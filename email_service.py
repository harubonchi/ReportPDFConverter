"""Email utilities for sending the merged PDF results via Gmail API."""

from __future__ import annotations

import base64
import os
import sys
from dataclasses import dataclass
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


# Gmail送信のみ（最小権限）
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]


def get_app_dir() -> Path:
    """
    Return the directory where the application is located.

    - When running as an exe (PyInstaller): directory of the executable
    - When running as a Python script: directory of this file
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


@dataclass(slots=True)
class EmailConfig:
    """Holds Gmail API configuration required to deliver emails."""

    sender: str
    display_name: str
    credentials_json: Path
    token_json: Path

    @property
    def is_configured(self) -> bool:
        return self.sender and self.credentials_json.exists()

    @classmethod
    def from_env(cls) -> "EmailConfig":
        app_dir = get_app_dir()
        return cls(
            sender=os.getenv(
                "EMAIL_SENDER", "roboken.report.tool@gmail.com"
            ),
            display_name=os.getenv(
                "EMAIL_DISPLAY_NAME", "ロボ研報告書作成ツール"
            ),
            credentials_json=Path(
                os.getenv(
                    "GMAIL_CREDENTIALS_JSON",
                    app_dir / "credentials.json",
                )
            ),
            token_json=Path(
                os.getenv(
                    "GMAIL_TOKEN_JSON",
                    app_dir / "token.json",
                )
            ),
        )


def _get_gmail_service(config: EmailConfig):
    creds: Credentials | None = None

    if config.token_json.exists():
        creds = Credentials.from_authorized_user_file(
            config.token_json, SCOPES
        )

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                config.credentials_json, SCOPES
            )
            creds = flow.run_local_server(port=0)

        config.token_json.write_text(
            creds.to_json(), encoding="utf-8"
        )

    return build("gmail", "v1", credentials=creds)


def send_email_with_attachment(
    *,
    config: EmailConfig,
    recipients: list[str] | str,
    cc_recipients: list[str] | None = None,
    subject: str,
    body: str,
    attachment_path: Path,
) -> None:
    """Send an email with ``attachment_path`` attached as a PDF document via Gmail API."""

    if not config.is_configured:
        raise RuntimeError("Email configuration is incomplete.")

    def _clean_addresses(value: list[str] | str) -> list[str]:
        if isinstance(value, str):
            items = [value]
        elif isinstance(value, list):
            items = value
        else:
            return []
        cleaned: list[str] = []
        seen: set[str] = set()
        for addr in items:
            normalized = str(addr or "").strip()
            if not normalized:
                continue
            key = normalized.lower()
            if key in seen:
                continue
            seen.add(key)
            cleaned.append(normalized)
        return cleaned

    to_addresses = _clean_addresses(recipients)
    cc_addresses = _clean_addresses(cc_recipients or [])

    if not to_addresses and cc_addresses:
        to_addresses = cc_addresses
        cc_addresses = []

    if not to_addresses:
        raise ValueError("At least one recipient is required.")

    service = _get_gmail_service(config)

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = formataddr((config.display_name, config.sender))
    message["To"] = ", ".join(to_addresses)
    if cc_addresses:
        message["Cc"] = ", ".join(cc_addresses)
    message.set_content(body)

    with attachment_path.open("rb") as attachment_file:
        message.add_attachment(
            attachment_file.read(),
            maintype="application",
            subtype="pdf",
            filename=attachment_path.name,
        )

    encoded_message = base64.urlsafe_b64encode(
        message.as_bytes()
    ).decode("utf-8")

    service.users().messages().send(
        userId="me",
        body={"raw": encoded_message},
    ).execute()
