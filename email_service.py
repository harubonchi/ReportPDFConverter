"""Email utilities for sending the merged PDF results via Gmail API."""

from __future__ import annotations

import base64
import os
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
        return cls(
            sender=os.getenv("EMAIL_SENDER", ""),
            display_name=os.getenv(
                "EMAIL_DISPLAY_NAME", "ロボ研報告書作成ツール"
            ),
            credentials_json=Path(
                os.getenv("GMAIL_CREDENTIALS_JSON", "credentials.json")
            ),
            token_json=Path(
                os.getenv("GMAIL_TOKEN_JSON", "token.json")
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
    recipient: str,
    subject: str,
    body: str,
    attachment_path: Path,
) -> None:
    """Send an email with ``attachment_path`` attached as a PDF document via Gmail API."""

    if not config.is_configured:
        raise RuntimeError("Email configuration is incomplete.")

    service = _get_gmail_service(config)

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = formataddr((config.display_name, config.sender))
    message["To"] = recipient
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
