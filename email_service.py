"""Email utilities for sending the merged PDF results."""

from __future__ import annotations

import os
import smtplib
from dataclasses import dataclass
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path


def _env_bool(name: str, default: bool) -> bool:
    """Return a boolean value parsed from an environment variable."""

    raw = os.getenv(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on"}


@dataclass(slots=True)
class EmailConfig:
    """Holds SMTP configuration required to deliver emails."""

    sender: str
    display_name: str
    username: str
    password: str
    smtp_server: str
    smtp_port: int
    use_tls: bool = True

    @property
    def is_configured(self) -> bool:
        """Return ``True`` when the configuration is sufficient for delivery."""

        return all(
            [
                self.sender,
                self.username,
                self.password,
                self.smtp_server,
                self.smtp_port,
            ]
        )

    @classmethod
    def from_env(cls) -> "EmailConfig":
        """Build a configuration instance from environment variables."""

        sender = os.getenv("EMAIL_SENDER", "")
        display_name = os.getenv("EMAIL_DISPLAY_NAME", "ロボ研報告書作成ツール")
        username = os.getenv("EMAIL_USERNAME", sender)
        password = os.getenv("EMAIL_PASSWORD", "")
        smtp_server = os.getenv("EMAIL_SMTP_SERVER", "")

        try:
            smtp_port = int(os.getenv("EMAIL_SMTP_PORT", ""))
        except (TypeError, ValueError):
            smtp_port = 0

        use_tls = _env_bool("EMAIL_USE_TLS", default=True)
        return cls(
            sender=sender,
            display_name=display_name,
            username=username,
            password=password,
            smtp_server=smtp_server,
            smtp_port=smtp_port,
            use_tls=use_tls,
        )


def send_email_with_attachment(
    *,
    config: EmailConfig,
    recipient: str,
    subject: str,
    body: str,
    attachment_path: Path,
) -> None:
    """Send an email with ``attachment_path`` attached as a PDF document."""

    if not config.is_configured:
        raise RuntimeError("Email configuration is incomplete.")

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = formataddr((config.display_name, config.sender))
    message["To"] = recipient
    message.set_content(body)

    with attachment_path.open("rb") as attachment_file:
        file_data = attachment_file.read()
    message.add_attachment(
        file_data,
        maintype="application",
        subtype="pdf",
        filename=attachment_path.name,
    )

    with smtplib.SMTP(config.smtp_server, config.smtp_port) as server:
        if config.use_tls:
            server.starttls()
        server.login(config.username, config.password)
        server.send_message(message)
