from __future__ import annotations

import os
import smtplib
from dataclasses import dataclass
from email.message import EmailMessage
from pathlib import Path


@dataclass
class EmailConfig:
    sender: str
    username: str
    password: str
    smtp_server: str
    smtp_port: int
    use_tls: bool = True

    @property
    def is_configured(self) -> bool:
        return all([self.sender, self.username, self.password, self.smtp_server, self.smtp_port])

    @classmethod
    def from_env(cls) -> "EmailConfig":
        sender = os.getenv("EMAIL_SENDER", "")
        username = os.getenv("EMAIL_USERNAME", sender)
        password = os.getenv("EMAIL_PASSWORD", "")
        smtp_server = os.getenv("EMAIL_SMTP_SERVER", "")

        try:
            smtp_port = int(os.getenv("EMAIL_SMTP_PORT", ""))
        except (TypeError, ValueError):
            smtp_port = 0

        use_tls_value = os.getenv("EMAIL_USE_TLS", "true")
        use_tls = str(use_tls_value).strip().lower() in {"1", "true", "yes", "on"}
        return cls(
            sender=sender,
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
    if not config.is_configured:
        raise RuntimeError("Email configuration is incomplete.")

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = config.sender
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