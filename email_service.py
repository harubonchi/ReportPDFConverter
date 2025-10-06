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
        sender = 'harubonchi@gmail.com'
        username = 'harubonchi@gmail.com'
        password = 'zuaa oqrg star rtcx'
        smtp_server = 'smtp.gmail.com'
        smtp_port = 587
        use_tls = 'true'
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