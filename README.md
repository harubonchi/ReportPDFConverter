# Gain Report Emailer Web Interface

This project provides a Flask-based web application that converts Word documents from an
uploaded ZIP archive into PDF, merges them, and emails the final report to a chosen
recipient.

## Features

- Upload ZIP archives that contain `.doc` or `.docx` files.
- Choose the order in which documents appear in the merged PDF through the web UI.
- Background processing of extraction, conversion, merging, and emailing.
- Remembers the most recently used order and suggests it for future uploads.

## Prerequisites

- Python 3.10+
- LibreOffice (optional, required only for legacy `.doc` conversion).
- SMTP credentials for sending the merged PDF via email.

## Environment variables

Copy `.env.example` to `.env` (or set the environment variables directly) and provide
your email credentials:

```
EMAIL_SENDER=example@example.com
EMAIL_USERNAME=example@example.com
EMAIL_PASSWORD=your-password
EMAIL_SMTP_SERVER=smtp.example.com
EMAIL_SMTP_PORT=587
EMAIL_USE_TLS=true
```

## Running the application

```bash
pip install -r requirements.txt
python app.py
```

The application runs on `http://localhost:5000` by default.