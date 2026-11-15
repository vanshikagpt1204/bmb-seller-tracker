# BuyMeABook Seller Tracker

A Python-based seller monitoring and reporting application.

## Features
- eBay Singapore Books listing integration
- Seller and listing event tracking
- SQLite database
- CSV report generation
- SMTP email reporting
- Daily scheduled execution
- Docker support

## Technology Stack
- Python 3.10
- SQLite
- Docker
- eBay Browse API
- Requests
- OpenPyXL

## Local Setup

1. Install dependencies:

   pip install -r requirements.txt

2. Create environment file:

   Copy-Item .env.example .env

3. Add authorized API and SMTP credentials to .env.

4. Run:

   python -m app.main

## Docker Setup

Build:

   docker build -t bmb-seller-tracker:latest .

Run:

   docker run -d --name bmb-seller-tracker --env-file .env -v bmb-seller-data:/app/data bmb-seller-tracker:latest

## Current Limitations

- Live eBay API access requires approved developer credentials.
- Email delivery requires SMTP credentials.
- Public eBay Browse API does not provide private seller email addresses, phone numbers or registration dates.
- Live API and email delivery have not yet been verified.

## Security

Never upload .env files, passwords, API credentials or private seller data.
