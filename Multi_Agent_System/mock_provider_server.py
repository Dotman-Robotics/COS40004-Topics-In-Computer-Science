"""
mock_provider_server.py

Serves the Brightline Catering Co. mock website on localhost:8080.
Run this before testing the search agent scraper integration.

Usage:
    python mock_provider_server.py

The site will be available at http://localhost:8080
Emails on the site: hello@brightlinecatering.com.au, bookings@brightlinecatering.com.au

For the test, point the search agent at http://localhost:8080 directly,
or use the search agent normally and manually add localhost:8080 to the
results list if DDG doesn't return it.
"""

import os
from flask import Flask, send_from_directory

app = Flask(__name__)
SITE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mock_provider")


@app.route("/")
@app.route("/index.html")
def home():
    return send_from_directory(SITE_DIR, "index.html")


@app.route("/contact")
@app.route("/contact.html")
def contact():
    return send_from_directory(SITE_DIR, "contact.html")


@app.route("/services")
@app.route("/services.html")
def services():
    return send_from_directory(SITE_DIR, "services.html")


if __name__ == "__main__":
    print("Mock provider site running at http://localhost:8080")
    print("Emails: hello@brightlinecatering.com.au, bookings@brightlinecatering.com.au")
    app.run(port=8080, debug=False)