"""Weekly Audit Intelligence Agent (AIA) - Main Execution Script.

Fetches intelligence sources, calls the Google Gemini API using the modern
`google-genai` SDK, saves the structured report in Markdown format, and dispatches
it via email to configured stakeholders.
"""

import datetime
import json
import logging
import os
import smtplib
import sys
import urllib.request
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Dict, List, Optional

from google import genai
from google.genai import types

# Configure structured logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s - %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger("AIA-Main")


class AuditIntelligenceAgent:
    """Orchestrates source collection, AI report generation, artifact saving, and email dispatch."""

    def __init__(
        self,
        config_path: str = "config/settings.json",
        prompt_path: str = "prompts/system.md",
        sources_path: str = "sources/sources.txt",
    ) -> None:
        """Initialize the agent, loading configurations and setting up Gemini client."""
        self.base_dir = Path(__file__).parent.parent
        self.config = self._load_json(self.base_dir / config_path)
        self.system_prompt = self._load_text(self.base_dir / prompt_path)
        self.sources = self._load_sources(self.base_dir / sources_path)

        api_key = os.environ.get("GEMINI_API_KEY")
        if not api_key:
            logger.error("GEMINI_API_KEY environment variable is missing.")
            raise ValueError("GEMINI_API_KEY environment variable required.")

        # Initialize official modern Google GenAI Client
        self.client = genai.Client(api_key=api_key)

    def _load_json(self, filepath: Path) -> dict:
        """Load JSON configuration file safely."""
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"Failed to load JSON configuration from {filepath}: {e}")
            raise

    def _load_text(self, filepath: Path) -> str:
        """Load plain text or Markdown file safely."""
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception as e:
            logger.error(f"Failed to load text file from {filepath}: {e}")
            raise

    def _load_sources(self, filepath: Path) -> List[str]:
        """Load source URLs from a line-separated text file."""
        if not filepath.exists():
            logger.warning(f"Sources file not found at {filepath}. Using defaults from config.")
            return self.config.get("default_sources", [])
        
        sources = []
        with open(filepath, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#"):
                    sources.append(line)
        return sources

def gather_source_context(self) -> str:
        """Fetch content metadata or content summaries from configured authoritative sources."""
        logger.info("Gathering source references and contextual metadata...")
        context_blocks = []
        
        for url in self.sources:
            try:
                req = urllib.request.Request(
                    url, 
                    headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
                )
                # Reduced socket timeout to prevent long network hangs
                with urllib.request.urlopen(req, timeout=5) as response:
                    status = response.status
                    context_blocks.append(f"Source URL: {url} (Status: {status})")
            except Exception as e:
                logger.warning(f"Unable to directly reach {url}: {e}. Including URL as text reference.")
                context_blocks.append(f"Source URL: {url} (Reference only)")

        aggregated_context = "\n".join(context_blocks)
        return f"Target Authoritative Reference Sources:\n{aggregated_context}"
    def generate_report(self, source_context: str) -> str:
        """Call Gemini API using official google-genai SDK to produce the intelligence report."""
        logger.info("Requesting report generation from Gemini API...")
        
        model_name = self.config.get("model_name", "gemini-2.5-flash")
        
        today_str = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d")
        
        user_prompt = (
            f"Date of Report: {today_str}\n\n"
            f"Please research and analyze the latest developments, whitepapers, standards, "
            f"and publications from the following target sources and domain areas:\n\n"
            f"{source_context}\n\n"
            f"Ensure strict adherence to the output format, language requirements (English Executive Summary, "
            f"Arabic Main Sections), and professional auditing tone described in your system instructions."
        )

        try:
            response = self.client.models.generate_content(
                model=model_name,
                contents=user_prompt,
                config=types.GenerateContentConfig(
                    system_instruction=self.system_prompt,
                    temperature=self.config.get("temperature", 0.2),
                    top_p=self.config.get("top_p", 0.95),
                ),
            )
            
            if not response.text:
                raise ValueError("Gemini API returned an empty response.")
                
            logger.info("Report generated successfully from Gemini API.")
            return response.text

        except Exception as e:
            logger.error(f"Error during Gemini API generation: {e}")
            raise

    def save_report(self, report_md: str) -> Path:
        """Save generated Markdown report to reports/ directory with standard date naming."""
        reports_dir = self.base_dir / "reports"
        reports_dir.mkdir(parents=True, exist_ok=True)

        date_suffix = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d")
        file_path = reports_dir / f"audit_intelligence_{date_suffix}.md"

        with open(file_path, "w", encoding="utf-8") as f:
            f.write(report_md)

        logger.info(f"Saved weekly report artifact to {file_path}")
        return file_path

    def send_email(self, report_md: str) -> None:
        """Send email with report content to all configured recipients via SMTP."""
        smtp_server = os.environ.get("SMTP_SERVER", "smtp.gmail.com")
        smtp_port = int(os.environ.get("SMTP_PORT", "587"))
        smtp_user = os.environ.get("SMTP_USER")
        smtp_password = os.environ.get("SMTP_PASSWORD")

        recipients = self.config.get("email_recipients", [
            "andro.ios.developer@gmail.com",
            "karimbenkarim@yahoo.fr"
        ])

        if not smtp_user or not smtp_password:
            logger.warning(
                "SMTP credentials (SMTP_USER / SMTP_PASSWORD) not configured. "
                "Skipping automated email delivery. Report remains safely archived in repository."
            )
            return

        date_str = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d")
        subject = f"Weekly Audit Intelligence Report - {date_str}"

        # Construct email with HTML rendering for RTL text support
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"] = smtp_user
        msg["To"] = ", ".join(recipients)

        # Plain text version
        msg.attach(MIMEText(report_md, "plain", "utf-8"))

        # Basic HTML wrapper supporting Arabic RTL reading
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
                pre {{ background-color: #f4f4f4; padding: 15px; border-radius: 5px; white-space: pre-wrap; }}
            </style>
        </head>
        <body>
            <div dir="auto">
                <pre>{report_md}</pre>
            </div>
        </body>
        </html>
        """
        msg.attach(MIMEText(html_content, "html", "utf-8"))

        try:
            logger.info(f"Connecting to SMTP server {smtp_server}:{smtp_port}...")
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(smtp_user, smtp_password)
                server.sendmail(smtp_user, recipients, msg.as_string())
            logger.info(f"Email successfully dispatched to: {', '.join(recipients)}")
        except Exception as e:
            logger.error(f"Failed to send email via SMTP: {e}")
            # Non-fatal error; report is already stored in repository artifacts

    def run(self) -> None:
        """Execute the end-to-end pipeline."""
        logger.info("Starting Audit Intelligence Agent execution...")
        context = self.gather_source_context()
        report_md = self.generate_report(context)
        self.save_report(report_md)
        self.send_email(report_md)
        logger.info("Audit Intelligence Agent execution completed successfully.")


if __name__ == "__main__":
    agent = AuditIntelligenceAgent()
    agent.run()
