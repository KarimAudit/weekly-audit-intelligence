"""Weekly Audit Intelligence Agent (AIA) - Main Execution Script."""

import datetime
import json
import logging
import os
import re
import smtplib
import sys
import urllib.request
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import List

from google import genai
from google.genai import types
import markdown
from docx import Document

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
        self.base_dir = Path(__file__).parent.parent
        self.config = self._load_json(self.base_dir / config_path)
        self.system_prompt = self._load_text(self.base_dir / prompt_path)
        self.sources = self._load_sources(self.base_dir / sources_path)

        api_key = os.environ.get("GEMINI_API_KEY")
        if not api_key:
            logger.error("GEMINI_API_KEY environment variable is missing.")
            raise ValueError("GEMINI_API_KEY environment variable required.")

        self.client = genai.Client(api_key=api_key)

    def _load_json(self, filepath: Path) -> dict:
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"Failed to load JSON configuration from {filepath}: {e}")
            raise

    def _load_text(self, filepath: Path) -> str:
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception as e:
            logger.error(f"Failed to load text file from {filepath}: {e}")
            raise

    def _load_sources(self, filepath: Path) -> List[str]:
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
        logger.info("Gathering source references and contextual metadata...")
        context_blocks = []
        
        for entry in self.sources:
            url_match = re.search(r"https?://[^\s]+", entry)
            target_url = url_match.group(0) if url_match else None

            if not target_url:
                context_blocks.append(f"Source Reference: {entry}")
                continue

            try:
                req = urllib.request.Request(
                    target_url, 
                    headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
                )
                with urllib.request.urlopen(req, timeout=5) as response:
                    status = response.status
                    context_blocks.append(f"Source Reference: {entry} (Status: {status})")
            except Exception as e:
                logger.warning(f"Unable to directly reach {target_url}: {e}")
                context_blocks.append(f"Source Reference: {entry}")

        aggregated_context = "\n".join(context_blocks)
        return f"Target Authoritative Reference Sources:\n{aggregated_context}"

    def generate_report(self, source_context: str) -> str:
        logger.info("Requesting report generation from Gemini API...")
        
        primary_model = self.config.get("model_name", "gemini-3.5-flash")
        candidate_models = [primary_model, "gemini-3.6-flash", "gemini-3.5-flash-lite"]
        candidate_models = list(dict.fromkeys(candidate_models))
        
        today_str = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d")
        
        user_prompt = (
            f"Date of Report: {today_str}\n\n"
            f"Please research and analyze the latest developments, whitepapers, standards, "
            f"and publications from the following target sources and domain areas:\n\n"
            f"{source_context}\n\n"
            f"CRITICAL INSTRUCTION: Do NOT include any hyperlinks, URLs, or web links anywhere in the generated text. "
            f"Present all information directly as professional editorial text. "
            f"Ensure strict adherence to language requirements (English Executive Summary, Arabic Main Sections)."
        )

        last_exception = None
        for model_name in candidate_models:
            try:
                logger.info(f"Attempting report generation with model: {model_name}")
                response = self.client.models.generate_content(
                    model=model_name,
                    contents=user_prompt,
                    config=types.GenerateContentConfig(
                        system_instruction=self.system_prompt,
                        temperature=self.config.get("temperature", 0.2),
                        top_p=self.config.get("top_p", 0.95),
                    ),
                )
                if response.text:
                    logger.info(f"Report successfully generated using model: {model_name}")
                    return response.text
            except Exception as e:
                logger.warning(f"Failed generation attempt with {model_name}: {e}")
                last_exception = e

        if last_exception:
            raise last_exception
        raise RuntimeError("Failed to generate report with available Gemini models.")

    def strip_all_links(self, text: str) -> str:
        """Completely strip URLs and Markdown hyperlink structures from text."""
        # Convert Markdown links [Text](URL) to just "Text"
        text = re.sub(r'\[([^\]]+)\]\([^)]+\)', r'\1', text)
        # Strip plain URLs starting with http:// or https://
        text = re.sub(r'https?://[^\s]+', '', text)
        return text

    def save_artifacts(self, report_md: str) -> Path:
        """Save clean report artifacts."""
        reports_dir = self.base_dir / "reports"
        reports_dir.mkdir(parents=True, exist_ok=True)
        date_suffix = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d")

        clean_text = self.strip_all_links(report_md)

        # Markdown
        md_path = reports_dir / f"audit_intelligence_{date_suffix}.md"
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(clean_text)

        # Word Document (.docx)
        docx_path = reports_dir / f"audit_intelligence_{date_suffix}.docx"
        doc = Document()
        doc.add_heading("Weekly Audit Intelligence Report", 0)
        for paragraph in clean_text.split("\n\n"):
            if paragraph.strip():
                doc.add_paragraph(paragraph.strip())
        doc.save(str(docx_path))

        logger.info(f"Saved weekly report artifacts to {reports_dir}")
        return docx_path

    def build_executive_newsletter_html(self, report_md: str) -> str:
        """Construct a McKinsey/EY styled HTML email layout with zero external links."""
        clean_text = self.strip_all_links(report_md)
        
        # Convert markdown text to valid HTML elements
        html_content = markdown.markdown(clean_text, extensions=['tables', 'fenced_code'])
        
        date_str = datetime.datetime.now(datetime.timezone.utc).strftime("%B %d, %Y")

        return f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                body {{
                    font-family: 'Georgia', 'Times New Roman', serif;
                    background-color: #f8f9fa;
                    color: #222222;
                    margin: 0;
                    padding: 0;
                    -webkit-font-smoothing: antialiased;
                }}
                .email-wrapper {{
                    max-width: 680px;
                    margin: 30px auto;
                    background: #ffffff;
                    border: 1px solid #e2e8f0;
                    border-radius: 4px;
                    overflow: hidden;
                    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
                }}
                .email-header {{
                    background-color: #00205b; /* Executive Navy Blue */
                    color: #ffffff;
                    padding: 35px 40px;
                    border-bottom: 4px solid #c59b27; /* Gold accent line */
                }}
                .email-header h1 {{
                    font-family: 'Georgia', serif;
                    font-size: 26px;
                    font-weight: normal;
                    margin: 0 0 8px 0;
                    letter-spacing: 0.5px;
                }}
                .email-header .sub-header {{
                    font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
                    font-size: 13px;
                    text-transform: uppercase;
                    letter-spacing: 1.5px;
                    color: #cbd5e1;
                    margin: 0;
                }}
                .email-body {{
                    padding: 40px;
                    font-size: 16px;
                    line-height: 1.8;
                    color: #2d3748;
                }}
                .email-body h1, .email-body h2, .email-body h3 {{
                    font-family: 'Georgia', serif;
                    color: #00205b;
                    margin-top: 30px;
                    margin-bottom: 12px;
                    font-weight: 600;
                }}
                .email-body h1 {{ font-size: 22px; border-bottom: 1px solid #e2e8f0; padding-bottom: 8px; }}
                .email-body h2 {{ font-size: 19px; }}
                .email-body h3 {{ font-size: 17px; }}
                .email-body p {{
                    margin-bottom: 18px;
                }}
                .email-body ul, .email-body ol {{
                    margin-bottom: 20px;
                    padding-left: 24px;
                }}
                .email-body li {{
                    margin-bottom: 8px;
                }}
                .email-body blockquote {{
                    margin: 24px 0;
                    padding: 16px 20px;
                    background-color: #f7fafc;
                    border-left: 4px solid #00205b;
                    font-style: italic;
                    color: #4a5568;
                }}
                .email-footer {{
                    background-color: #f1f5f9;
                    padding: 24px 40px;
                    border-top: 1px solid #e2e8f0;
                    font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
                    font-size: 12px;
                    color: #64748b;
                    text-align: center;
                }}
            </style>
        </head>
        <body>
            <div class="email-wrapper">
                <div class="email-header">
                    <p class="sub-header">Executive Briefing &bull; {date_str}</p>
                    <h1>Audit Intelligence Weekly</h1>
                </div>
                <div class="email-body" dir="auto">
                    {html_content}
                </div>
                <div class="email-footer">
                    Confidential &bull; Generated for Internal Audit Leadership
                </div>
            </div>
        </body>
        </html>
        """

    def send_email(self, report_md: str, docx_path: Path) -> None:
        """Send formatted executive email with word document attachment and zero links."""
        smtp_server = os.environ.get("SMTP_SERVER", "smtp.gmail.com")
        smtp_port = int(os.environ.get("SMTP_PORT", "587"))
        smtp_user = os.environ.get("SMTP_USER")
        smtp_password = os.environ.get("SMTP_PASSWORD")

        recipients = self.config.get("email_recipients", [
            "andro.ios.developer@gmail.com",
            "karimbenkarim@yahoo.fr"
        ])

        if not smtp_user or not smtp_password:
            logger.warning("SMTP credentials missing. Skipping email delivery.")
            return

        date_str = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d")
        subject = f"Audit Intelligence Executive Briefing - {date_str}"

        msg = MIMEMultipart()
        msg["Subject"] = subject
        msg["From"] = smtp_user
        msg["To"] = ", ".join(recipients)

        # HTML Newsletter Body
        html_body = self.build_executive_newsletter_html(report_md)
        msg.attach(MIMEText(html_body, "html", "utf-8"))

        # Word Document Attachment
        if docx_path.exists():
            with open(docx_path, "rb") as f:
                part = MIMEApplication(f.read(), Name=docx_path.name)
            part['Content-Disposition'] = f'attachment; filename="{docx_path.name}"'
            msg.attach(part)

        try:
            logger.info(f"Dispatching executive newsletter to: {', '.join(recipients)}...")
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(smtp_user, smtp_password)
                server.sendmail(smtp_user, recipients, msg.as_string())
            logger.info("Executive email successfully delivered.")
        except Exception as e:
            logger.error(f"Failed to send email via SMTP: {e}")

    def run(self) -> None:
        logger.info("Starting Audit Intelligence Agent execution...")
        context = self.gather_source_context()
        report_md = self.generate_report(context)
        docx_path = self.save_artifacts(report_md)
        self.send_email(report_md, docx_path)
        logger.info("Audit Intelligence Agent execution completed successfully.")


if __name__ == "__main__":
    agent = AuditIntelligenceAgent()
    agent.run()
