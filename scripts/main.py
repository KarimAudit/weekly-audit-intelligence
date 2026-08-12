"""Weekly Audit Intelligence Agent (AIA) - Main Execution Script (World-Class Arabic Magazine Layout)."""

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
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s - %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger("AIA-Main")


class AuditIntelligenceAgent:
    """Orchestrates source collection, AI report generation (Arabic focus), magazine-style artifact saving, and email dispatch."""

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

        primary_model = self.config.get("model_name", "gemini-3.6-flash")
        candidate_models = [primary_model, "gemini-3.6-flash", "gemini-3.6-pro", "gemini-3.5-flash"]
        candidate_models = list(dict.fromkeys(candidate_models))

        today_str = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d")

        user_prompt = (
            f"Date of Newsletter: {today_str}\n\n"
            f"Target Authoritative Reference Sources & Topics:\n"
            f"{source_context}\n\n"
            f"Instruction: Generate a world-class, McKinsey/EY-style executive newsletter in formal business Arabic (اللغة العربية الفصحى).\n"
            f"Formatting Requirements:\n"
            f"- Incorporate high-impact section headers with contextual Arabic business terminology.\n"
            f"- Use professional emojis/icons strategically (e.g., 📊, 🛡️, 💡, ⚖️, 🎯, 🔍) to improve visual structure without cluttering.\n"
            f"- Include actionable Executive Summaries, Key Risk Indicators, and Governance Insights.\n"
            f"- Keep technical English terms or abbreviations (e.g., IIA Standards, ISO 27001, ESG, COSO) in parentheses where necessary.\n"
            f"- Include clickable, functional markdown links [المصدر](URL) in the references section."
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
                        temperature=self.config.get("temperature", 0.3),
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

    def _set_cell_background(self, cell, hex_color: str):
        """Helper to set background color for docx table cells."""
        tc_pr = cell._element.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), hex_color)
        tc_pr.append(shd)

    def save_artifacts(self, report_md: str) -> Path:
        """Save report artifacts as Markdown and Magazine-Style formatted DOCX."""
        reports_dir = self.base_dir / "reports"
        reports_dir.mkdir(parents=True, exist_ok=True)
        date_suffix = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d")

        # Save Markdown
        md_path = reports_dir / f"audit_intelligence_{date_suffix}.md"
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(report_md)

        # Build World-Class Magazine DOCX
        docx_path = reports_dir / f"audit_intelligence_{date_suffix}.docx"
        doc = Document()

        # Set RTL for Document
        section = doc.sections[0]
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

        # 1. Magazine Cover Header Banner (Table-based styling)
        banner_table = doc.add_table(rows=1, cols=1)
        banner_table.autofit = False
        banner_table.columns[0].width = Inches(6.8)
        cell = banner_table.cell(0, 0)
        self._set_cell_background(cell, "0F172A")  # Slate Navy

        cell_p = cell.paragraphs[0]
        cell_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        cell_p.paragraph_format.space_before = Pt(14)
        cell_p.paragraph_format.space_after = Pt(4)

        run_sub = cell_p.add_run(f"موجز تنفيذي رفيع المستوى  •  {date_suffix}\n")
        run_sub.font.name = "Arial"
        run_sub.font.size = Pt(10)
        run_sub.font.color.rgb = RGBColor(148, 163, 184)

        run_title = cell_p.add_run("النشرة الذكية للتدقيق والحوكمة\nAudit & Governance Intelligence Brief")
        run_title.font.name = "Arial"
        run_title.font.size = Pt(20)
        run_title.font.bold = True
        run_title.font.color.rgb = RGBColor(255, 255, 255)

        doc.add_paragraph().paragraph_format.space_after = Pt(12)

        # Parse paragraphs and apply high-end styles
        lines = report_md.split("\n")
        for line in lines:
            line_str = line.strip()
            if not line_str:
                continue

            if line_str.startswith("# "):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.paragraph_format.space_before = Pt(18)
                p.paragraph_format.space_after = Pt(6)
                run = p.add_run(line_str.replace("# ", ""))
                run.font.name = "Arial"
                run.font.size = Pt(18)
                run.font.bold = True
                run.font.color.rgb = RGBColor(15, 23, 42)

            elif line_str.startswith("## "):
                # Ey-style Banner Header
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.paragraph_format.space_before = Pt(14)
                p.paragraph_format.space_after = Pt(4)
                run = p.add_run("  " + line_str.replace("## ", "") + "  ")
                run.font.name = "Arial"
                run.font.size = Pt(14)
                run.font.bold = True
                run.font.color.rgb = RGBColor(37, 99, 235)

            elif line_str.startswith("### "):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.paragraph_format.space_before = Pt(10)
                p.paragraph_format.space_after = Pt(2)
                run = p.add_run(line_str.replace("### ", ""))
                run.font.name = "Arial"
                run.font.size = Pt(12)
                run.font.bold = True
                run.font.color.rgb = RGBColor(15, 23, 42)

            elif line_str.startswith("> "):
                # Executive Highlight Box
                box_table = doc.add_table(rows=1, cols=1)
                box_cell = box_table.cell(0, 0)
                self._set_cell_background(box_cell, "F1F5F9")
                bp = box_cell.paragraphs[0]
                bp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                brun = bp.add_run(line_str.replace("> ", ""))
                brun.font.name = "Arial"
                brun.font.size = Pt(10.5)
                brun.font.italic = True
                brun.font.color.rgb = RGBColor(51, 65, 85)

            else:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.paragraph_format.space_after = Pt(6)
                run = p.add_run(line_str)
                run.font.name = "Arial"
                run.font.size = Pt(10.5)
                run.font.color.rgb = RGBColor(30, 41, 59)

        doc.save(str(docx_path))
        logger.info(f"Saved weekly newsletter artifacts to {reports_dir}")
        return docx_path

    def build_executive_newsletter_html(self, report_md: str) -> str:
        """Construct an EY/McKinsey-caliber HTML email layout with styled headers and RTL support."""
        html_content = markdown.markdown(report_md, extensions=['tables', 'fenced_code'])
        date_str = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d")

        return f"""
        <!DOCTYPE html>
        <html lang="ar" dir="rtl">
        <head>
            <meta charset="utf-8">
            <style>
                @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700&display=swap');

                body {{
                    font-family: 'Tajawal', 'Segoe UI', Arial, sans-serif;
                    background-color: #0f172a;
                    color: #1e293b;
                    margin: 0;
                    padding: 20px 0;
                    direction: rtl;
                    text-align: right;
                }}
                .email-wrapper {{
                    max-width: 720px;
                    margin: 0 auto;
                    background: #ffffff;
                    border-radius: 8px;
                    overflow: hidden;
                    box-shadow: 0 10px 25px rgba(0, 0, 0, 0.2);
                }}
                /* McKinsey / EY Header Banner */
                .email-header {{
                    background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
                    color: #ffffff;
                    padding: 40px;
                    border-bottom: 4px solid #059669; /* Emerald Accent */
                }}
                .email-header .tagline {{
                    display: inline-block;
                    background-color: rgba(255, 255, 255, 0.1);
                    color: #38bdf8;
                    font-size: 12px;
                    font-weight: 700;
                    padding: 4px 12px;
                    border-radius: 20px;
                    text-transform: uppercase;
                    margin-bottom: 12px;
                }}
                .email-header h1 {{
                    font-size: 26px;
                    font-weight: 700;
                    margin: 0 0 8px 0;
                    line-height: 1.3;
                }}
                .email-header .meta {{
                    font-size: 13px;
                    color: #94a3b8;
                    margin: 0;
                }}
                /* Main Body Styling */
                .email-body {{
                    padding: 40px;
                    font-size: 15.5px;
                    line-height: 1.9;
                    color: #334155;
                }}
                /* Section Headings Styling */
                .email-body h1 {{
                    font-size: 21px;
                    color: #0f172a;
                    background: #f8fafc;
                    padding: 10px 16px;
                    border-right: 5px solid #2563eb;
                    border-radius: 0 6px 6px 0;
                    margin-top: 32px;
                    margin-bottom: 16px;
                    font-weight: 700;
                }}
                .email-body h2 {{
                    font-size: 18px;
                    color: #0f172a;
                    background: #f1f5f9;
                    padding: 8px 14px;
                    border-right: 4px solid #059669;
                    border-radius: 0 4px 4px 0;
                    margin-top: 24px;
                    margin-bottom: 12px;
                    font-weight: 700;
                }}
                .email-body h3 {{
                    font-size: 16px;
                    color: #1e293b;
                    margin-top: 18px;
                    margin-bottom: 8px;
                    font-weight: 700;
                }}
                /* Highlights & Quotes */
                .email-body blockquote {{
                    margin: 20px 0;
                    padding: 16px 20px;
                    background-color: #f0fdf4;
                    border-right: 4px solid #16a34a;
                    border-radius: 4px;
                    color: #166534;
                    font-weight: 500;
                }}
                .email-body a {{
                    color: #2563eb;
                    text-decoration: underline;
                    font-weight: 600;
                }}
                /* Corporate Tables */
                .email-body table {{
                    width: 100%;
                    border-collapse: collapse;
                    margin: 24px 0;
                    font-size: 14.5px;
                }}
                .email-body th {{
                    background-color: #0f172a;
                    color: #ffffff;
                    font-weight: 600;
                    padding: 12px;
                    text-align: right;
                }}
                .email-body td {{
                    border-bottom: 1px solid #e2e8f0;
                    padding: 10px 12px;
                    text-align: right;
                }}
                .email-body tr:nth-child(even) {{
                    background-color: #f8fafc;
                }}
                /* Executive Footer */
                .email-footer {{
                    background-color: #f8fafc;
                    padding: 24px 40px;
                    border-top: 1px solid #e2e8f0;
                    font-size: 12px;
                    color: #64748b;
                    text-align: center;
                }}
            </style>
        </head>
        <body>
            <div class="email-wrapper">
                <div class="email-header">
                    <span class="tagline">موجز الاستشارات التنفيذية</span>
                    <h1>النشرة الذكية للتدقيق والحوكمة 🏛️</h1>
                    <p class="meta">اصدار قيادي أسبوعي  &bull;  {date_str}</p>
                </div>
                <div class="email-body">
                    {html_content}
                </div>
                <div class="email-footer">
                    سري للغاية &bull; تم التطوير خصيصاً لقيادات التدقيق الداخلي والحوكمة المؤسسية
                </div>
            </div>
        </body>
        </html>
        """

    def send_email(self, report_md: str, docx_path: Path) -> None:
        """Send formatted newsletter email with docx attachment and active links."""
        smtp_server = os.environ.get("SMTP_SERVER", "smtp.gmail.com")
        smtp_port = int(os.environ.get("SMTP_PORT", "587"))
        smtp_user = os.environ.get("SMTP_USER")
        smtp_password = os.environ.get("SMTP_PASSWORD")

        recipients = self.config.get("email_recipients", [])

        if not recipients:
            logger.warning("No email recipients configured in settings.json. Skipping dispatch.")
            return

        if not smtp_user or not smtp_password:
            logger.warning("SMTP credentials missing. Skipping email delivery.")
            return

        date_str = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d")
        subject = f"🏛️ موجز التدقيق والحوكمة الأسبوعي - {date_str}"

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
            logger.info(f"Dispatching executive newsletter to {len(recipients)} recipients...")
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
