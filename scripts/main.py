import os
import json
import logging
import smtplib
import subprocess
import platform
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime
from typing import List, Dict, Any

import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

# Setup logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Corporate Palette (Customized)
NAVY_PRIMARY = "B734F7"     # Dark Header Fill / Primary Accent
ACCENT_BLUE  = "B12763"     # Secondary Highlight / Borders
GOLD_ACCENT  = "D4AF37"     # Premium Highlight
BG_LIGHT     = "F8FAFC"     # Callout & Dashboard Fill
TEXT_DARK    = "1E293B"     # Body Text Color
TEXT_WHITE   = "FFFFFF"     # High Contrast Header Text

# Risk Colors
RISK_COLORS = {
    "حرج": "991B1B",      # Red Dark
    "عالٍ": "C2410C",     # Orange Dark
    "متوسط": "854D0E",   # Yellow/Brown Dark
    "منخفض": "166534"    # Green Dark
}

def set_cell_background(cell, color_hex: str):
    """Sets the background color of a Word table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Sets padding inside table cells."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for margin_name, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{margin_name}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_executive_dashboard(doc, blocks: List[Dict[str, Any]]):
    """Generates an Executive Dashboard Summary Box at the top of the report."""
    tbl = doc.add_table(rows=2, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    total_items = len(blocks)
    critical_high = sum(1 for b in blocks if b.get("risk_level") in ["حرج", "عالٍ"])
    categories = list(set(b.get("category", "عام") for b in blocks))
    
    headers = ["إجمالي التحديثات", "تحديثات عالية الخطورة", "المجالات المغطاة"]
    for i, h in enumerate(headers):
        cell = tbl.cell(0, i)
        set_cell_background(cell, NAVY_PRIMARY)
        set_cell_margins(cell, top=120, bottom=120)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.font.name = "Traditional Arabic"
        run.font.bold = True
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    values = [str(total_items), str(critical_high), ", ".join(categories[:2])]
    for i, v in enumerate(values):
        cell = tbl.cell(1, i)
        set_cell_background(cell, BG_LIGHT)
        set_cell_margins(cell, top=140, bottom=140)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(v)
        run.font.name = "Traditional Arabic"
        run.font.bold = True
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(0xB7, 0x34, 0xF7) if i != 1 else RGBColor(0x99, 0x1B, 0x1B)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

def create_section_header(doc, title: str, risk: str = "متوسط", category: str = "حوكمة"):
    """Creates a high-contrast dark section title bar with integrated risk badge."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, NAVY_PRIMARY)
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)
    
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    run_t = p.add_run(f"{title}  ")
    run_t.font.name = "Traditional Arabic"
    run_t.font.bold = True
    run_t.font.size = Pt(15)
    run_t.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    
    run_tag = p.add_run(f" [{category} | مستوى الخطورة: {risk}]")
    run_tag.font.name = "Traditional Arabic"
    run_tag.font.size = Pt(11)
    run_tag.font.bold = True
    run_tag.font.color.rgb = RGBColor(0xD4, 0xAF, 0x37)

def add_action_box(doc, why_it_matters: str, recommended_actions: str):
    """Creates a structured callout box for business impact and audit actions."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, BG_LIGHT)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'  <w:top w:val="none"/>'
        f'  <w:left w:val="none"/>'
        f'  <w:bottom w:val="none"/>'
        f'  <w:right w:val="single" w:sz="24" w:space="0" w:color="{ACCENT_BLUE}"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)
    set_cell_margins(cell, top=120, bottom=120, left=180, right=180)
    
    p = cell.paragraphs[0]
    p.paragraph_format.line_spacing = 1.25
    p.paragraph_format.space_after = Pt(4)
    
    r1 = p.add_run("🎯 الأثر المباشر على المؤسسة (Why It Matters):\n")
    r1.font.name = "Traditional Arabic"
    r1.font.bold = True
    r1.font.size = Pt(12)
    r1.font.color.rgb = RGBColor(0xB1, 0x27, 0x63)
    
    r2 = p.add_run(f"{why_it_matters}\n\n")
    r2.font.name = "Traditional Arabic"
    r2.font.size = Pt(11.5)
    r2.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)
    
    r3 = p.add_run("🛠️ إجراءات التدقيق والحوكمة الموصى بها:\n")
    r3.font.name = "Traditional Arabic"
    r3.font.bold = True
    r3.font.size = Pt(12)
    r3.font.color.rgb = RGBColor(0xB1, 0x27, 0x63)
    
    r4 = p.add_run(f"{recommended_actions}")
    r4.font.name = "Traditional Arabic"
    r4.font.size = Pt(11.5)
    r4.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)

def build_word_document(content_blocks: List[Dict[str, Any]], reports_dir: str = "reports") -> str:
    """Builds a corporate Word document inside the designated reports directory."""
    os.makedirs(reports_dir, exist_ok=True)
    
    filename = os.path.join(reports_dir, f"Governance_Weekly_Brief_{datetime.now().strftime('%Y-%m-%d')}.docx")
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
        footer = section.footer
        f_p = footer.paragraphs[0]
        f_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        f_run = f_p.add_run("النشرة الأسبوعية للحوكمة والتدقيق — سري وللاستخدام الداخلي فقط")
        f_run.font.name = "Traditional Arabic"
        f_run.font.size = Pt(9)
        f_run.font.color.rgb = RGBColor(0x94, 0xA3, 0xB8)
        
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("النشرة الأسبوعية للحوكمة والتدقيق الداخلي")
    run_title.font.name = "Traditional Arabic"
    run_title.font.size = Pt(22)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0xB7, 0x34, 0xF7)
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = p_sub.add_run(f"تقرير رصد التحليلات والتطورات التنظيمية — {datetime.now().strftime('%Y-%m-%d')}")
    run_sub.font.name = "Traditional Arabic"
    run_sub.font.size = Pt(12)
    run_sub.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)
    doc.add_paragraph().paragraph_format.space_after = Pt(8)
    
    add_executive_dashboard(doc, content_blocks)
    
    for block in content_blocks:
        create_section_header(
            doc, 
            block.get("title", "موضوع رئيسي"), 
            risk=block.get("risk_level", "متوسط"),
            category=block.get("category", "حوكمة")
        )
        
        p_body = doc.add_paragraph()
        p_body.paragraph_format.line_spacing = 1.3
        p_body.paragraph_format.space_after = Pt(6)
        p_body.paragraph_format.space_before = Pt(6)
        run_b = p_body.add_run(block.get("summary", ""))
        run_b.font.name = "Traditional Arabic"
        run_b.font.size = Pt(12.5)
        run_b.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)
        
        add_action_box(
            doc, 
            why_it_matters=block.get("why_it_matters", ""), 
            recommended_actions=block.get("recommended_actions", "")
        )
        
        doc.add_paragraph().paragraph_format.space_after = Pt(10)
        
    doc.save(filename)
    logging.info(f"Generated Word document in reports folder: {filename}")
    return filename

def convert_docx_to_pdf(docx_path: str) -> str:
    """Converts the DOCX document to PDF inside the reports folder."""
    pdf_path = docx_path.rsplit('.', 1)[0] + ".pdf"
    
    try:
        # Preferred method for Windows / macOS (Requires MS Word installed)
        from docx2pdf import convert
        convert(docx_path, pdf_path)
        logging.info(f"Successfully converted DOCX to PDF using docx2pdf: {pdf_path}")
        return pdf_path
    except Exception as e_docx2pdf:
        logging.warning(f"docx2pdf conversion failed or unavailable: {e_docx2pdf}. Attempting LibreOffice fallback...")
        
    try:
        # Fallback method (Ideal for Linux / headless servers with LibreOffice)
        output_dir = os.path.dirname(docx_path)
        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", docx_path, "--outdir", output_dir],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        logging.info(f"Successfully converted DOCX to PDF using LibreOffice: {pdf_path}")
        return pdf_path
    except Exception as e_libreoffice:
        logging.error(f"Failed to convert DOCX to PDF using LibreOffice: {e_libreoffice}")
        
    # Return docx_path if conversion fails completely
    return docx_path

def generate_email_html(content_blocks: List[Dict[str, Any]]) -> str:
    """Generates a responsive high-contrast HTML Email."""
    cards_html = ""
    for block in content_blocks:
        risk = block.get("risk_level", "متوسط")
        risk_color = RISK_COLORS.get(risk, "854D0E")
        
        cards_html += f"""
        <div style="background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; margin-bottom: 24px; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.03);">
            <div style="background-color: #{NAVY_PRIMARY}; padding: 14px 18px; text-align: right; display: flex; justify-content: space-between; align-items: center;">
                <h2 style="color: #{TEXT_WHITE}; margin: 0; font-family: 'Segoe UI', Tahoma, sans-serif; font-size: 16px; font-weight: bold;">
                    {block.get('title', '')}
                </h2>
                <span style="background-color: #{risk_color}; color: #ffffff; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: bold; margin-right: 10px;">
                    {risk}
                </span>
            </div>
            <div style="padding: 18px; text-align: right; font-family: 'Segoe UI', Tahoma, sans-serif;">
                <p style="color: #{TEXT_DARK}; font-size: 14.5px; line-height: 1.7; margin-top: 0; margin-bottom: 14px;">
                    {block.get('summary', '')}
                </p>
                <div style="background-color: #{BG_LIGHT}; border-right: 4px solid #{ACCENT_BLUE}; padding: 14px 16px; border-radius: 4px;">
                    <div style="margin-bottom: 8px;">
                        <strong style="color: #{ACCENT_BLUE}; font-size: 13.5px;">🎯 الأثر المباشر (Why it Matters):</strong>
                        <p style="color: #334155; font-size: 13.5px; margin: 4px 0 0 0; line-height: 1.6;">{block.get('why_it_matters', '')}</p>
                    </div>
                    <div style="margin-top: 10px;">
                        <strong style="color: #{ACCENT_BLUE}; font-size: 13.5px;">🛠️ إجراءات العمل الموصى بها:</strong>
                        <p style="color: #334155; font-size: 13.5px; margin: 4px 0 0 0; line-height: 1.6;">{block.get('recommended_actions', '')}</p>
                    </div>
                </div>
            </div>
        </div>
        """

    return f"""
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body style="background-color: #f1f5f9; margin: 0; padding: 20px; font-family: 'Segoe UI', Tahoma, sans-serif;">
        <div style="max-width: 680px; margin: 0 auto;">
            <div style="background: linear-gradient(135deg, #{NAVY_PRIMARY} 0%, #7C22A8 100%); padding: 28px 20px; border-radius: 8px 8px 0 0; text-align: center;">
                <h1 style="color: #ffffff; margin: 0; font-size: 21px; font-weight: bold;">
                    النشرة الأسبوعية للحوكمة والتدقيق الداخلي
                </h1>
                <p style="color: #f3e8ff; margin: 8px 0 0 0; font-size: 13px;">
                    رصد التحليلات والتطورات التنظيمية — {datetime.now().strftime('%Y-%m-%d')}
                </p>
            </div>
            <div style="padding: 20px 0;">
                {cards_html}
            </div>
            <div style="text-align: center; padding: 15px; font-size: 12px; color: #64748b; border-top: 1px solid #e2e8f0;">
                هذه النشرة مُعدّة آلياً عبر نظام رصد الرؤى والتطورات التنظيمية للحوكمة والتدقيق.<br>
                جميع الحقوق محفوظة © {datetime.now().year}
            </div>
        </div>
    </body>
    </html>
    """

def send_email_report(html_content: str, pdf_path: str):
    """Sends the HTML report along with the PDF attachment via SMTP."""
    smtp_server = os.getenv("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(os.getenv("SMTP_PORT", 465))
    sender_email = os.getenv("SENDER_EMAIL")
    sender_password = os.getenv("SENDER_PASSWORD")
    recipient_raw = os.getenv("RECIPIENT_EMAILS", "")
    recipient_emails = [e.strip() for e in recipient_raw.split(",") if e.strip()]

    if not sender_email or not sender_password or not recipient_emails:
        logging.warning("SMTP environment variables missing or invalid. Email was not dispatched.")
        return

    # Use 'mixed' multipart to allow both inline HTML body and attached files
    msg = MIMEMultipart("mixed")
    msg["Subject"] = f"النشرة الأسبوعية للحوكمة والتدقيق الداخلي — {datetime.now().strftime('%Y-%m-%d')}"
    msg["From"] = sender_email
    msg["To"] = ", ".join(recipient_emails)

    # Attach HTML Body
    msg.attach(MIMEText(html_content, "html", "utf-8"))

    # Attach PDF Document
    if os.path.exists(pdf_path) and pdf_path.endswith(".pdf"):
        with open(pdf_path, "rb") as f:
            attachment = MIMEApplication(f.read(), _subtype="pdf")
            attachment.add_header("Content-Disposition", "attachment", filename=os.path.basename(pdf_path))
            msg.attach(attachment)
            logging.info(f"Attached PDF report: {os.path.basename(pdf_path)}")
    else:
        logging.warning(f"Target PDF file was not found or is unreadable: {pdf_path}")

    try:
        if smtp_port == 465:
            # SSL Connection
            with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                server.login(sender_email, sender_password)
                server.sendmail(sender_email, recipient_emails, msg.as_string())
        else:
            # TLS Connection (Standard for port 587)
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.sendmail(sender_email, recipient_emails, msg.as_string())
                
        logging.info(f"Email newsletter successfully dispatched to {len(recipient_emails)} recipient(s).")
    except Exception as e:
        logging.error(f"Failed to dispatch email report: {str(e)}")

def run_pipeline():
    """Main pipeline entry point."""
    sample_data = [
        {
            "title": "1. تحديثات معايير التدقيق القائم على إدارة المخاطر (ERM)",
            "category": "تدقيق داخلي",
            "risk_level": "عالٍ",
            "summary": "أصدرت الهيئات التنظيمية خطوطاً إرشادية جديدة تركز على رفع كفاءة التقييم المستمر لمخاطر الأمن السيبراني والتحول الرقمي ضمن خطط التدقيق السنوية.",
            "why_it_matters": "تزايد المخاطر السيبرانية قد يؤدي إلى انكشاف النظم الرقابية وغرامات تنظيمية غير متوقعة.",
            "recommended_actions": "تحديث مصفوفة مخاطر التدقيق الداخلي وإعادة تقييم ضوابط الوصول للأنظمة الحساسة قبل نهاية الربع."
        },
        {
            "title": "2. تعزيز أطر الحوكمة في التطبيقات المالية المؤتمتة",
            "category": "حوكمة وتقنية",
            "risk_level": "حرج",
            "summary": "التأكيد على ضرورة فصل المهام وتفعيل آليات المراجعة المزدوجة للعمليات المالية المؤتمتة لمنع التعارض والأخطاء التشغيلية.",
            "why_it_matters": "غياب الرقابة المزدوجة يزيد احتمال الاحتيال المالي والأخطاء المحاسبية غير المكتشفة.",
            "recommended_actions": "إجراء فحص عينات للأنظمة المؤتمتة للتحقق من مطابقة صلاحيات المستخدمين لمصفوفة (SoD)."
        }
    ]

    # 1. Build Word Report inside 'reports' folder
    docx_filepath = build_word_document(sample_data, reports_dir="reports")
    
    # 2. Convert Word Document to PDF
    pdf_filepath = convert_docx_to_pdf(docx_filepath)
    
    # 3. Build HTML Email Content
    email_html = generate_email_html(sample_data)
    
    # 4. Save preview file locally
    with open("email_preview.html", "w", encoding="utf-8") as f:
        f.write(email_html)
        
    # 5. Dispatch Email with PDF Attachment
    send_email_report(email_html, pdf_filepath)

if __name__ == "__main__":
    run_pipeline()
