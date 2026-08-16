# -*- coding: utf-8 -*-
"""
============================================================================
النشرة الأسبوعية للحوكمة والتدقيق الداخلي — Governance & Audit Weekly
============================================================================
Pipeline:
  1) Gemini (with live Google Search grounding) researches this week's most
     relevant developments across 8 professional domains, using a curated,
     deep-researched list of authoritative sources, and returns strict
     structured JSON.
  2) A luxury-styled DOCX is built from that JSON (cover, executive
     dashboard, quick wins, "management flashback", per-domain briefs with
     real hyperlinks, further-reading / book picks).
  3) DOCX -> PDF.
  4) A matching luxury HTML email is generated and dispatched.

  If step 1 fails for any reason, the pipeline STOPS (non-zero exit) and
  does NOT send an email. A previous version emailed a "sorry, this failed"
  placeholder to the real distribution list, which is worse than sending
  nothing — that behaviour has been removed.

----------------------------------------------------------------------------
REQUIRED ENVIRONMENT VARIABLES
----------------------------------------------------------------------------
  GEMINI_API_KEY  (or GOOGLE_API_KEY)   Google AI Studio API key
  SENDER_EMAIL, SENDER_PASSWORD, RECIPIENT_EMAILS   SMTP credentials

OPTIONAL:
  GEMINI_MODEL            default "gemini-2.5-flash"
  SMTP_SERVER, SMTP_PORT  default smtp.gmail.com / 465

----------------------------------------------------------------------------
REQUIRED PYTHON PACKAGES (requirements.txt)
----------------------------------------------------------------------------
  google-genai          <- the CURRENT SDK. The old "google-generativeai"
                            package reached end-of-life on 2025-11-30 and
                            must not be used.
  python-docx
  python-dotenv          (optional, only for local .env loading)
  docx2pdf                (optional — Windows/macOS only; on Linux CI the
                            pipeline falls back to LibreOffice automatically)

----------------------------------------------------------------------------
REQUIRED SYSTEM PACKAGES ON THE CI RUNNER (GitHub Actions ubuntu-latest)
----------------------------------------------------------------------------
This is the fix for Arabic text rendering as tofu/question marks in the
PDF: standard Ubuntu CI runners do NOT ship "Calibri" and do NOT ship any
font with proper Arabic script shaping. Word/LibreOffice then silently
substitute a fallback font that has no Arabic glyphs, which is exactly
what produced the "?????" you saw. Add this step to the workflow BEFORE
the step that runs this script:

    - name: Install fonts + LibreOffice
      run: |
        sudo apt-get update
        sudo apt-get install -y \\
          libreoffice \\
          fonts-hosny-amiri \\
          fonts-crosextra-carlito

  fonts-hosny-amiri    -> installs "Amiri", a classical, professionally
                           designed Naskh Arabic typeface (used on 65,000+
                           sites, served via Google Fonts) — this is the
                           font this script now asks for explicitly.
  fonts-crosextra-carlito -> installs "Carlito", the metric-compatible
                           open-source clone of Calibri, for the Latin /
                           English-technical-term portions of the text.

At start-up this script checks (via `fc-list`) whether Amiri is actually
installed and logs a loud, explicit warning — instead of silently
producing broken PDFs again — if it is not.
============================================================================
"""

import os
import re
import sys
import json
import logging
import smtplib
import subprocess
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.header import Header
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
log = logging.getLogger("newsletter")

# ============================================================================
# 1. BRAND / DESIGN SYSTEM
# ============================================================================
# A restrained "private-bank" palette: deep navy + warm gold on ivory.
# (The previous palette had a bug: colours named NAVY/BLUE were actually
#  magenta/pink hex values. Fixed here.)

NAVY_DARK    = "0B1F3A"   # Cover / primary header fill
NAVY_PRIMARY = "13294B"   # Section header fill
GOLD_ACCENT  = "C9A227"   # Premium accent, rules, badges
GOLD_LIGHT   = "E8D9A0"   # Subtle gold tint
IVORY_BG     = "FBF9F4"   # Card / callout background
HAIRLINE     = "D8D2C2"   # Thin borders
TEXT_DARK    = "1C2430"   # Body text
TEXT_MUTED   = "5B6472"   # Secondary text
TEXT_WHITE   = "FFFFFF"

RISK_COLORS = {
    "حرج": "8C1D1D",
    "عالٍ": "B5541A",
    "متوسط": "8A6D1A",
    "منخفض": "1E6B3E",
}

# Two distinct fonts, deliberately: one for Arabic (complex-script) runs
# and one for Latin (ascii) runs within the same paragraph. python-docx's
# font.name only sets w:ascii — Word/LibreOffice then pick whatever their
# *default* complex-script font is for the Arabic characters, which on a
# bare Linux CI box has no Arabic glyphs at all. We set w:cs explicitly
# instead of leaving it to chance. See the module docstring for the apt
# packages that must be installed for these to actually be present.
FONT_ARABIC = "Amiri"     # complex-script (w:cs) — classical Naskh, elegant
FONT_LATIN = "Carlito"    # ascii (w:ascii) — open, metric-compatible Calibri

ISSUE_NO_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".issue_number")


def verify_fonts_installed():
    """Checks (via fontconfig) whether the fonts this document depends on
    are actually present, and logs a loud, actionable warning if not —
    so a broken-looking PDF is diagnosable instead of a silent surprise."""
    for font_name, apt_pkg in [(FONT_ARABIC, "fonts-hosny-amiri"), (FONT_LATIN, "fonts-crosextra-carlito")]:
        try:
            result = subprocess.run(["fc-list", f":family={font_name}"], capture_output=True, text=True, timeout=10)
            if not result.stdout.strip():
                log.warning(
                    f"⚠ Font '{font_name}' was NOT found by fontconfig on this machine. "
                    f"Arabic/Latin text may render as missing glyphs or question marks. "
                    f"Fix: sudo apt-get install -y {apt_pkg}"
                )
            else:
                log.info(f"Font check OK: '{font_name}' is installed.")
        except FileNotFoundError:
            log.warning("fc-list not found — cannot verify font availability (fontconfig may be missing).")
        except Exception as e:
            log.warning(f"Font check for '{font_name}' failed: {e}")


def next_issue_number() -> int:
    """Persists and increments a simple issue counter across runs."""
    n = 1
    try:
        if os.path.exists(ISSUE_NO_FILE):
            with open(ISSUE_NO_FILE, "r") as f:
                n = int(f.read().strip() or "0") + 1
    except Exception:
        n = 1
    try:
        with open(ISSUE_NO_FILE, "w") as f:
            f.write(str(n))
    except Exception:
        pass
    return n


# ============================================================================
# 2. CURATED SOURCE LIBRARY  (used to steer Gemini's grounded research)
# ============================================================================
# Base list = exactly what was supplied originally. Additions below are
# marked "# added" and were individually verified by live web search
# (not guessed) — each is a currently-live, authoritative body directly
# relevant to one of the newsletter's 8 domains but missing from the
# original sample list.
SOURCES = {
    "تدقيق داخلي وأداء (Internal / Performance Audit)": [
        ("IIA - The Institute of Internal Auditors", "https://www.theiia.org/"),
        ("INTOSAI", "https://www.intosai.org/"),
        ("ISSAI - INTOSAI Standards", "https://www.issai.org/"),               # added
        ("ECIIA - European Confederation of Institutes of Internal Auditing", "https://www.eciia.eu/"),  # added
        ("NAO UK", "https://www.nao.org.uk"),
        ("GAO US", "https://www.gao.gov"),
        ("AuditNet", "https://www.auditnet.org"),
        ("Internal Audit 360", "https://www.internalaudit360.com"),
        ("IIA Australia", "https://www.iia.org.au"),
        ("Saudi Internal Audit Association", "https://www.saia.gov.sa"),
    ],
    "الحوكمة والقطاع العام (Governance / Public Sector)": [
        ("OECD Public Governance", "https://www.oecd.org/governance/"),
        ("World Bank Governance", "https://www.worldbank.org/"),
        ("LSE - London School of Economics", "https://www.lse.ac.uk"),
        ("ICGN - International Corporate Governance Network", "https://www.icgn.org/"),  # added
        ("Transparency International", "https://www.transparency.org/"),      # added
        ("Ash Center for Democratic Governance, Harvard Kennedy School", "https://ash.harvard.edu/"),  # added
    ],
    "الرقابة الداخلية وإدارة المخاطر (Internal Control / Risk / COSO)": [
        ("COSO", "https://www.coso.org/"),
        ("IFAC", "https://www.ifac.org/"),
        ("Protiviti", "https://www.protiviti.com"),
        ("ISO 31000 - Risk Management", "https://www.iso.org/standards/popular/iso-31000-family"),  # added
        ("GARP - Global Association of Risk Professionals", "https://www.garp.org/"),  # added
        ("FERMA - Federation of European Risk Management Associations", "https://ferma.eu/"),  # added
    ],
    "الموارد البشرية (HR)": [                                                  # added (new category)
        ("SHRM - Society for Human Resource Management", "https://www.shrm.org/"),
        ("CIPD - Chartered Institute of Personnel and Development", "https://www.cipd.org/"),
    ],
    "المحاسبة الإدارية (Management Accounting)": [                              # added (new category)
        ("IMA - Institute of Management Accountants", "https://www.imanet.org/"),
    ],
    "استشارات وأفضل الممارسات (Big Four / Strategy Insights)": [
        ("Deloitte Insights", "https://www.deloitte.com/"),
        ("EY Insights", "https://www.ey.com/"),
        ("KPMG Insights", "https://kpmg.com/"),
        ("PwC Insights", "https://www.pwc.com/"),
        ("McKinsey", "https://www.mckinsey.com"),
        ("Harvard Business Review", "https://hbr.org"),
    ],
}

DOMAINS = [
    "التدقيق الداخلي (Internal Audit)",
    "تدقيق الأداء (Performance Audit)",
    "الحوكمة (Governance)",
    "الرقابة الداخلية (Internal Controls)",
    "إدارة المخاطر (Risk Management)",
    "الموارد البشرية (HR)",
    "المحاسبة الإدارية (Management Accounting)",
    "إصلاح وتحول القطاع العام (Public Sector Reform)",
]

# Rotates weekly so the "flashback" never repeats two weeks running.
FLASHBACK_TOPICS = [
    "إطار COSO للرقابة الداخلية (COSO Internal Control Framework)",
    "إدارة الموارد البشرية الاستراتيجية (Strategic HR Management)",
    "المراقبة التسييرية (Contrôle de Gestion / Management Control)",
    "إدارة الأداء المؤسسي (Performance Management)",
    "تطوير المواهب وتخطيط الإحلال الوظيفي (Talent Development & Succession Planning)",
    "المعايير المحاسبية الدولية (IPSAS / IFRS Essentials)",
    "إدارة التغيير المؤسسي (Change Management - Kotter's 8 Steps)",
    "نموذج النضج الرقابي وثلاثة خطوط الدفاع (Three Lines Model)",
]


def pick_flashback_topic() -> str:
    week = datetime.now().isocalendar()[1]
    return FLASHBACK_TOPICS[week % len(FLASHBACK_TOPICS)]


# ============================================================================
# 3. CONTENT GENERATION — Gemini with live Google Search grounding
# ============================================================================
RESPONSE_SCHEMA_HINT = """
أعد النتيجة ككائن JSON صِرف واحد فقط (بدون أي نص قبله أو بعده، بدون Markdown fences)
مطابقًا تمامًا لهذا الهيكل:

{
  "issue_theme": "عنوان جذاب لموضوع العدد هذا الأسبوع (جملة قصيرة)",
  "editor_note": "فقرة افتتاحية قصيرة (3-4 أسطر) بأسلوب تنفيذي راقٍ",
  "quick_wins": [
    {"tip_ar": "نصيحة عملية قابلة للتطبيق فورًا بجملة أو جملتين مع المصطلح الإنجليزي بين قوسين"}
    // 3 عناصر بالضبط
  ],
  "flashback": {
    "topic": "اسم الموضوع الإداري الكلاسيكي المرسل إليك بالضبط",
    "content_ar": "شرح مكثف وعملي (فقرة واحدة، 4-6 أسطر) لماذا هذا المفهوم لا يزال أساسيًا اليوم",
    "key_points": ["نقطة تطبيقية 1", "نقطة تطبيقية 2", "نقطة تطبيقية 3"]
  },
  "domain_updates": [
    {
      "domain": "اسم المجال كما ورد في القائمة",
      "title": "عنوان التحديث/الخبر",
      "category": "تصنيف قصير",
      "risk_level": "حرج | عالٍ | متوسط | منخفض",
      "summary": "ملخص التطور الأحدث في هذا المجال (2-3 جمل) مع ترجمة أي مصطلح تقني إلى الإنجليزية بين قوسين",
      "why_it_matters": "الأثر المباشر على المؤسسة (جملتان)",
      "recommended_actions": "إجراءات عملية وقابلة للتنفيذ فورًا (جملتان، صيغة أفعال أمر)",
      "source_name": "اسم المصدر الفعلي الذي استندت إليه",
      "source_url": "رابط حقيقي وقابل للفتح للمصدر"
    }
    // عنصر واحد لكل مجال من المجالات الثمانية المرسلة، بحد أقصى 6 عناصر (اختر الأهم)
  ],
  "further_reading": [
    {"title": "عنوان المقال/التقرير", "source_name": "اسم الجهة", "url": "رابط حقيقي", "type": "مقال | تقرير | دراسة"}
    // 4 عناصر
  ],
  "book_recommendation": {
    "title": "عنوان كتاب معروف وذو صلة بأحد المجالات",
    "author": "اسم المؤلف",
    "why_ar": "لماذا يستحق القراءة هذا الأسبوع تحديدًا (جملتان)"
  }
}

قواعد إلزامية:
- اكتب المحتوى بالعربية الفصحى الاحترافية، وضع كل مصطلح تقني بالإنجليزية بين قوسين عند أول ورود له، مثال: «معايير الأداء (Performance Standards)».
- استخدم فقط أخبارًا وتطورات حقيقية وحديثة (لا تختلق حقائق)، واستشهد بروابط قابلة للفتح فعليًا من نتائج البحث.
- لا تكرر نفس الخبر في أكثر من مجال.
- اجعل اللهجة عملية ومباشرة (hands-on)، بلا حشو إنشائي.
"""


def build_research_prompt() -> str:
    sources_block = "\n".join(
        f"- {cat}:\n  " + "، ".join(f"{name} ({url})" for name, url in items)
        for cat, items in SOURCES.items()
    )
    domains_block = "\n".join(f"- {d}" for d in DOMAINS)
    flashback_topic = pick_flashback_topic()

    return f"""أنت محرر تنفيذي متخصص في التدقيق الداخلي والحوكمة وإدارة المخاطر، تُعِدّ نشرة أسبوعية
احترافية لجمهور من كبار المدققين والمسؤولين الحكوميين والماليين التنفيذيين. مستوى الجودة
المطلوب أعلى من نشرات Deloitte وMcKinsey وHBR: مكثف، عملي، بلا حشو، بقيمة مضافة حقيقية.

استخدم بحث Google المباشر للعثور على أحدث التطورات والمنشورات (آخر 1-3 أسابيع قدر الإمكان)
من هذه المصادر الموثوقة حصرًا أو مصادر بنفس المستوى من المصداقية:

{sources_block}

غطِّ المجالات التالية (اختر الأهم وليس بالضرورة كل مجال إن لم يوجد جديد يستحق الذكر):
{domains_block}

موضوع "ومضة إدارية" (Flashback) لهذا الأسبوع تحديدًا هو: {flashback_topic}
اشرحه بإيجاز عملي دون البحث عنه (معرفة عامة كافية).

{RESPONSE_SCHEMA_HINT}
"""


def _extract_json(text: str) -> Optional[dict]:
    """Gemini sometimes wraps JSON in ```json fences despite instructions."""
    if not text:
        return None
    cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", text.strip(), flags=re.MULTILINE)
    match = re.search(r"\{.*\}", cleaned, flags=re.DOTALL)
    candidate = match.group(0) if match else cleaned
    try:
        return json.loads(candidate)
    except json.JSONDecodeError as e:
        log.error(f"JSON parse failure: {e}")
        return None


def generate_newsletter_content() -> Tuple[Optional[Dict[str, Any]], bool]:
    """Calls Gemini (grounded with live Google Search) to research and draft
    this week's issue.

    Returns (content, success). On ANY failure, success=False and content is
    None — the caller (run_pipeline) must NOT proceed to build/send a
    newsletter in that case. We deliberately do not silently substitute
    placeholder content here any more: a previous version did that and the
    placeholder text ended up being emailed to the real distribution list,
    which is a worse outcome than the run simply failing loudly.
    """
    api_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
    model_name = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")

    if not api_key:
        log.error(
            "FATAL: neither GEMINI_API_KEY nor GOOGLE_API_KEY is set in the environment. "
            "Set one of them as a GitHub Actions secret and expose it to this job's env block."
        )
        return None, False

    try:
        from google import genai
        from google.genai import types
    except ImportError as e:
        log.error(
            f"FATAL: could not import the 'google-genai' package ({e}). "
            "It is very likely missing from requirements.txt (or the old, "
            "end-of-life 'google-generativeai' package is installed instead). "
            "Fix: add 'google-genai' to requirements.txt and re-run "
            "`pip install -r requirements.txt`."
        )
        return None, False

    try:
        client = genai.Client(api_key=api_key)
        grounding_tool = types.Tool(google_search=types.GoogleSearch())
        config = types.GenerateContentConfig(tools=[grounding_tool], temperature=0.4)
        response = client.models.generate_content(
            model=model_name,
            contents=build_research_prompt(),
            config=config,
        )
    except Exception as e:
        log.error(
            f"FATAL: the Gemini API call itself failed ({type(e).__name__}: {e}). "
            "Common causes: invalid/expired API key, the API key's project has no "
            "billing/quota enabled, or the model name in GEMINI_MODEL doesn't exist. "
            "Check https://aistudio.google.com/app/apikey for key status.",
            exc_info=True,
        )
        return None, False

    if not getattr(response, "text", None):
        log.error(
            "FATAL: Gemini returned an empty response (no text). This usually means the "
            "prompt was blocked by a safety filter, or the model returned only a "
            "function/tool call. Inspect response.candidates in a local debug run."
        )
        return None, False

    data = _extract_json(response.text)
    if not data:
        log.error(
            "FATAL: Gemini's response could not be parsed as JSON. Raw response has been "
            "logged below for debugging — check for extra prose around the JSON object."
        )
        log.error(f"Raw Gemini response (first 2000 chars): {response.text[:2000]}")
        return None, False

    data.setdefault("flashback", {}).setdefault("topic", pick_flashback_topic())
    log.info(f"Content generated successfully: {len(data.get('domain_updates', []))} domain updates.")
    return data, True


# ============================================================================
# 4. DOCX HELPERS
# ============================================================================
def set_rtl(paragraph):
    """Marks a paragraph as right-to-left / bidirectional so Word renders
    Arabic runs correctly regardless of alignment."""
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement("w:bidi")
    pPr.append(bidi)


def style_run(run, size=11, bold=False, color=TEXT_DARK, italic=False):
    """Applies size/weight/colour, and — critically — sets DIFFERENT fonts
    for the Latin (w:ascii) and Arabic/complex-script (w:cs) glyph ranges
    within the same run, so mixed Arabic + English-technical-term text
    renders correctly instead of falling back to a font with no Arabic
    glyphs (the cause of the "?????" rendering bug)."""
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor.from_string(color)
    run.font.name = FONT_LATIN
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), FONT_LATIN)
    rFonts.set(qn("w:hAnsi"), FONT_LATIN)
    rFonts.set(qn("w:cs"), FONT_ARABIC)
    rFonts.set(qn("w:eastAsia"), FONT_ARABIC)
    return run


def add_hyperlink(paragraph, text, url, color=GOLD_ACCENT, underline=True, size=10.5, bold=False):
    """Inserts a real, clickable hyperlink run into a docx paragraph."""
    part = paragraph.part
    r_id = part.relate_to(
        url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), FONT_LATIN)
    rFonts.set(qn("w:hAnsi"), FONT_LATIN)
    rFonts.set(qn("w:cs"), FONT_ARABIC)
    rFonts.set(qn("w:eastAsia"), FONT_ARABIC)
    rPr.append(rFonts)

    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size * 2)))
    rPr.append(sz)

    if bold:
        rPr.append(OxmlElement("w:b"))
    if underline:
        u = OxmlElement("w:u")
        u.set(qn("w:val"), "single")
        rPr.append(u)

    color_el = OxmlElement("w:color")
    color_el.set(qn("w:val"), color)
    rPr.append(color_el)

    new_run.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink


def set_cell_background(cell, color_hex: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for name, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{name}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def add_gold_rule(doc, space_before=2, space_after=10):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), GOLD_ACCENT)
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


# ============================================================================
# 5. DOCX BUILD
# ============================================================================
def add_cover_page(doc, content: Dict[str, Any], issue_no: int):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, NAVY_DARK)
    set_cell_margins(cell, top=900, bottom=900, left=500, right=500)

    p_kicker = cell.paragraphs[0]
    p_kicker.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_rtl(p_kicker)
    style_run(p_kicker.add_run(f"العدد {issue_no}  •  {datetime.now().strftime('%Y-%m-%d')}"),
              size=11, color=GOLD_LIGHT, bold=True)

    p_title = cell.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_rtl(p_title)
    p_title.paragraph_format.space_before = Pt(14)
    style_run(p_title.add_run("النشرة الأسبوعية للحوكمة والتدقيق الداخلي"),
              size=26, bold=True, color=TEXT_WHITE)

    p_theme = cell.add_paragraph()
    p_theme.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_rtl(p_theme)
    p_theme.paragraph_format.space_before = Pt(10)
    style_run(p_theme.add_run(content.get("issue_theme", "")), size=14, italic=True, color=GOLD_ACCENT)

    p_line = cell.add_paragraph()
    p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_line.paragraph_format.space_before = Pt(16)
    pPr = p_line._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), "6")
    top.set(qn("w:space"), "1")
    top.set(qn("w:color"), GOLD_ACCENT)
    pBdr.append(top)
    pPr.append(pBdr)

    p_sub = cell.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_rtl(p_sub)
    p_sub.paragraph_format.space_before = Pt(10)
    style_run(
        p_sub.add_run("تدقيق داخلي • تدقيق أداء • حوكمة • رقابة داخلية • إدارة مخاطر • "
                       "موارد بشرية • محاسبة إدارية • إصلاح القطاع العام"),
        size=10, color=GOLD_LIGHT,
    )
    doc.add_paragraph()


def add_editor_note(doc, content: Dict[str, Any]):
    p = doc.add_paragraph()
    set_rtl(p)
    p.paragraph_format.space_after = Pt(4)
    style_run(p.add_run("كلمة التحرير"), size=13, bold=True, color=NAVY_PRIMARY)

    p2 = doc.add_paragraph()
    set_rtl(p2)
    p2.paragraph_format.line_spacing = 1.3
    style_run(p2.add_run(content.get("editor_note", "")), size=11.5, color=TEXT_MUTED, italic=True)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def add_executive_dashboard(doc, updates: List[Dict[str, Any]]):
    tbl = doc.add_table(rows=2, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    total = len(updates)
    critical_high = sum(1 for b in updates if b.get("risk_level") in ("حرج", "عالٍ"))
    domains_covered = len(set(b.get("domain", "") for b in updates))

    headers = ["إجمالي التحديثات", "تحديثات عالية الأولوية", "المجالات المغطاة"]
    values = [str(total), str(critical_high), str(domains_covered)]
    for i, (h, v) in enumerate(zip(headers, values)):
        top_cell = tbl.cell(0, i)
        set_cell_background(top_cell, NAVY_PRIMARY)
        set_cell_margins(top_cell, top=100, bottom=80)
        p = top_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_rtl(p)
        style_run(p.add_run(h), size=10, bold=True, color=TEXT_WHITE)

        bottom_cell = tbl.cell(1, i)
        set_cell_background(bottom_cell, IVORY_BG)
        set_cell_margins(bottom_cell, top=140, bottom=140)
        p2 = bottom_cell.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        style_run(p2.add_run(v), size=18, bold=True,
                  color=(GOLD_ACCENT if i != 1 else RISK_COLORS["حرج"]))
    doc.add_paragraph().paragraph_format.space_after = Pt(8)


def add_quick_wins(doc, quick_wins: List[Dict[str, str]]):
    if not quick_wins:
        return
    header = doc.add_paragraph()
    set_rtl(header)
    style_run(header.add_run("انتصارات سريعة (Quick Wins)"), size=13, bold=True, color=NAVY_PRIMARY)

    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    set_cell_background(cell, IVORY_BG)
    set_cell_margins(cell, top=120, bottom=120, left=200, right=200)
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="none"/><w:bottom w:val="none"/><w:left w:val="none"/>'
        f'<w:right w:val="single" w:sz="20" w:space="0" w:color="{GOLD_ACCENT}"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)
    first = True
    for item in quick_wins:
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        set_rtl(p)
        p.paragraph_format.space_after = Pt(6)
        style_run(p.add_run("— "), size=11, bold=True, color=GOLD_ACCENT)
        style_run(p.add_run(item.get("tip_ar", "")), size=11, color=TEXT_DARK)
    doc.add_paragraph().paragraph_format.space_after = Pt(8)


def add_flashback(doc, flashback: Dict[str, Any]):
    if not flashback or not flashback.get("content_ar"):
        return
    header = doc.add_paragraph()
    set_rtl(header)
    style_run(header.add_run(f"ومضة إدارية — {flashback.get('topic', '')}"),
              size=13, bold=True, color=NAVY_PRIMARY)

    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    set_cell_background(cell, "FFFFFF")
    set_cell_margins(cell, top=120, bottom=120, left=200, right=200)
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="{HAIRLINE}"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="{HAIRLINE}"/>'
        f'<w:left w:val="single" w:sz="4" w:space="0" w:color="{HAIRLINE}"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="{HAIRLINE}"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)

    p = cell.paragraphs[0]
    set_rtl(p)
    p.paragraph_format.line_spacing = 1.3
    style_run(p.add_run(flashback.get("content_ar", "")), size=11, color=TEXT_DARK)

    for kp in flashback.get("key_points", []):
        pk = cell.add_paragraph()
        set_rtl(pk)
        pk.paragraph_format.space_before = Pt(4)
        style_run(pk.add_run("— "), size=10.5, bold=True, color=GOLD_ACCENT)
        style_run(pk.add_run(kp), size=10.5, color=TEXT_MUTED)
    doc.add_paragraph().paragraph_format.space_after = Pt(8)


def create_section_header(doc, title: str, risk: str, category: str):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, NAVY_PRIMARY)
    set_cell_margins(cell, top=130, bottom=130, left=200, right=200)

    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_rtl(p)
    style_run(p.add_run(f"{title}  "), size=13.5, bold=True, color=TEXT_WHITE)
    risk_color = RISK_COLORS.get(risk, "8A6D1A")
    tag = p.add_run(f" [{category} | {risk}]")
    style_run(tag, size=10, bold=True, color=GOLD_LIGHT)


def add_domain_card(doc, block: Dict[str, Any]):
    create_section_header(
        doc,
        block.get("domain", block.get("title", "تحديث")),
        risk=block.get("risk_level", "متوسط"),
        category=block.get("category", "عام"),
    )
    if block.get("title"):
        p_title = doc.add_paragraph()
        set_rtl(p_title)
        p_title.paragraph_format.space_before = Pt(4)
        style_run(p_title.add_run(block["title"]), size=12, bold=True, color=NAVY_PRIMARY)

    p_body = doc.add_paragraph()
    set_rtl(p_body)
    p_body.paragraph_format.line_spacing = 1.3
    p_body.paragraph_format.space_after = Pt(6)
    style_run(p_body.add_run(block.get("summary", "")), size=11, color=TEXT_DARK)

    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    set_cell_background(cell, IVORY_BG)
    set_cell_margins(cell, top=110, bottom=110, left=180, right=180)
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/>'
        f'<w:right w:val="single" w:sz="20" w:space="0" w:color="{RISK_COLORS.get(block.get("risk_level"), GOLD_ACCENT)}"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)

    p1 = cell.paragraphs[0]
    set_rtl(p1)
    style_run(p1.add_run("الأثر (Why It Matters): "), size=10.5, bold=True, color=NAVY_PRIMARY)
    style_run(p1.add_run(block.get("why_it_matters", "")), size=10.5, color=TEXT_DARK)

    p2 = cell.add_paragraph()
    set_rtl(p2)
    p2.paragraph_format.space_before = Pt(4)
    style_run(p2.add_run("إجراءات موصى بها: "), size=10.5, bold=True, color=NAVY_PRIMARY)
    style_run(p2.add_run(block.get("recommended_actions", "")), size=10.5, color=TEXT_DARK)

    if block.get("source_url"):
        p3 = cell.add_paragraph()
        set_rtl(p3)
        p3.paragraph_format.space_before = Pt(4)
        style_run(p3.add_run("المصدر: "), size=9.5, bold=True, color=TEXT_MUTED)
        add_hyperlink(p3, block.get("source_name", block["source_url"]), block["source_url"], size=9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)


def add_further_reading(doc, items: List[Dict[str, str]], book: Dict[str, str]):
    if not items and not (book and book.get("title")):
        return
    header = doc.add_paragraph()
    set_rtl(header)
    style_run(header.add_run("للقراءة هذا الأسبوع (Further Reading)"), size=13, bold=True, color=NAVY_PRIMARY)

    for it in items:
        p = doc.add_paragraph()
        set_rtl(p)
        p.paragraph_format.space_after = Pt(4)
        style_run(p.add_run(f"• {it.get('type', 'مقال')} — "), size=10.5, color=TEXT_MUTED)
        if it.get("url"):
            add_hyperlink(p, it.get("title", ""), it["url"], size=10.5, bold=True)
        else:
            style_run(p.add_run(it.get("title", "")), size=10.5, bold=True, color=TEXT_DARK)
        if it.get("source_name"):
            style_run(p.add_run(f"  ({it['source_name']})"), size=9.5, color=TEXT_MUTED)

    if book and book.get("title"):
        add_gold_rule(doc, space_before=8, space_after=6)
        p = doc.add_paragraph()
        set_rtl(p)
        style_run(p.add_run("كتاب الأسبوع: "), size=11, bold=True, color=NAVY_PRIMARY)
        style_run(p.add_run(f"{book['title']} — {book.get('author', '')}"), size=11, bold=True, color=TEXT_DARK)
        p2 = doc.add_paragraph()
        set_rtl(p2)
        style_run(p2.add_run(book.get("why_ar", "")), size=10.5, italic=True, color=TEXT_MUTED)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)


def build_word_document(content: Dict[str, Any], reports_dir: str = "reports") -> str:
    os.makedirs(reports_dir, exist_ok=True)
    issue_no = next_issue_number()
    filename = os.path.join(
        reports_dir, f"Governance_Weekly_Brief_{datetime.now().strftime('%Y-%m-%d')}_Issue{issue_no}.docx"
    )
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

        footer = section.footer
        f_p = footer.paragraphs[0]
        f_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_rtl(f_p)
        style_run(f_p.add_run("النشرة الأسبوعية للحوكمة والتدقيق — سري وللاستخدام الداخلي فقط  |  "),
                  size=8, color=TEXT_MUTED)
        fldChar1 = OxmlElement("w:fldChar"); fldChar1.set(qn("w:fldCharType"), "begin")
        instrText = OxmlElement("w:instrText"); instrText.set(qn("xml:space"), "preserve"); instrText.text = "PAGE"
        fldChar2 = OxmlElement("w:fldChar"); fldChar2.set(qn("w:fldCharType"), "end")
        run = f_p.add_run()
        run._r.append(fldChar1); run._r.append(instrText); run._r.append(fldChar2)
        style_run(run, size=8, color=TEXT_MUTED)

    add_cover_page(doc, content, issue_no)
    add_editor_note(doc, content)
    add_executive_dashboard(doc, content.get("domain_updates", []))
    add_quick_wins(doc, content.get("quick_wins", []))
    add_flashback(doc, content.get("flashback", {}))

    updates_header = doc.add_paragraph()
    set_rtl(updates_header)
    style_run(updates_header.add_run("أبرز التطورات حسب المجال"), size=13, bold=True, color=NAVY_PRIMARY)
    add_gold_rule(doc, space_before=2, space_after=8)

    for block in content.get("domain_updates", []):
        add_domain_card(doc, block)

    add_further_reading(doc, content.get("further_reading", []), content.get("book_recommendation", {}))

    doc.save(filename)
    log.info(f"Generated Word document: {filename}")
    return filename


# ============================================================================
# 6. PDF CONVERSION
# ============================================================================
def convert_docx_to_pdf(docx_path: str) -> str:
    pdf_path = docx_path.rsplit(".", 1)[0] + ".pdf"
    try:
        from docx2pdf import convert
        convert(docx_path, pdf_path)
        log.info(f"Converted to PDF via docx2pdf: {pdf_path}")
        return pdf_path
    except Exception as e1:
        log.warning(f"docx2pdf unavailable/failed: {e1}. Trying LibreOffice...")
    try:
        output_dir = os.path.dirname(docx_path)
        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", docx_path, "--outdir", output_dir],
            check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        )
        log.info(f"Converted to PDF via LibreOffice: {pdf_path}")
        return pdf_path
    except Exception as e2:
        log.warning(f"LibreOffice conversion failed: {e2}. Falling back to DOCX attachment.")
    return docx_path


# ============================================================================
# 7. HTML EMAIL (mirrors the DOCX design language)
# ============================================================================
def _html_escape(s: str) -> str:
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def generate_email_html(content: Dict[str, Any], issue_no: int) -> str:
    quick_wins_html = "".join(
        f'<li style="margin-bottom:8px; color:#1C2430;">{_html_escape(q.get("tip_ar",""))}</li>'
        for q in content.get("quick_wins", [])
    )

    fb = content.get("flashback", {})
    flashback_html = ""
    if fb.get("content_ar"):
        pts = "".join(f'<li style="margin-bottom:4px;">{_html_escape(k)}</li>' for k in fb.get("key_points", []))
        flashback_html = f"""
        <div style="background:#FFFFFF; border:1px solid #D8D2C2; border-radius:10px; padding:18px 20px; margin-bottom:26px;">
            <div style="color:#C9A227; font-weight:bold; font-size:12px; letter-spacing:0.5px; margin-bottom:6px; text-transform:uppercase;">ومضة إدارية — {_html_escape(fb.get('topic',''))}</div>
            <p style="color:#1C2430; font-size:14px; line-height:1.7; margin:0 0 8px 0;">{_html_escape(fb.get('content_ar',''))}</p>
            <ul style="margin:0; padding-right:18px; color:#5B6472; font-size:13px;">{pts}</ul>
        </div>"""

    cards_html = ""
    for block in content.get("domain_updates", []):
        risk = block.get("risk_level", "متوسط")
        risk_color = RISK_COLORS.get(risk, "8A6D1A")
        source_html = ""
        if block.get("source_url"):
            source_html = (
                f'<a href="{block["source_url"]}" style="color:#13294B; font-size:12px; text-decoration:underline;">'
                f'المصدر: {_html_escape(block.get("source_name", ""))}</a>'
            )
        cards_html += f"""
        <div style="background:#ffffff; border:1px solid #D8D2C2; border-radius:10px; margin-bottom:20px; overflow:hidden;">
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#13294B;">
              <tr>
                <td style="padding:12px 18px; text-align:right; color:#ffffff; font-size:15px; font-weight:bold; font-family:Calibri, Tahoma, sans-serif;">{_html_escape(block.get('domain',''))}</td>
                <td style="padding:12px 18px; text-align:left; white-space:nowrap;"><span style="background:#{risk_color}; color:#fff; padding:2px 10px; border-radius:10px; font-size:11px; font-weight:bold;">{risk}</span></td>
              </tr>
            </table>
            <div style="padding:16px 18px; text-align:right; font-family:Calibri, Tahoma, sans-serif;">
                <div style="font-weight:bold; color:#13294B; font-size:13.5px; margin-bottom:6px;">{_html_escape(block.get('title',''))}</div>
                <p style="color:#1C2430; font-size:13.5px; line-height:1.7; margin:0 0 10px 0;">{_html_escape(block.get('summary',''))}</p>
                <div style="background:#FBF9F4; border-right:3px solid #C9A227; padding:10px 14px; border-radius:4px;">
                    <div style="margin-bottom:6px;"><strong style="color:#13294B; font-size:12.5px;">الأثر:</strong>
                    <span style="color:#334155; font-size:12.5px;"> {_html_escape(block.get('why_it_matters',''))}</span></div>
                    <div><strong style="color:#13294B; font-size:12.5px;">إجراءات:</strong>
                    <span style="color:#334155; font-size:12.5px;"> {_html_escape(block.get('recommended_actions',''))}</span></div>
                </div>
                <div style="margin-top:10px;">{source_html}</div>
            </div>
        </div>"""

    reading_html = ""
    for it in content.get("further_reading", []):
        link = f'<a href="{it.get("url","#")}" style="color:#13294B; font-weight:bold; text-decoration:none;">{_html_escape(it.get("title",""))}</a>'
        reading_html += (
            f'<li style="margin-bottom:8px; color:#1C2430; font-size:13px;">'
            f'<span style="color:#5B6472;">{_html_escape(it.get("type","مقال"))} — </span>{link}'
            f'<span style="color:#5B6472; font-size:12px;"> ({_html_escape(it.get("source_name",""))})</span></li>'
        )

    book = content.get("book_recommendation", {})
    book_html = ""
    if book and book.get("title"):
        book_html = f"""
        <div style="border-top:1px solid #D8D2C2; margin-top:14px; padding-top:14px;">
            <div style="color:#13294B; font-weight:bold; font-size:13.5px;">كتاب الأسبوع: {_html_escape(book['title'])} — {_html_escape(book.get('author',''))}</div>
            <p style="color:#5B6472; font-size:12.5px; font-style:italic; margin:6px 0 0 0;">{_html_escape(book.get('why_ar',''))}</p>
        </div>"""

    return f"""<!DOCTYPE html>
<html dir="rtl" lang="ar">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"></head>
<body style="background:#F1EFE8; margin:0; padding:20px; font-family:Calibri, Tahoma, sans-serif;">
  <div style="max-width:680px; margin:0 auto;">
    <div style="background:linear-gradient(135deg, #0B1F3A 0%, #13294B 100%); padding:30px 20px; border-radius:10px 10px 0 0; text-align:center;">
      <div style="color:#E8D9A0; font-size:12px; letter-spacing:1px; margin-bottom:6px;">العدد {issue_no} • {datetime.now().strftime('%Y-%m-%d')}</div>
      <h1 style="color:#ffffff; margin:0; font-size:21px;">النشرة الأسبوعية للحوكمة والتدقيق الداخلي</h1>
      <div style="color:#C9A227; margin-top:8px; font-size:14px; font-style:italic;">{_html_escape(content.get('issue_theme',''))}</div>
    </div>

    <div style="background:#ffffff; padding:18px 22px; text-align:right;">
      <p style="color:#5B6472; font-size:13.5px; line-height:1.7; margin:0; font-style:italic;">{_html_escape(content.get('editor_note',''))}</p>
    </div>

    <div style="padding:20px 0 0 0;">
      <div style="background:#FBF9F4; border-right:4px solid #C9A227; border-radius:8px; padding:16px 20px; margin-bottom:26px; text-align:right;">
        <div style="color:#13294B; font-weight:bold; font-size:14px; margin-bottom:8px;">انتصارات سريعة</div>
        <ul style="margin:0; padding-right:18px; font-size:13.5px;">{quick_wins_html}</ul>
      </div>

      <div style="text-align:right;">{flashback_html}</div>

      <div style="text-align:right;">
        <div style="color:#13294B; font-weight:bold; font-size:15px; margin-bottom:14px; border-bottom:2px solid #C9A227; padding-bottom:6px;">أبرز التطورات حسب المجال</div>
        {cards_html}
      </div>

      <div style="background:#ffffff; border:1px solid #D8D2C2; border-radius:10px; padding:18px 20px; text-align:right;">
        <div style="color:#13294B; font-weight:bold; font-size:14px; margin-bottom:10px;">للقراءة هذا الأسبوع</div>
        <ul style="margin:0; padding-right:18px;">{reading_html}</ul>
        {book_html}
      </div>
    </div>

    <div style="text-align:center; padding:18px; font-size:11px; color:#5B6472; border-top:1px solid #D8D2C2; margin-top:10px;">
      نشرة داخلية سرية — مُعدّة عبر نظام رصد الرؤى الحوكمية والتدقيقية.<br>
      جميع الحقوق محفوظة © {datetime.now().year}
    </div>
  </div>
</body>
</html>"""


# ============================================================================
# 8. EMAIL DISPATCH
# ============================================================================
def send_email_report(html_content: str, report_filepath: str, issue_theme: str, issue_no: int):
    smtp_server = os.getenv("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(os.getenv("SMTP_PORT", 465))
    sender_email = os.getenv("SENDER_EMAIL")
    sender_password = os.getenv("SENDER_PASSWORD")
    recipient_raw = os.getenv("RECIPIENT_EMAILS", "")
    recipient_emails = [e.strip() for e in recipient_raw.split(",") if e.strip()]

    if not sender_email or not sender_password or not recipient_emails:
        log.error("CRITICAL: SMTP credentials missing (SENDER_EMAIL / SENDER_PASSWORD / RECIPIENT_EMAILS).")
        raise ValueError("Missing required SMTP environment variables.")

    msg = MIMEMultipart("mixed")
    subject_text = f"العدد {issue_no} | {issue_theme} — {datetime.now().strftime('%Y-%m-%d')}"
    # Explicit RFC 2047 header encoding: without this, some mail transfer
    # agents mangle non-ASCII Subject headers (a classic cause of Arabic
    # text turning into "?" specifically in the subject line).
    msg["Subject"] = Header(subject_text, "utf-8")
    msg["From"] = sender_email
    msg["To"] = ", ".join(recipient_emails)
    msg.attach(MIMEText(html_content, "html", "utf-8"))

    if os.path.exists(report_filepath):
        subtype = "pdf" if report_filepath.endswith(".pdf") else \
            "vnd.openxmlformats-officedocument.wordprocessingml.document"
        with open(report_filepath, "rb") as f:
            attachment = MIMEApplication(f.read(), _subtype=subtype)
            attachment.add_header("Content-Disposition", "attachment", filename=os.path.basename(report_filepath))
            msg.attach(attachment)
            log.info(f"Attached: {os.path.basename(report_filepath)}")
    else:
        log.warning(f"Report file missing: {report_filepath}")

    try:
        if smtp_port == 465:
            with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                server.login(sender_email, sender_password)
                server.sendmail(sender_email, recipient_emails, msg.as_string())
        else:
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.sendmail(sender_email, recipient_emails, msg.as_string())
        log.info(f"Newsletter dispatched to {len(recipient_emails)} recipient(s).")
    except Exception as e:
        log.error(f"Failed to dispatch email: {e}", exc_info=True)
        raise


# ============================================================================
# 9. PIPELINE
# ============================================================================
def run_pipeline():
    log.info("Step 0/4 — Verifying required fonts are installed...")
    verify_fonts_installed()

    log.info("Step 1/4 — Researching & drafting content with Gemini (grounded search)...")
    content, ok = generate_newsletter_content()
    if not ok:
        log.error(
            "Pipeline STOPPED before building or sending anything. "
            "See the FATAL log line above for the exact cause. No email was sent."
        )
        sys.exit(1)

    log.info("Step 2/4 — Building luxury Word document...")
    docx_filepath = build_word_document(content, reports_dir="reports")
    issue_no = int(re.search(r"Issue(\d+)", docx_filepath).group(1))

    log.info("Step 3/4 — Converting to PDF...")
    report_filepath = convert_docx_to_pdf(docx_filepath)

    log.info("Step 4/4 — Building HTML email and dispatching...")
    email_html = generate_email_html(content, issue_no)
    with open("email_preview.html", "w", encoding="utf-8") as f:
        f.write(email_html)

    send_email_report(email_html, report_filepath, content.get("issue_theme", "النشرة الأسبوعية"), issue_no)
    log.info("Pipeline complete.")


if __name__ == "__main__":
    run_pipeline()
