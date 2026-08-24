# -*- coding: utf-8 -*-
"""
النشرة الأسبوعية للحوكمة والتدقيق الداخلي — Governance & Audit Weekly

Pipeline:
  1) Tavily API للبحث عن آخر التطورات في المجالات الثمانية بشكل مستقل.
  2) نموذج LLM (GLM-4-Flash / DeepSeek) لتوليد JSON منظم بناءً على نتائج البحث.
  3) بناء ملف Word فاخر التصميم من الـ JSON.
  4) تحويل DOCX -> PDF.
  5) بناء إيميل HTML مطابق وإرساله.

المتغيرات البيئية المطلوبة:
  TAVILY_API_KEY
  LLM_PROVIDER (glm | deepseek | openai)
  LLM_API_KEY
  LLM_MODEL (e.g., glm-4-flash)
  SENDER_EMAIL, SENDER_PASSWORD, RECIPIENT_EMAILS

اختياري:
  SMTP_SERVER, SMTP_PORT    افتراضي smtp.gmail.com / 465

الحزم المطلوبة (requirements.txt): requests, python-docx,
python-dotenv (اختياري), docx2pdf (اختياري).

خطوط عربية على CI: يجب تثبيت fonts-hosny-amiri و fonts-crosextra-carlito
قبل تشغيل هذا السكربت.
"""

import os
import re
import sys
import json
import time
import logging
import smtplib
import subprocess
import requests
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
# 1. نظام الألوان والخطوط
# ============================================================================
NAVY_DARK    = "0B1F3A"
NAVY_PRIMARY = "13294B"
GOLD_ACCENT  = "C9A227"
GOLD_LIGHT   = "E8D9A0"
IVORY_BG     = "FBF9F4"
HAIRLINE     = "D8D2C2"
TEXT_DARK    = "1C2430"
TEXT_MUTED   = "5B6472"
TEXT_WHITE   = "FFFFFF"

RISK_COLORS = {
    "حرج": "8C1D1D",
    "عالٍ": "B5541A",
    "متوسط": "8A6D1A",
    "منخفض": "1E6B3E",
}

FONT_ARABIC = "Amiri"
FONT_LATIN = "Carlito"

ISSUE_NO_FILE = os.path.join(os.getcwd(), ".issue_number")
# ملاحظة: بيُكتب في مجلد التشغيل الحالي (جذر الـ repo عند تشغيله كـ
# "python scripts/main.py" من GitHub Actions) عمدًا، وليس بجانب هذا
# الملف، عشان يطابق أمر "git add .issue_number" في الـ workflow.

def verify_fonts_installed():
    for font_name, apt_pkg in [(FONT_ARABIC, "fonts-hosny-amiri"), (FONT_LATIN, "fonts-crosextra-carlito")]:
        try:
            result = subprocess.run(["fc-list", f":family={font_name}"], capture_output=True, text=True, timeout=10)
            if not result.stdout.strip():
                log.warning(f"⚠ خط '{font_name}' غير مثبت. ثبّته عبر: sudo apt-get install -y {apt_pkg}")
            else:
                log.info(f"Font check OK: '{font_name}' is installed.")
        except FileNotFoundError:
            log.warning("fc-list غير موجود — لا يمكن التحقق من الخطوط.")
        except Exception as e:
            log.warning(f"فشل التحقق من خط '{font_name}': {e}")

def next_issue_number() -> int:
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
# 2. مكتبة المصادر الموثوقة والمجالات
# ============================================================================
SOURCES = {
    "تدقيق داخلي وأداء (Internal / Performance Audit)": [
        ("IIA", "https://www.theiia.org/"),
        ("INTOSAI", "https://www.intosai.org/"),
        ("GAO US", "https://www.gao.gov"),
    ],
    "الحوكمة والقطاع العام (Governance / Public Sector)": [
        ("OECD", "https://www.oecd.org/governance/"),
        ("World Bank", "https://www.worldbank.org/"),
    ],
    "الرقابة الداخلية وإدارة المخاطر (Internal Control / Risk / COSO)": [
        ("COSO", "https://www.coso.org/"),
        ("IFAC", "https://www.ifac.org/"),
    ],
    "الموارد البشرية (HR)": [
        ("SHRM", "https://www.shrm.org/")
    ],
    "المحاسبة الإدارية (Management Accounting)": [
        ("IMA", "https://www.imanet.org/")
    ],
    "استشارات وأفضل الممارسات (Big Four / Strategy Insights)": [
        ("Deloitte", "https://www.deloitte.com/"),
        ("McKinsey", "https://www.mckinsey.com"),
        ("HBR", "https://hbr.org"),
    ],
}

# نطاقات مجرّدة (بدون https:// أو www.) لاستخدامها مباشرة مع بارامتر
# include_domains في Tavily. لازم تكون مطابقة فعليًا لأسماء النطاقات في
# SOURCES أعلاه — لو غيّرت مصدر هناك، حدّث هنا كمان.
TRUSTED_DOMAINS_BY_CATEGORY = {
    "تدقيق داخلي وأداء": ["theiia.org", "intosai.org", "gao.gov"],
    "الحوكمة والقطاع العام": ["oecd.org", "worldbank.org"],
    "الرقابة الداخلية وإدارة المخاطر": ["coso.org", "ifac.org"],
    "الموارد البشرية": ["shrm.org"],
    "المحاسبة الإدارية": ["imanet.org"],
    "استشارات وأفضل الممارسات": ["deloitte.com", "mckinsey.com", "hbr.org"],
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

# ربط كل مجال من المجالات الثمانية بفئة (أو أكثر) من TRUSTED_DOMAINS_BY_CATEGORY
# — هذا هو الربط الفعلي المفقود سابقًا بين قائمة المصادر الموثوقة واستعلامات البحث.
DOMAIN_TO_SOURCE_CATEGORIES = {
    DOMAINS[0]: ["تدقيق داخلي وأداء"],                                    # التدقيق الداخلي
    DOMAINS[1]: ["تدقيق داخلي وأداء"],                                    # تدقيق الأداء
    DOMAINS[2]: ["الحوكمة والقطاع العام"],                                # الحوكمة
    DOMAINS[3]: ["الرقابة الداخلية وإدارة المخاطر"],                       # الرقابة الداخلية
    DOMAINS[4]: ["الرقابة الداخلية وإدارة المخاطر"],                       # إدارة المخاطر
    DOMAINS[5]: ["الموارد البشرية"],                                      # الموارد البشرية
    DOMAINS[6]: ["المحاسبة الإدارية"],                                    # المحاسبة الإدارية
    DOMAINS[7]: ["الحوكمة والقطاع العام", "استشارات وأفضل الممارسات"],     # إصلاح القطاع العام
}

# نطاقات مخصصة لقسم "للقراءة هذا الأسبوع" — تُستخدم في بحث منفصل لضمان إن
# further_reading مبني على نتائج بحث حقيقية من Deloitte/McKinsey/HBR بدل ما
# يكون توليدًا حرًا من الموديل (كان السبب الرئيسي لضعف هذا القسم سابقًا).
FURTHER_READING_DOMAINS = TRUSTED_DOMAINS_BY_CATEGORY["استشارات وأفضل الممارسات"]

FLASHBACK_TOPICS = [
    "إطار COSO للرقابة الداخلية (COSO Internal Control Framework)",
    "إدارة الموارد البشرية الاستراتيجية (Strategic HR Management)",
    "المراقبة التسييرية (Management Control)",
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
# 3. توليد المحتوى — Tavily Search + OpenAI-Compatible LLMs (GLM/DeepSeek)
# ============================================================================

def _request_with_retries(method: str, url: str, max_retries: int = 3, backoff_base: float = 2.0, **kwargs):
    """
    غلاف صغير حول requests يعيد المحاولة مع backoff أسّي عند أخطاء الشبكة
    المؤقتة أو أكواد 429/5xx، بدل ما يفشل الـ pipeline كله من أول عثرة اتصال.
    """
    last_exc = None
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.request(method, url, **kwargs)
            if resp.status_code == 429 or resp.status_code >= 500:
                log.warning(f"محاولة {attempt}/{max_retries}: كود {resp.status_code} من {url} — إعادة محاولة...")
                last_exc = requests.HTTPError(f"HTTP {resp.status_code}")
                if attempt < max_retries:
                    time.sleep(backoff_base ** attempt)
                    continue
                resp.raise_for_status()
            return resp
        except (requests.ConnectionError, requests.Timeout) as e:
            last_exc = e
            log.warning(f"محاولة {attempt}/{max_retries} فشلت لـ {url}: {e}")
            if attempt < max_retries:
                time.sleep(backoff_base ** attempt)
    if last_exc:
        raise last_exc
    raise RuntimeError(f"فشل الطلب إلى {url} لأسباب غير معروفة.")


def _flatten_trusted_domains(categories: List[str]) -> List[str]:
    seen = []
    for cat in categories:
        for d in TRUSTED_DOMAINS_BY_CATEGORY.get(cat, []):
            if d not in seen:
                seen.append(d)
    return seen


def _tavily_search_call(api_key: str, query: str, include_domains: Optional[List[str]] = None,
                         max_results: int = 4, topic: str = "general") -> dict:
    payload = {
        "api_key": api_key,
        "query": query,
        "search_depth": "advanced",
        "topic": topic,          # "general" (وليس "news") لأن المصادر المؤسسية
                                  # (IIA, COSO, IFAC...) غالبًا غير مصنّفة كأخبار،
                                  # وفهرس الأخبار في Tavily كان يتجاهلها فعليًا.
        "time_range": "month",   # يعوّض عن topic="news" في تحديد الحداثة
        "max_results": max_results,
        "include_answer": True,
    }
    if include_domains:
        payload["include_domains"] = include_domains
    url = "https://api.tavily.com/search"
    response = _request_with_retries("POST", url, json=payload, timeout=30)
    response.raise_for_status()
    return response.json()


def _format_search_results(data: dict, header: str) -> str:
    ctx = f"### {header}:\n"
    if data.get("answer"):
        ctx += f"ملخص عام: {data['answer']}\n"
    for res in data.get("results", []):
        title = res.get("title", "")
        res_url = res.get("url", "")
        content = res.get("content", "")[:500]  # تقليل الحجم لتوفير الـ Tokens
        ctx += f"- العنوان: {title}\n  الرابط: {res_url}\n  المحتوى: {content}\n"
    return ctx


def search_with_tavily(domains: List[str]) -> str:
    """
    يستخدم Tavily API للبحث الفعلي عن أحدث التطورات لكل مجال، مقيّدًا بالنطاقات
    الموثوقة المحددة في TRUSTED_DOMAINS_BY_CATEGORY. لو البحث المقيّد رجع بدون
    نتائج (المصادر المؤسسية أحيانًا محدودة الفهرسة لموضوع بعينه هذا الشهر)،
    يتم عمل بحث احتياطي أوسع مع تحذير واضح في الـ log، بدل ما يفشل المجال بالكامل.
    """
    api_key = os.getenv("TAVILY_API_KEY")
    if not api_key:
        log.error("FATAL: TAVILY_API_KEY غير موجود في البيئة.")
        return ""

    all_context = []
    any_success = False

    for domain in domains:
        categories = DOMAIN_TO_SOURCE_CATEGORIES.get(domain, [])
        include_domains = _flatten_trusted_domains(categories)
        query = f"أحدث التطورات والمستجدات في {domain}"
        try:
            log.info(f"Tavily searching (مقيّد بالمصادر الموثوقة {include_domains}): {domain}...")
            data = _tavily_search_call(api_key, query, include_domains=include_domains, max_results=4)

            if not data.get("results"):
                log.warning(
                    f"لا نتائج من المصادر الموثوقة لمجال '{domain}' خلال آخر شهر. "
                    f"إجراء بحث احتياطي أوسع (غير مقيّد بالنطاقات المتفق عليها)..."
                )
                data = _tavily_search_call(api_key, query, include_domains=None, max_results=4)
                header = f"نتائج البحث لـ {domain} (تنبيه: بحث احتياطي عام، ليس من المصادر المتفق عليها)"
            else:
                header = f"نتائج البحث لـ {domain} (من المصادر الموثوقة: {', '.join(include_domains)})"

            all_context.append(_format_search_results(data, header))
            if data.get("results"):
                any_success = True
        except Exception as e:
            log.error(f"فشل البحث في Tavily لمجال {domain}: {e}")
            all_context.append(f"### نتائج البحث لـ {domain}:\n(تعذر جلب النتائج)\n")

    # بحث منفصل مخصص لقسم "للقراءة هذا الأسبوع" من مصادر Deloitte/McKinsey/HBR
    # تحديدًا، بدل ما يعتمد الموديل على توليد حر بلا سياق بحثي حقيقي.
    try:
        log.info(f"Tavily searching for further-reading sources: {FURTHER_READING_DOMAINS}...")
        fr_data = _tavily_search_call(
            api_key,
            "أحدث المقالات والتقارير في التدقيق الداخلي والحوكمة وإدارة المخاطر",
            include_domains=FURTHER_READING_DOMAINS,
            max_results=4,
        )
        if fr_data.get("results"):
            all_context.append(_format_search_results(
                fr_data, f"مصادر مخصصة لقسم further_reading (من: {', '.join(FURTHER_READING_DOMAINS)})"
            ))
            any_success = True
        else:
            log.warning("لا نتائج حديثة من مصادر further_reading الموثوقة هذا الشهر.")
    except Exception as e:
        log.error(f"فشل بحث further_reading: {e}")

    if not any_success:
        log.error("FATAL: فشل البحث في كل المجالات (المقيّدة والاحتياطية) — لا يوجد سياق حقيقي لإرساله للـ LLM.")
        return ""

    return "\n\n".join(all_context)



def generate_with_llm(prompt: str) -> Optional[str]:
    """يستدعي نموذج LLM متوافق مع OpenAI (GLM-4-Flash / DeepSeek) لتوليد JSON."""
    provider = os.getenv("LLM_PROVIDER", "glm").lower()
    api_key = os.getenv("LLM_API_KEY")
    model = os.getenv("LLM_MODEL", "glm-4-flash")

    if not api_key:
        log.error("FATAL: LLM_API_KEY غير موجود في البيئة.")
        return None

    # تحديد Base URL حسب المزود
    if provider == "glm":
        base_url = "https://open.bigmodel.cn/api/paas/v4"
    elif provider == "deepseek":
        base_url = "https://api.deepseek.com/v1"
    elif provider == "openai":
        base_url = "https://api.openai.com/v1"
    else:
        log.error(f"LLM_PROVIDER غير مدعوم: {provider}")
        return None

    url = f"{base_url}/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}

    # max_tokens صريح ومرتفع: بدون تحديده بعض المزودين (خصوصًا GLM) بيستخدموا حد
    # افتراضي صغير نسبيًا، فالرد بيتقطع في نص الـ JSON (عادة قبل آخر حقلين
    # further_reading و book_recommendation لأنهم آخر شيء في الهيكل). قابل
    # للتعديل عبر متغير بيئي LLM_MAX_TOKENS لو احتجت رفعه أكتر.
    max_tokens = int(os.getenv("LLM_MAX_TOKENS", "6000"))

    # إجبار النموذج على إرجاع JSON صرف
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.4,
        "max_tokens": max_tokens,
        "response_format": {"type": "json_object"}
    }

    response = None  # مُعرّف مسبقًا حتى لا ينفجر الـ except لو فشل الطلب قبل استلام أي رد
    try:
        log.info(f"Generating content with {provider}/{model} (max_tokens={max_tokens})...")
        response = _request_with_retries("POST", url, json=payload, headers=headers, timeout=90)
        response.raise_for_status()
        resp_json = response.json()
        choice = resp_json["choices"][0]
        finish_reason = choice.get("finish_reason")
        content_text = choice["message"]["content"]
        if finish_reason == "length":
            log.warning(
                f"تحذير: الرد اتقطع بسبب حد الطول (finish_reason=length) رغم "
                f"max_tokens={max_tokens}. طول الرد المستلم: {len(content_text)} حرف. "
                f"جرّب رفع LLM_MAX_TOKENS في الـ workflow."
            )
        else:
            log.info(f"LLM finished normally (finish_reason={finish_reason}), response length: {len(content_text)} chars.")
        return content_text
    except Exception as e:
        log.error(f"FATAL: فشل استدعاء نموذج التوليد ({type(e).__name__}): {e}")
        if response is not None:
            log.error(f"Response Body: {response.text[:2000]}")
        return None


def build_generation_prompt(search_context: str) -> str:
    flashback_topic = pick_flashback_topic()

    return f"""أنت محرر تنفيذي متخصص في التدقيق الداخلي والحوكمة وإدارة المخاطر، تُعِدّ نشرة أسبوعية
احترافية لجمهور من كبار المدققين والمسؤولين الحكوميين والماليين التنفيذيين. مستوى الجودة
المطلوب أعلى من نشرات Deloitte وMcKinsey وHBR: مكثف، عملي، بلا حشو، بقيمة مضافة حقيقية.

تم جلب نتائج بحث فعلية وحديثة من الإنترنت، بعضها من مصادر مهنية موثوقة محددة
مسبقًا (IIA, INTOSAI, GAO, OECD, World Bank, COSO, IFAC, SHRM, IMA, Deloitte,
McKinsey, HBR) وبعضها الآخر بحث احتياطي عام (مُعلَّم بوضوح في العنوان لو كان كذلك).
مهمتك هي قراءتها وتنظيمها في النشرة الأسبوعية. لا تقم باختلاق أي أخبار أو روابط
غير موجودة في الـ Context المرفق، ولا تستبدل رابطًا حقيقيًا من الـ Context برابط
من معرفتك العامة.

=== نتائج البحث من الإنترنت (Search Context) ===
{search_context}
================================================

غطِّ المجالات التالية واختر الأهم من نتائج البحث الموسومة بها (فضّل دائمًا
النتائج القادمة من "المصادر الموثوقة" على النتائج القادمة من "بحث احتياطي عام"
لو كان الاثنان متاحين لنفس المجال):
{chr(10).join(f'- {d}' for d in DOMAINS)}

لقسم further_reading تحديدًا: استخدم فقط النتائج الموجودة تحت القسم المعنون
"مصادر مخصصة لقسم further_reading" في الـ Context أعلاه (من Deloitte/McKinsey/HBR).
لو هذا القسم فاضي أو غير موجود في الـ Context، أعد further_reading كقائمة فارغة []
بدل اختلاق عناصر غير موجودة فعليًا في نتائج البحث.

موضوع "ومضة إدارية" (Flashback) لهذا الأسبوع هو: {flashback_topic}
اشرحه بإيجاز عملي بناءً على معرفتك العامة.

أعد النتيجة ككائن JSON صِرف واحد فقط (بدون أي نص قبله أو بعده، بدون Markdown fences)
مطابقًا تمامًا لهذا الهيكل:

{{
  "issue_theme": "عنوان جذاب لموضوع العدد هذا الأسبوع (جملة قصيرة)",
  "editor_note": "فقرة افتتاحية قصيرة (3-4 أسطر) بأسلوب تنفيذي راقٍ",
  "quick_wins": [
    {{"tip_ar": "نصيحة عملية قابلة للتطبيق فورًا بجملة أو جملتين مع المصطلح الإنجليزي بين قوسين"}}
    // 3 عناصر بالضبط
  ],
  "flashback": {{
    "topic": "اسم الموضوع الإداري الكلاسيكي المرسل إليك بالضبط",
    "content_ar": "شرح مكثف وعملي (فقرة واحدة، 4-6 أسطر) لماذا هذا المفهوم لا يزال أساسيًا اليوم",
    "key_points": ["نقطة تطبيقية 1", "نقطة تطبيقية 2", "نقطة تطبيقية 3"]
  }},
  "domain_updates": [
    {{
      "domain": "اسم المجال كما ورد في القائمة",
      "title": "عنوان التحديث/الخبر من نتائج البحث",
      "category": "تصنيف قصير",
      "risk_level": "حرج | عالٍ | متوسط | منخفض",
      "summary": "ملخص التطور الأحدث في هذا المجال (2-3 جمل) مع ترجمة أي مصطلح تقني إلى الإنجليزية بين قوسين",
      "why_it_matters": "الأثر المباشر على المؤسسة (جملتان)",
      "recommended_actions": "إجراءات عملية وقابلة للتنفيذ فورًا (جملتان، صيغة أفعال أمر)",
      "source_name": "اسم المصدر الفعلي من نتائج البحث",
      "source_url": "رابط حقيقي وقابل للفتح من نتائج البحث"
    }}
    // عنصر واحد لكل مجال من المجالات الثمانية المرسلة، بحد أقصى 6 عناصر (اختر الأهم)
  ],
  "further_reading": [
    {{"title": "عنوان المقال/التقرير", "source_name": "اسم الجهة", "url": "رابط حقيقي", "type": "مقال | تقرير | دراسة"}}
    // 4 عناصر
  ],
  "book_recommendation": {{
    "title": "عنوان كتاب معروف وذو صلة بأحد المجالات",
    "author": "اسم المؤلف",
    "why_ar": "لماذا يستحق القراءة هذا الأسبوع تحديدًا (جملتان)"
  }}
}}

قواعد إلزامية:
- اكتب المحتوى بالعربية الفصحى الاحترافية، وضع كل مصطلح تقني بالإنجليزية بين قوسين.
- استخدم فقط الأخبار والروابط الموجودة في نتائج البحث المرفقة (Search Context).
- اجعل اللهجة عملية ومباشرة (hands-on)، بلا حشو إنشائي.
"""


def _extract_json(text: str) -> Optional[dict]:
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


def _validate_content_shape(data: dict) -> bool:
    """
    تحقق دفاعي بسيط: تأكد إن المفاتيح الأساسية اللي باقي السكربت (DOCX/PDF/HTML)
    يعتمد عليها موجودة، حتى لو القيم فاضية، بدل ما ينهار لاحقًا بـ KeyError غامض.
    """
    required_keys = [
        "issue_theme", "editor_note", "quick_wins", "flashback",
        "domain_updates", "further_reading", "book_recommendation",
    ]
    missing = [k for k in required_keys if k not in data]
    if missing:
        log.error(f"استجابة LLM ناقصة المفاتيح التالية: {missing}")
        return False
    if not isinstance(data.get("domain_updates"), list) or not data["domain_updates"]:
        log.error("حقل domain_updates فاضي أو ليس قائمة — لا فائدة من نشرة بدون تحديثات.")
        return False
    return True


def generate_newsletter_content() -> Tuple[Optional[Dict[str, Any]], bool]:
    """يبني المحتوى بالكامل: بحث Tavily ثم توليد LLM."""

    # 1. البحث الفعلي
    search_context = search_with_tavily(DOMAINS)
    if not search_context:
        log.error("Pipeline STOPPED: تعذر جلب نتائج البحث من Tavily.")
        return None, False

    # 2. توليد JSON
    prompt = build_generation_prompt(search_context)
    raw_response = generate_with_llm(prompt)

    if not raw_response:
        return None, False

    data = _extract_json(raw_response)
    if not data:
        log.error("FATAL: تعذر تحليل استجابة الـ LLM كـ JSON.")
        log.error(f"Raw LLM response (first 2000 chars): {raw_response[:2000]}")
        return None, False

    data.setdefault("flashback", {}).setdefault("topic", pick_flashback_topic())

    if not _validate_content_shape(data):
        log.error(
            f"Raw LLM response length: {len(raw_response)} chars. "
            f"Last 500 chars (لمعرفة هل الرد اتقطع في النهاية): ...{raw_response[-500:]}"
        )
        return None, False

    log.info(f"Content generated successfully: {len(data.get('domain_updates', []))} domain updates.")
    return data, True

# ============================================================================
# 4. أدوات DOCX
# ============================================================================
def set_rtl(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement("w:bidi")
    pPr.append(bidi)

def style_run(run, size=11, bold=False, color=TEXT_DARK, italic=False):
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
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), FONT_LATIN); rFonts.set(qn("w:hAnsi"), FONT_LATIN)
    rFonts.set(qn("w:cs"), FONT_ARABIC); rFonts.set(qn("w:eastAsia"), FONT_ARABIC)
    rPr.append(rFonts)
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(int(size * 2))); rPr.append(sz)
    if bold: rPr.append(OxmlElement("w:b"))
    if underline:
        u = OxmlElement("w:u"); u.set(qn("w:val"), "single"); rPr.append(u)
    color_el = OxmlElement("w:color"); color_el.set(qn("w:val"), color); rPr.append(color_el)
    new_run.append(rPr)
    t = OxmlElement("w:t"); t.text = text; new_run.append(t)
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
    bottom.set(qn("w:val"), "single"); bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "1"); bottom.set(qn("w:color"), GOLD_ACCENT)
    pBdr.append(bottom); pPr.append(pBdr)
    return p

# ============================================================================
# 5. بناء DOCX
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
    style_run(p_kicker.add_run(f"العدد {issue_no}  •  {datetime.now().strftime('%Y-%m-%d')}"), size=11, color=GOLD_LIGHT, bold=True)

    p_title = cell.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_rtl(p_title)
    p_title.paragraph_format.space_before = Pt(14)
    style_run(p_title.add_run("النشرة الأسبوعية للحوكمة والتدقيق الداخلي"), size=26, bold=True, color=TEXT_WHITE)

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
    top.set(qn("w:val"), "single"); top.set(qn("w:sz"), "6")
    top.set(qn("w:space"), "1"); top.set(qn("w:color"), GOLD_ACCENT)
    pBdr.append(top); pPr.append(pBdr)

    p_sub = cell.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_rtl(p_sub)
    p_sub.paragraph_format.space_before = Pt(10)
    style_run(p_sub.add_run("تدقيق داخلي • تدقيق أداء • حوكمة • رقابة داخلية • إدارة مخاطر • موارد بشرية • محاسبة إدارية • إصلاح القطاع العام"), size=10, color=GOLD_LIGHT)
    doc.add_paragraph()

def add_editor_note(doc, content: Dict[str, Any]):
    p = doc.add_paragraph(); set_rtl(p)
    p.paragraph_format.space_after = Pt(4)
    style_run(p.add_run("كلمة التحرير"), size=13, bold=True, color=NAVY_PRIMARY)

    p2 = doc.add_paragraph(); set_rtl(p2)
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
        p = top_cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER; set_rtl(p)
        style_run(p.add_run(h), size=10, bold=True, color=TEXT_WHITE)

        bottom_cell = tbl.cell(1, i)
        set_cell_background(bottom_cell, IVORY_BG)
        set_cell_margins(bottom_cell, top=140, bottom=140)
        p2 = bottom_cell.paragraphs[0]; p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        style_run(p2.add_run(v), size=18, bold=True, color=(GOLD_ACCENT if i != 1 else RISK_COLORS["حرج"]))
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

def add_quick_wins(doc, quick_wins: List[Dict[str, str]]):
    if not quick_wins: return
    header = doc.add_paragraph(); set_rtl(header)
    style_run(header.add_run("انتصارات سريعة (Quick Wins)"), size=13, bold=True, color=NAVY_PRIMARY)

    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    set_cell_background(cell, IVORY_BG)
    set_cell_margins(cell, top=120, bottom=120, left=200, right=200)
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="none"/><w:bottom w:val="none"/><w:left w:val="none"/><w:right w:val="single" w:sz="20" w:space="0" w:color="{GOLD_ACCENT}"/></w:tcBorders>')
    tcPr.append(borders)
    first = True
    for item in quick_wins:
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False; set_rtl(p)
        p.paragraph_format.space_after = Pt(6)
        style_run(p.add_run("— "), size=11, bold=True, color=GOLD_ACCENT)
        style_run(p.add_run(item.get("tip_ar", "")), size=11, color=TEXT_DARK)
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

def add_flashback(doc, flashback: Dict[str, Any]):
    if not flashback or not flashback.get("content_ar"): return
    header = doc.add_paragraph(); set_rtl(header)
    style_run(header.add_run(f"ومضة إدارية — {flashback.get('topic', '')}"), size=13, bold=True, color=NAVY_PRIMARY)

    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    set_cell_background(cell, "FFFFFF")
    set_cell_margins(cell, top=120, bottom=120, left=200, right=200)
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="4" w:space="0" w:color="{HAIRLINE}"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="{HAIRLINE}"/><w:left w:val="single" w:sz="4" w:space="0" w:color="{HAIRLINE}"/><w:right w:val="single" w:sz="4" w:space="0" w:color="{HAIRLINE}"/></w:tcBorders>')
    tcPr.append(borders)

    p = cell.paragraphs[0]; set_rtl(p); p.paragraph_format.line_spacing = 1.3
    style_run(p.add_run(flashback.get("content_ar", "")), size=11, color=TEXT_DARK)

    for kp in flashback.get("key_points", []):
        pk = cell.add_paragraph(); set_rtl(pk)
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

    p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT; set_rtl(p)
    style_run(p.add_run(f"{title}  "), size=13.5, bold=True, color=TEXT_WHITE)
    risk_color = RISK_COLORS.get(risk, "8A6D1A")
    tag = p.add_run(f" [{category} | {risk}]")
    style_run(tag, size=10, bold=True, color=GOLD_LIGHT)

def add_domain_card(doc, block: Dict[str, Any]):
    create_section_header(doc, block.get("domain", block.get("title", "تحديث")), risk=block.get("risk_level", "متوسط"), category=block.get("category", "عام"))
    if block.get("title"):
        p_title = doc.add_paragraph(); set_rtl(p_title)
        p_title.paragraph_format.space_before = Pt(4)
        style_run(p_title.add_run(block["title"]), size=12, bold=True, color=NAVY_PRIMARY)

    p_body = doc.add_paragraph(); set_rtl(p_body)
    p_body.paragraph_format.line_spacing = 1.3
    p_body.paragraph_format.space_after = Pt(6)
    style_run(p_body.add_run(block.get("summary", "")), size=11, color=TEXT_DARK)

    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    set_cell_background(cell, IVORY_BG)
    set_cell_margins(cell, top=110, bottom=110, left=180, right=180)
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="single" w:sz="20" w:space="0" w:color="{RISK_COLORS.get(block.get("risk_level"), GOLD_ACCENT)}"/></w:tcBorders>')
    tcPr.append(borders)

    p1 = cell.paragraphs[0]; set_rtl(p1)
    style_run(p1.add_run("الأثر (Why It Matters): "), size=10.5, bold=True, color=NAVY_PRIMARY)
    style_run(p1.add_run(block.get("why_it_matters", "")), size=10.5, color=TEXT_DARK)

    p2 = cell.add_paragraph(); set_rtl(p2)
    p2.paragraph_format.space_before = Pt(4)
    style_run(p2.add_run("إجراءات موصى بها: "), size=10.5, bold=True, color=NAVY_PRIMARY)
    style_run(p2.add_run(block.get("recommended_actions", "")), size=10.5, color=TEXT_DARK)

    if block.get("source_url"):
        p3 = cell.add_paragraph(); set_rtl(p3)
        p3.paragraph_format.space_before = Pt(4)
        style_run(p3.add_run("المصدر: "), size=9.5, bold=True, color=TEXT_MUTED)
        add_hyperlink(p3, block.get("source_name", block["source_url"]), block["source_url"], size=9.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(10)

def add_further_reading(doc, items: List[Dict[str, str]], book: Dict[str, str]):
    if not items and not (book and book.get("title")): return
    header = doc.add_paragraph(); set_rtl(header)
    style_run(header.add_run("للقراءة هذا الأسبوع (Further Reading)"), size=13, bold=True, color=NAVY_PRIMARY)

    for it in items:
        p = doc.add_paragraph(); set_rtl(p)
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
        p = doc.add_paragraph(); set_rtl(p)
        style_run(p.add_run("كتاب الأسبوع: "), size=11, bold=True, color=NAVY_PRIMARY)
        style_run(p.add_run(f"{book['title']} — {book.get('author', '')}"), size=11, bold=True, color=TEXT_DARK)
        p2 = doc.add_paragraph(); set_rtl(p2)
        style_run(p2.add_run(book.get("why_ar", "")), size=10.5, italic=True, color=TEXT_MUTED)
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

def build_word_document(content: Dict[str, Any], reports_dir: str = "reports") -> str:
    os.makedirs(reports_dir, exist_ok=True)
    issue_no = next_issue_number()
    filename = os.path.join(reports_dir, f"Governance_Weekly_Brief_{datetime.now().strftime('%Y-%m-%d')}_Issue{issue_no}.docx")
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.7); section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.8); section.right_margin = Inches(0.8)
        footer = section.footer
        f_p = footer.paragraphs[0]; f_p.alignment = WD_ALIGN_PARAGRAPH.CENTER; set_rtl(f_p)
        style_run(f_p.add_run("النشرة الأسبوعية للحوكمة والتدقيق — سري وللاستخدام الداخلي فقط  |  "), size=8, color=TEXT_MUTED)
        fldChar1 = OxmlElement("w:fldChar"); fldChar1.set(qn("w:fldCharType"), "begin")
        instrText = OxmlElement("w:instrText"); instrText.set(qn("xml:space"), "preserve"); instrText.text = "PAGE"
        fldChar2 = OxmlElement("w:fldChar"); fldChar2.set(qn("w:fldCharType"), "end")
        run = f_p.add_run(); run._r.append(fldChar1); run._r.append(instrText); run._r.append(fldChar2)
        style_run(run, size=8, color=TEXT_MUTED)

    add_cover_page(doc, content, issue_no)
    add_editor_note(doc, content)
    add_executive_dashboard(doc, content.get("domain_updates", []))
    add_quick_wins(doc, content.get("quick_wins", []))
    add_flashback(doc, content.get("flashback", {}))

    updates_header = doc.add_paragraph(); set_rtl(updates_header)
    style_run(updates_header.add_run("أبرز التطورات حسب المجال"), size=13, bold=True, color=NAVY_PRIMARY)
    add_gold_rule(doc, space_before=2, space_after=8)

    for block in content.get("domain_updates", []):
        add_domain_card(doc, block)

    add_further_reading(doc, content.get("further_reading", []), content.get("book_recommendation", {}))
    doc.save(filename)
    log.info(f"Generated Word document: {filename}")
    return filename

# ============================================================================
# 6. تحويل PDF
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
        subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", docx_path, "--outdir", output_dir], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        log.info(f"Converted to PDF via LibreOffice: {pdf_path}")
        return pdf_path
    except Exception as e2:
        log.warning(f"LibreOffice conversion failed: {e2}. Falling back to DOCX attachment.")
    return docx_path

# ============================================================================
# 7. إيميل HTML
# ============================================================================
def _html_escape(s: str) -> str:
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def generate_email_html(content: Dict[str, Any], issue_no: int) -> str:
    quick_wins_html = "".join(f'<li style="margin-bottom:8px; color:#1C2430;">{_html_escape(q.get("tip_ar",""))}</li>' for q in content.get("quick_wins", []))
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
            source_html = f'<a href="{block["source_url"]}" style="color:#13294B; font-size:12px; text-decoration:underline;">المصدر: {_html_escape(block.get("source_name", ""))}</a>'
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
        reading_html += f'<li style="margin-bottom:8px; color:#1C2430; font-size:13px;"><span style="color:#5B6472;">{_html_escape(it.get("type","مقال"))} — </span>{link}<span style="color:#5B6472; font-size:12px;"> ({_html_escape(it.get("source_name",""))})</span></li>'

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
# 8. إرسال الإيميل
# ============================================================================
def send_email_report(html_content: str, report_filepath: str, issue_theme: str, issue_no: int):
    smtp_server = os.getenv("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(os.getenv("SMTP_PORT", 465))
    sender_email = os.getenv("SENDER_EMAIL")
    sender_password = os.getenv("SENDER_PASSWORD")
    recipient_raw = os.getenv("RECIPIENT_EMAILS", "")
    recipient_emails = [e.strip() for e in recipient_raw.split(",") if e.strip()]

    if not sender_email or not sender_password or not recipient_emails:
        log.error("CRITICAL: بيانات SMTP ناقصة (SENDER_EMAIL / SENDER_PASSWORD / RECIPIENT_EMAILS).")
        raise ValueError("Missing required SMTP environment variables.")

    msg = MIMEMultipart("mixed")
    subject_text = f"العدد {issue_no} | {issue_theme} — {datetime.now().strftime('%Y-%m-%d')}"
    msg["Subject"] = Header(subject_text, "utf-8")
    msg["From"] = sender_email
    msg["To"] = ", ".join(recipient_emails)
    msg.attach(MIMEText(html_content, "html", "utf-8"))

    if os.path.exists(report_filepath):
        subtype = "pdf" if report_filepath.endswith(".pdf") else "vnd.openxmlformats-officedocument.wordprocessingml.document"
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
# 9. الـ Pipeline
# ============================================================================
def run_pipeline():
    log.info("Step 0/4 — Verifying required fonts are installed...")
    verify_fonts_installed()

    log.info("Step 1/4 — Researching with Tavily & drafting content with LLM...")
    content, ok = generate_newsletter_content()
    if not ok:
        log.error("Pipeline STOPPED before building or sending anything. No email was sent.")
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
