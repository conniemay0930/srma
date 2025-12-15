# app.py
# =========================================================
# 一句話帶你完成 MA（BYOK / Traditional Chinese）
# Author: Ya Hsin Yao
#
# 免責聲明：本工具僅供學術研究/教學用途，不構成醫療建議或法律意見；
# 使用者須自行驗證所有結果、數值、引用與全文內容；請勿上傳可識別之病人資訊。
#
# 校內資源/授權提醒（重要）：
# - 若文章來自校內訂閱（付費期刊/EZproxy/館藏系統），請勿將受版權保護之全文
#   上傳至任何第三方服務或公開部署之網站（包含本 app 的雲端部署）。
# - 請遵守圖書館授權條款：避免大量下載/自動化批次擷取、避免共享全文給未授權者。
# - 若不確定是否可上傳：建議改用「本機版」或僅上傳你有權分享的開放取用全文（OA/PMC）。
#
# Privacy notice (BYOK):
# - Key only used for this session; do not use on untrusted deployments;
# - do not upload identifiable patient info.
# =========================================================

from __future__ import annotations

import io
import re
import math
import json
import html
import time
from dataclasses import dataclass
from typing import Dict, List, Tuple, Optional, Any

import requests
import pandas as pd
import streamlit as st
import xml.etree.ElementTree as ET

# Optional: PDF text extraction
try:
    from PyPDF2 import PdfReader  # type: ignore
    HAS_PYPDF2 = True
except Exception:
    HAS_PYPDF2 = False

# Optional: Plotly for forest plot
try:
    import plotly.graph_objects as go  # type: ignore
    HAS_PLOTLY = True
except Exception:
    HAS_PLOTLY = False

# Optional: Matplotlib fallback
try:
    import matplotlib.pyplot as plt  # type: ignore
    HAS_MPL = True
except Exception:
    HAS_MPL = False

# Optional: Word export
try:
    from docx import Document  # type: ignore
    from docx.shared import Pt  # type: ignore
    HAS_DOCX = True
except Exception:
    HAS_DOCX = False


# -------------------- Page config --------------------
st.set_page_config(page_title="一句話帶你完成 MA（繁體中文）", layout="wide")

# -------------------- Styles --------------------
CSS = """
<style>
:root{
  --bg: #ffffff;
  --muted: #6b7280;
  --line: #e5e7eb;
  --soft: #f7f7fb;
  --warn-bg:#fff7ed;
  --warn-line:#f59e0b;
  --ok-bg:#ecfdf5;
  --ok-line:#10b981;
  --bad-bg:#fef2f2;
  --bad-line:#ef4444;
}
.small { font-size: 0.92rem; color: var(--muted); }
.muted { color: var(--muted); }
.wrap { white-space: normal; }
.card { border: 1px solid var(--line); border-radius: 16px; padding: 0.95rem 1.05rem; background: var(--bg); margin-bottom: 0.9rem; box-shadow: 0 1px 0 rgba(0,0,0,0.03); }
.notice { border-left: 5px solid var(--warn-line); background: var(--warn-bg); padding: 0.95rem 1.05rem; border-radius: 14px; }
.kpi { border: 1px solid var(--line); border-radius: 16px; padding: 0.8rem 1rem; background: var(--soft); }
.kpi .label { font-size: 0.84rem; color: var(--muted); }
.kpi .value { font-size: 1.35rem; font-weight: 800; color: #111827; }
.badge { display:inline-block; padding:0.18rem 0.6rem; border-radius:999px; font-size:0.78rem; margin-right:0.35rem; border:1px solid rgba(0,0,0,0.06); background:#f3f4f6; }
.badge-ok { background: var(--ok-bg); border-color: rgba(16,185,129,0.25); color:#065f46; }
.badge-warn { background: #fef3c7; border-color: rgba(245,158,11,0.25); color:#92400e; }
.badge-bad { background: var(--bad-bg); border-color: rgba(239,68,68,0.25); color:#991b1b; }
.hr { border:none; border-top:1px solid #eef2f7; margin: 0.9rem 0; }
a { text-decoration: none; }
.red { color: #b91c1c; font-weight: 650; }
.green { color: #065f46; font-weight: 650; }
.flow { display:grid; grid-template-columns: 1fr; gap: 10px; }
.flow-row{ display:grid; grid-template-columns: 1fr; gap: 10px; }
.flow-box{ border:1px solid var(--line); border-radius: 14px; padding: 10px 12px; background: #fff; }
.flow-box .t{ font-weight: 800; margin-bottom: 2px; }
.flow-box .n{ color: var(--muted); font-size: 0.92rem; }
.flow-arrow{ text-align:center; color: var(--muted); font-size: 1.1rem; }
@media (min-width: 900px){ .flow-row{ grid-template-columns: 1fr 1fr; gap: 12px; } }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


# -------------------- Minimal i18n --------------------
I18N = {
    "zh-TW": {
        "app_title": "一句話帶你完成 MA",
        "author": "作者：Ya Hsin Yao",
        "lang_label": "語言",
        "tips_title": "小叮嚀／注意事項",
        "byok_title": "LLM（使用者自備 key）",
        "byok_toggle": "啟用 LLM（使用者自備 key）",
        "byok_notice": "Key only used for this session; do not use on untrusted deployments; do not upload identifiable patient info.",
        "byok_consent": "我了解並同意：不在不受信任的部署輸入 key、不上傳可識別病人資訊",
        "byok_clear": "Clear key",
        "settings": "設定",
        "links_title": "文獻與資源連結（可選）",
        "resolver": "OpenURL resolver base",
        "ezproxy": "EZproxy prefix",
        "search_settings": "檢索設定",
        "article_type": "文章類型 filter",
        "custom_filter": "自訂 filter（可留空）",
        "goal_mode": "研究目標取向",
        "question_notice": "輸入一句研究問題",
        "question_help": "例如：「FLACS 是否比傳統 phaco 好？」或「Diffractive vs nondiffractive EDOF IOL visual performance」。",
        "question_label": "Research question (one sentence)",
        "run": "開始（自動產出完整流程）",
        "tabs_overview": "總覽（PRISMA）",
        "tabs_step1": "Step 1 搜尋式（可手改）",
        "tabs_step2": "Step 2 可行性（SR/MA/NMA）",
        "tabs_step34": "Step 3+4 Records + 粗篩（合併）",
        "tabs_ft": "Step 4b Full text review",
        "tabs_extract": "Step 5 Data extraction（寬表）",
        "tabs_ma": "Step 6 MA + 森林圖",
        "tabs_rob2": "Step 6b ROB 2.0（含理由）",
        "tabs_ms": "Step 7 稿件草稿（分段呈現）",
        "tabs_diag": "Diagnostics",
        "pubmed_edit": "PubMed query（editable）",
        "pubmed_refetch": "以此搜尋式重新抓 PubMed",
        "pubmed_restore": "恢復為自動產生",
        "download_query": "下載搜尋式（txt）",
        "feas_title": "可行性掃描（既有 SR/MA/NMA）+ 綜合建議",
        "feas_optional": "（可選）BYOK：可行性報告",
        "records_none": "沒有抓到 records。",
        "ft_bulk_upload": "批次上傳 PDF（可選）",
        "ft_single_upload": "PDF 上傳（單篇，可選）",
        "ft_extract_text": "抽字（PDF→文字）",
        "ft_text_area": "Full-text text（可貼上；建議先 OCR 再貼/上傳）",
        "ft_ai_fill": "AI 讀全文 + 回填（對這篇）",
        "extract_schema": "Extraction schema（可自行規劃欄位；一行一欄）",
        "extract_quick_add": "快速新增一筆（一次輸入完再寫入寬表）",
        "extract_editor": "進階：一次編輯整張表（按儲存才 commit）",
        "extract_save": "儲存/commit 寬表修改",
        "ma_outcome_label": "Outcome label（手動輸入；用來篩選要做 MA 的列）",
        "ma_measure": "Effect measure（OR/RR/HR/MD/SMD）",
        "ma_run": "Run MA + 森林圖（按鈕執行）",
        "ms_generate": "（可選）BYOK：生成更完整稿件（請人工核對）",
        "export_docx": "匯出 Word（DOCX）",
    },
    "en": {
        "app_title": "From one question to Meta-analysis",
        "author": "Author: Ya Hsin Yao",
        "lang_label": "Language",
        "tips_title": "Notes / Warnings",
        "byok_title": "LLM (Bring Your Own Key)",
        "byok_toggle": "Enable LLM (BYOK)",
        "byok_notice": "Key only used for this session; do not use on untrusted deployments; do not upload identifiable patient info.",
        "byok_consent": "I understand and agree: do not use key on untrusted deployments; do not upload identifiable patient info",
        "byok_clear": "Clear key",
        "settings": "Settings",
        "links_title": "Institution links (optional)",
        "resolver": "OpenURL resolver base",
        "ezproxy": "EZproxy prefix",
        "search_settings": "Search settings",
        "article_type": "Article type filter",
        "custom_filter": "Custom filter (optional)",
        "goal_mode": "Goal mode",
        "question_notice": "Enter one research question",
        "question_help": "e.g., “FLACS vs conventional phaco” or “Diffractive vs nondiffractive EDOF IOL visual performance”.",
        "question_label": "Research question (one sentence)",
        "run": "Run (end-to-end pipeline)",
        "tabs_overview": "Overview (PRISMA)",
        "tabs_step1": "Step 1 Query (editable)",
        "tabs_step2": "Step 2 Feasibility (SR/MA/NMA)",
        "tabs_step34": "Step 3+4 Records + Screening",
        "tabs_ft": "Step 4b Full-text review",
        "tabs_extract": "Step 5 Data extraction",
        "tabs_ma": "Step 6 MA + Forest",
        "tabs_rob2": "Step 6b RoB 2.0",
        "tabs_ms": "Step 7 Manuscript draft",
        "tabs_diag": "Diagnostics",
        "pubmed_edit": "PubMed query (editable)",
        "pubmed_refetch": "Refetch PubMed using this query",
        "pubmed_restore": "Restore auto query",
        "download_query": "Download query (txt)",
        "feas_title": "Feasibility scan + recommendations",
        "feas_optional": "(Optional) BYOK: feasibility report",
        "records_none": "No records retrieved.",
        "ft_bulk_upload": "Bulk upload PDFs (optional)",
        "ft_single_upload": "Upload PDF (optional)",
        "ft_extract_text": "Extract text (PDF → text)",
        "ft_text_area": "Full-text text (paste here; OCR recommended)",
        "ft_ai_fill": "AI full-text review + extraction (this record)",
        "extract_schema": "Extraction schema (one column per line)",
        "extract_quick_add": "Quick add (enter once, then append)",
        "extract_editor": "Advanced: edit table (commit on save)",
        "extract_save": "Save / commit edits",
        "ma_outcome_label": "Outcome label (manual input)",
        "ma_measure": "Effect measure (OR/RR/HR/MD/SMD)",
        "ma_run": "Run MA + forest plot",
        "ms_generate": "(Optional) BYOK: generate richer draft (verify manually)",
        "export_docx": "Export Word (DOCX)",
    },
}


def t(key: str) -> str:
    lang = st.session_state.get("UI_LANG", "zh-TW")
    return I18N.get(lang, I18N["zh-TW"]).get(key, key)


# -------------------- Header --------------------
st.title(t("app_title"))
st.caption(f"{t('author')}　|　Language：{'繁體中文' if st.session_state.get('UI_LANG','zh-TW')=='zh-TW' else 'English'}　|　免責聲明：僅供學術用途；請自行驗證所有結果與引用。")


# -------------------- Helpers --------------------
def norm_text(x: Any) -> str:
    if x is None:
        return ""
    x = html.unescape(str(x))
    x = re.sub(r"\s+", " ", x).strip()
    return x


def short(s: str, n: int = 120) -> str:
    s = s or ""
    return (s[:n] + "…") if len(s) > n else s


def ensure_columns(df: pd.DataFrame, cols: List[str], default: Any = "") -> pd.DataFrame:
    if df is None:
        return pd.DataFrame({c: [] for c in cols})
    for c in cols:
        if c not in df.columns:
            df[c] = default
    return df


def safe_float(x: Any) -> Optional[float]:
    try:
        if x is None:
            return None
        s = str(x).strip()
        if s == "":
            return None
        return float(s)
    except Exception:
        return None


def pretty_json(d: Any) -> str:
    try:
        return json.dumps(d, ensure_ascii=False, indent=2)
    except Exception:
        return str(d)


def json_from_text(s: str) -> Optional[dict]:
    if not s:
        return None
    s = s.strip()
    try:
        return json.loads(s)
    except Exception:
        pass
    m = re.search(r"\{.*\}", s, flags=re.S)
    if m:
        try:
            return json.loads(m.group(0))
        except Exception:
            return None
    return None


def format_abstract(text: str) -> str:
    t0 = (text or "").strip()
    if not t0:
        return ""
    t0 = re.sub(r"\s*\n\s*", "\n", t0)
    t0 = re.sub(
        r"(?<!\n)\b(PURPOSE|METHODS|RESULTS|CONCLUSIONS|CONCLUSION|BACKGROUND|DESIGN|SETTING|PATIENTS|INTERVENTION|MAIN OUTCOME MEASURES|IMPORTANCE|OBJECTIVE|DATA SOURCES|STUDY SELECTION|DATA EXTRACTION|LIMITATIONS)\s*:\s*",
        r"\n\n\1: ",
        t0,
        flags=re.IGNORECASE,
    )
    if "\n" not in t0 and len(t0) > 800:
        t0 = re.sub(r"(?<=\.)\s+(?=[A-Z])", "\n\n", t0)
    return t0.strip()


def badge(label: str) -> str:
    label = label or "Unsure"
    if label == "Include":
        return f"<span class='badge badge-ok'>{label}</span>"
    if label == "Exclude":
        return f"<span class='badge badge-bad'>{label}</span>"
    return f"<span class='badge badge-warn'>{label}</span>"


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    if df is None:
        df = pd.DataFrame()
    return df.to_csv(index=False).encode("utf-8-sig")


# -------------------- Link helpers --------------------
def pubmed_link(pmid: str) -> str:
    pmid = str(pmid or "").strip().replace("PMID:", "").strip()
    return f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/" if pmid else ""


def doi_link(doi: str) -> str:
    doi = (doi or "").strip()
    return f"https://doi.org/{doi}" if doi else ""


def pmc_link(pmcid: str) -> str:
    pmcid = (pmcid or "").strip()
    if not pmcid:
        return ""
    if not pmcid.upper().startswith("PMC"):
        pmcid = "PMC" + pmcid
    return f"https://pmc.ncbi.nlm.nih.gov/articles/{pmcid}/"


def build_openurl(resolver_base: str, doi: str = "", pmid: str = "", title: str = "") -> str:
    resolver_base = (resolver_base or "").strip()
    if not resolver_base:
        return ""
    params = ["url_ver=Z39.88-2004", "ctx_ver=Z39.88-2004"]
    if doi:
        params.append("rft_id=" + requests.utils.quote(f"info:doi/{doi}"))
    elif pmid:
        params.append("rft_id=" + requests.utils.quote(f"info:pmid/{pmid}"))
    if title:
        params.append("rft.title=" + requests.utils.quote(title[:200]))
    joiner = "&" if "?" in resolver_base else "?"
    return resolver_base + joiner + "&".join(params)


def apply_ezproxy(ezproxy_prefix: str, url: str) -> str:
    ezproxy_prefix = (ezproxy_prefix or "").strip()
    url = (url or "").strip()
    if not ezproxy_prefix or not url:
        return url
    if "url=" in ezproxy_prefix:
        return ezproxy_prefix + requests.utils.quote(url, safe="")
    if ezproxy_prefix.endswith("/"):
        ezproxy_prefix = ezproxy_prefix[:-1]
    return ezproxy_prefix + "/login?url=" + requests.utils.quote(url, safe="")


# -------------------- MeSH suggestion (NLM) --------------------
@st.cache_data(show_spinner=False, ttl=60 * 60)
def mesh_suggest(term: str, limit: int = 6) -> List[str]:
    term = (term or "").strip()
    if not term:
        return []
    url = "https://id.nlm.nih.gov/mesh/lookup/descriptor"
    params = {"label": term, "match": "contains", "limit": str(limit)}
    try:
        r = requests.get(url, params=params, timeout=20)
        r.raise_for_status()
        data = r.json()
        labels = []
        for item in data:
            lab = item.get("label")
            if lab:
                labels.append(lab)
        out, seen = [], set()
        for x in labels:
            if x not in seen:
                out.append(x)
                seen.add(x)
        return out[:limit]
    except Exception:
        return []


def build_pubmed_block(term: str, mesh_label: Optional[str] = None) -> str:
    term = (term or "").strip()
    if not term:
        return ""
    if "[" in term and "]" in term:
        return term
    if mesh_label and mesh_label.strip():
        m = mesh_label.strip()
        return f'({term}[tiab] OR "{term}"[tiab] OR "{m}"[MeSH Terms])'
    return f'({term}[tiab] OR "{term}"[tiab] OR "{term}"[MeSH Terms])'


# -------------------- PubMed fetchers --------------------
NCBI_ESEARCH = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
NCBI_EFETCH = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"


def pubmed_esearch(query: str, retstart: int = 0, retmax: int = 200) -> Tuple[int, List[str], str, Dict[str, Any]]:
    params = {"db": "pubmed", "term": query, "retmode": "json", "retstart": retstart, "retmax": retmax}
    es_url = requests.Request("GET", NCBI_ESEARCH, params=params).prepare().url or ""
    diag: Dict[str, Any] = {"http_status": None, "content_type": None, "snippet": ""}
    try:
        r = requests.get(NCBI_ESEARCH, params=params, timeout=35, headers={"User-Agent": "srma-streamlit/1.0"})
        diag["http_status"] = r.status_code
        diag["content_type"] = r.headers.get("content-type", "")
        if "json" not in (diag["content_type"] or "").lower():
            diag["snippet"] = (r.text or "")[:2000]
            return 0, [], es_url, diag
        js = r.json().get("esearchresult", {})
        ids = js.get("idlist", []) or []
        count = int(js.get("count", 0) or 0)
        return count, ids, es_url, diag
    except Exception as e:
        diag["error"] = str(e)
        return 0, [], es_url, diag


def pubmed_efetch_xml(pmids: List[str]) -> Tuple[str, List[str]]:
    if not pmids:
        return "", []
    params = {"db": "pubmed", "id": ",".join(pmids), "retmode": "xml"}
    url = requests.Request("GET", NCBI_EFETCH, params=params).prepare().url or ""
    try:
        r = requests.get(NCBI_EFETCH, params=params, timeout=90, headers={"User-Agent": "srma-streamlit/1.0"})
        r.raise_for_status()
        return r.text, [url]
    except Exception:
        return "", [url]


def parse_pubmed_xml(xml_text: str) -> pd.DataFrame:
    if not (xml_text or "").strip():
        return pd.DataFrame()
    rows: List[Dict[str, Any]] = []
    try:
        root = ET.fromstring(xml_text)
    except Exception:
        return pd.DataFrame()

    for art in root.findall(".//PubmedArticle"):
        pmid = (art.findtext(".//PMID") or "").strip()
        title = norm_text(art.findtext(".//ArticleTitle") or "")
        ab_parts: List[str] = []
        for ab in art.findall(".//AbstractText"):
            lab = (ab.get("Label") or "").strip()
            t1 = norm_text(ab.text or "")
            if t1:
                ab_parts.append(f"{lab}: {t1}" if lab else t1)
        abstract = "\n".join(ab_parts).strip()

        year = (art.findtext(".//PubDate/Year") or "").strip()
        if not year:
            md = (art.findtext(".//PubDate/MedlineDate") or "").strip()
            m = re.search(r"(19|20)\d{2}", md)
            year = m.group(0) if m else ""

        journal = norm_text(art.findtext(".//Journal/Title") or "")

        # First author: handle both Individual and CollectiveName
        first_author = ""
        a0 = art.find(".//AuthorList/Author[1]")
        if a0 is not None:
            coll = (a0.findtext("CollectiveName") or "").strip()
            if coll:
                first_author = coll
            else:
                last = (a0.findtext("LastName") or "").strip()
                ini = (a0.findtext("Initials") or "").strip()
                fore = (a0.findtext("ForeName") or "").strip()
                if last and ini:
                    first_author = f"{last} {ini}".strip()
                elif last and fore:
                    ini2 = "".join([x[0] for x in fore.split() if x])[:4]
                    first_author = f"{last} {ini2}".strip()
                else:
                    first_author = last or fore

        doi = ""
        pmcid = ""
        for aid in art.findall(".//ArticleIdList/ArticleId"):
            idt = (aid.get("IdType") or "").lower()
            val = (aid.text or "").strip()
            if idt == "doi" and val and not doi:
                doi = val
            if idt == "pmc" and val and not pmcid:
                pmcid = val

        rows.append(
            {
                "record_id": f"PMID:{pmid}" if pmid else "",
                "pmid": pmid,
                "pmcid": pmcid,
                "doi": doi,
                "year": year,
                "journal": journal,
                "first_author": first_author,
                "title": title,
                "abstract": abstract,
                "source": "PubMed",
            }
        )

    df = pd.DataFrame(rows)
    df = ensure_columns(df, ["record_id", "pmid", "pmcid", "doi", "year", "journal", "first_author", "title", "abstract", "source"], "")
    if not df.empty and "pmid" in df.columns:
        df = df.drop_duplicates(subset=["pmid"], keep="first").reset_index(drop=True)
    return df


# -------------------- Protocol & heuristics --------------------
@dataclass
class Protocol:
    P: str = ""
    I: str = ""
    C: str = ""
    O: str = ""
    NOT: str = "animal OR mice OR rat OR in vitro OR case report"
    goal_mode: str = "Fast / feasible (gap-fill)"

    P_syn: List[str] = None
    I_syn: List[str] = None
    C_syn: List[str] = None
    O_syn: List[str] = None
    mesh_P: List[str] = None
    mesh_I: List[str] = None
    mesh_C: List[str] = None
    mesh_O: List[str] = None

    inclusion: str = ""
    exclusion: str = ""
    outcomes_plan: str = ""
    extraction_plan: str = ""
    feasibility_note: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "pico": {"P": self.P, "I": self.I, "C": self.C, "O": self.O, "NOT": self.NOT},
            "goal_mode": self.goal_mode,
            "search_expansion": {
                "P_synonyms": self.P_syn or [],
                "I_synonyms": self.I_syn or [],
                "C_synonyms": self.C_syn or [],
                "O_synonyms": self.O_syn or [],
                "NOT": [x.strip() for x in (self.NOT or "").split(" OR ") if x.strip()],
            },
            "mesh_candidates": {"P": self.mesh_P or [], "I": self.mesh_I or [], "C": self.mesh_C or [], "O": self.mesh_O or []},
            "criteria": {"inclusion": self.inclusion, "exclusion": self.exclusion},
            "plans": {"outcomes_plan": self.outcomes_plan, "extraction_plan": self.extraction_plan},
            "feasibility": {"note": self.feasibility_note},
        }


def guess_pico_from_question(q: str) -> Protocol:
    q = (q or "").strip()
    # Heuristic: if "vs" present, split into I/C. Otherwise keep as I, P left generic.
    I, C = q, ""
    m = re.split(r"\s+vs\.?\s+|\s+versus\s+|比較|對照", q, flags=re.IGNORECASE)
    if len(m) >= 2:
        I = m[0].strip()
        C = m[1].strip()

    proto = Protocol(P=q, I=I or q, C=C, O="")
    proto.P_syn, proto.I_syn, proto.C_syn, proto.O_syn = [], [], [], []
    proto.mesh_P, proto.mesh_I, proto.mesh_C, proto.mesh_O = [], [], [], []
    proto.inclusion = "『研究設計/族群/介入-比較/追蹤時間/主要結局』請依臨床情境與可行性報告細化。"
    proto.exclusion = "動物、體外、病例報告/小型 case series（依題目調整）。"
    proto.outcomes_plan = (
        "開始所有步驟前：先掃描既有 SR/MA/NMA，產出可行性報告並決定是否縮題/改方向。\n"
        "Outcome 規劃：請盤點所有 RCT 的 primary + secondary outcomes，並指定主要 outcome（MA primary）。"
    )
    proto.extraction_plan = (
        "Extraction table 不寫死：在 PICO 層級先規劃抽取欄位（基線特徵、介入細節、比較組、追蹤時間點、所有 primary/secondary outcomes、"
        "安全性事件、研究設計要素），並在比對既有 SR/MA 後調整（補 gap 或對齊既有結局）。\n"
        "若全文為掃描 PDF：建議先 OCR（Adobe/Drive OCR 等）再上傳；抽取 figure/table 時，請在 AI 提示中要求讀取 Figure/Table。"
    )
    return proto


# -------------------- LLM (BYOK) --------------------
def llm_available() -> bool:
    return bool(st.session_state.get("BYOK_ENABLED")) and bool(st.session_state.get("BYOK_KEY")) and bool(st.session_state.get("BYOK_CONSENT"))


def llm_chat(messages: List[Dict[str, str]], temperature: float = 0.2, timeout: int = 90) -> str:
    base = (st.session_state.get("BYOK_BASE_URL") or "https://api.openai.com/v1").rstrip("/")
    model = st.session_state.get("BYOK_MODEL") or "gpt-4o-mini"
    key = st.session_state.get("BYOK_KEY") or ""
    url = base + "/chat/completions"
    headers = {"Authorization": f"Bearer {key}", "Content-Type": "application/json"}
    payload = {"model": model, "messages": messages, "temperature": float(temperature), "max_tokens": 1400}
    r = requests.post(url, headers=headers, json=payload, timeout=timeout)
    r.raise_for_status()
    js = r.json()
    return (js.get("choices") or [{}])[0].get("message", {}).get("content", "") or ""


def llm_build_protocol(q: str, lang: str) -> Protocol:
    sys = "You are an expert systematic reviewer. Output ONLY valid JSON."
    user = f"""
Output language: {"Traditional Chinese" if lang=="zh-TW" else "English"}.
Given the research question below, propose a Protocol for SR/MA (and optionally NMA).
Return JSON with:
- P,I,C,O (strings)
- P_syn, I_syn, C_syn, O_syn (list of strings; include abbreviations, brand/model names, free-text terms)
- NOT (string for PubMed NOT block; include animals/in vitro/case report by default)
- goal_mode
- inclusion (PICO-level rules; specify study designs)
- exclusion
- outcomes_plan (must mention prior SR/MA/NMA scan and primary+secondary outcomes across RCTs)
- extraction_plan (must mention OCR and figure/table extraction hints; extraction sheet should be self-planned)
Return JSON only.

Research question:
{q}
""".strip()
    txt = llm_chat([{"role": "system", "content": sys}, {"role": "user", "content": user}], temperature=float(st.session_state.get("BYOK_TEMP", 0.2) or 0.2), timeout=120)
    js = json_from_text(txt) or {}
    proto = Protocol(
        P=str(js.get("P", q) or q),
        I=str(js.get("I", q) or q),
        C=str(js.get("C", "") or ""),
        O=str(js.get("O", "") or ""),
        NOT=str(js.get("NOT", "animal OR mice OR rat OR in vitro OR case report") or "animal OR mice OR rat OR in vitro OR case report"),
        goal_mode=str(js.get("goal_mode", "Fast / feasible (gap-fill)") or "Fast / feasible (gap-fill)"),
        P_syn=list(js.get("P_syn") or []),
        I_syn=list(js.get("I_syn") or []),
        C_syn=list(js.get("C_syn") or []),
        O_syn=list(js.get("O_syn") or []),
        mesh_P=[],
        mesh_I=[],
        mesh_C=[],
        mesh_O=[],
        inclusion=str(js.get("inclusion", "") or ""),
        exclusion=str(js.get("exclusion", "") or ""),
        outcomes_plan=str(js.get("outcomes_plan", "") or ""),
        extraction_plan=str(js.get("extraction_plan", "") or ""),
        feasibility_note="",
    )
    return proto


def llm_screen_title_abstract(batch: List[Dict[str, Any]], proto: Protocol, lang: str) -> Dict[str, Dict[str, Any]]:
    sys = "You are a strict SR/MA title/abstract screener. Output ONLY valid JSON."
    user = f"""
Output language: {"Traditional Chinese" if lang=="zh-TW" else "English"}.

Protocol:
{pretty_json(proto.to_dict())}

For each record, decide: Include / Exclude / Unsure for full-text review.
Return JSON object:
{{ "<record_id>": {{"decision": "...", "confidence": 0-1, "reason": "...", "matched_rules": "..."}} }}

Records:
{pretty_json(batch)}
""".strip()
    out = llm_chat([{"role": "system", "content": sys}, {"role": "user", "content": user}], temperature=0.1, timeout=180)
    js = json_from_text(out) or {}
    res: Dict[str, Dict[str, Any]] = {}
    for rid, v in (js or {}).items():
        if not isinstance(v, dict):
            continue
        dec = str(v.get("decision", "Unsure") or "Unsure")
        if dec not in ["Include", "Exclude", "Unsure"]:
            dec = "Unsure"
        conf = safe_float(v.get("confidence", 0.5)) or 0.5
        conf = max(0.0, min(1.0, conf))
        res[str(rid)] = {
            "decision": dec,
            "confidence": conf,
            "reason": str(v.get("reason", "") or ""),
            "matched_rules": str(v.get("matched_rules", "") or ""),
        }
    return res


def llm_feasibility_report(q: str, proto: Protocol, hits: List[Dict[str, Any]], lang: str) -> Dict[str, Any]:
    sys = "You are an SR/MA methodologist. Output ONLY valid JSON."
    user = f"""
Output language: {"Traditional Chinese" if lang=="zh-TW" else "English"}.

Research question: {q}

Existing SR/MA/NMA hits (top):
{pretty_json(hits[:15])}

Produce feasibility recommendation JSON:
- status: "Feasible" / "Needs refinement" / "Consider different question"
- why
- suggested_modifications (list)
- study_design_advice
- nma_advice
- extraction_schema_suggestions (list; must mention primary+secondary outcomes across RCTs)
- inclusion_criteria_suggestions (PICO-level; align with goal_mode)
Return JSON only.
""".strip()
    out = llm_chat([{"role": "system", "content": sys}, {"role": "user", "content": user}], temperature=0.2, timeout=180)
    return json_from_text(out) or {}


def llm_fulltext_extract_one(fulltext: str, proto: Protocol, schema: List[str], lang: str) -> Dict[str, Any]:
    sys = "You are an SR/MA full-text reviewer and extractor. Output ONLY valid JSON."
    user = f"""
Output language: {"Traditional Chinese" if lang=="zh-TW" else "English"}.

Protocol:
{pretty_json(proto.to_dict())}

Extraction schema columns:
{pretty_json(schema)}

Task:
1) Decide full-text eligibility:
- fulltext_decision: "Include for meta-analysis" / "Include (qualitative only)" / "Exclude after full-text"
- fulltext_reason: short justification aligned with inclusion/exclusion criteria.

2) Extract key fields into 'extracted_fields' dict mapping schema column -> value (string/number). If cannot extract, leave blank.

3) If you see usable effect size and 95% CI, populate:
meta: {{
  "effect_measure":"OR/RR/HR/MD/SMD",
  "effect": number,
  "lower_CI": number,
  "upper_CI": number,
  "outcome_label": "...",
  "timepoint": "...",
  "effect_unit":"..."
}}
If not available, return empty meta.

OCR/figure/table note:
- If scanned PDF, content may be missing; prefer extracting from tables/figures if present.
- Explicitly look for "Table" and "Figure".

Full text:
{(fulltext or "")[:160000]}
""".strip()
    out = llm_chat([{"role": "system", "content": sys}, {"role": "user", "content": user}], temperature=0.1, timeout=240)
    return json_from_text(out) or {}


def llm_rob2_suggest(fulltext: str, lang: str) -> Dict[str, Any]:
    sys = "You are an expert in Cochrane RoB 2. Output ONLY valid JSON."
    user = f"""
Output language: {"Traditional Chinese" if lang=="zh-TW" else "English"}.

Given the full text below, provide RoB 2 judgments with reasons.

Return JSON:
{{
  "randomization": {{"level":"Low/Some concerns/High/NA","reason":"..."}},
  "deviations": {{"level":"Low/Some concerns/High/NA","reason":"..."}},
  "missing_data": {{"level":"Low/Some concerns/High/NA","reason":"..."}},
  "measurement": {{"level":"Low/Some concerns/High/NA","reason":"..."}},
  "selection": {{"level":"Low/Some concerns/High/NA","reason":"..."}},
  "overall": {{"level":"Low/Some concerns/High/NA","reason":"..."}}
}}

Full text:
{(fulltext or "")[:200000]}
""".strip()
    out = llm_chat([{"role": "system", "content": sys}, {"role": "user", "content": user}], temperature=0.2, timeout=240)
    return json_from_text(out) or {}


# -------------------- Search query builder --------------------
ARTICLE_TYPE_FILTERS = {
    "不限": "",
    "RCT": '("randomized controlled trial"[pt] OR randomized[tiab] OR randomised[tiab] OR trial[tiab])',
    "Systematic review / Meta-analysis": '("systematic review"[pt] OR "meta-analysis"[pt] OR systematic[tiab] OR meta-analysis[tiab] OR metaanalysis[tiab])',
    "Cohort/Observational": '(cohort[tiab] OR observational[tiab] OR registry[tiab])',
}


def build_pubmed_query(proto: Protocol, article_type: str, custom_filter: str) -> str:
    P_terms = [proto.P] + (proto.P_syn or [])
    I_terms = [proto.I] + (proto.I_syn or [])
    C_terms = ([proto.C] + (proto.C_syn or [])) if proto.C else []
    O_terms = ([proto.O] + (proto.O_syn or [])) if proto.O else []

    P_mesh = (proto.mesh_P or [None])[0] if proto.mesh_P else None
    I_mesh = (proto.mesh_I or [None])[0] if proto.mesh_I else None
    C_mesh = (proto.mesh_C or [None])[0] if proto.mesh_C else None
    O_mesh = (proto.mesh_O or [None])[0] if proto.mesh_O else None

    blocks = []
    if P_terms:
        pt = " OR ".join([build_pubmed_block(t1, P_mesh) for t1 in P_terms if str(t1).strip()])
        if pt:
            blocks.append(f"({pt})")
    if I_terms:
        it = " OR ".join([build_pubmed_block(t1, I_mesh) for t1 in I_terms if str(t1).strip()])
        if it:
            blocks.append(f"({it})")
    if C_terms:
        ct = " OR ".join([build_pubmed_block(t1, C_mesh) for t1 in C_terms if str(t1).strip()])
        if ct:
            blocks.append(f"({ct})")
    if O_terms:
        ot = " OR ".join([build_pubmed_block(t1, O_mesh) for t1 in O_terms if str(t1).strip()])
        if ot:
            blocks.append(f"({ot})")

    core = " AND ".join(blocks) if blocks else f'("{proto.P}"[tiab])'
    not_block = (proto.NOT or "").strip()
    not_expr = f" NOT ({not_block})" if not_block else ""

    at = ARTICLE_TYPE_FILTERS.get(article_type, "")
    cf = (custom_filter or "").strip()
    filters = []
    if at:
        filters.append(f"({at})")
    if cf:
        filters.append(f"({cf})")
    filt = (" AND " + " AND ".join(filters)) if filters else ""
    return f"({core}){filt}{not_expr}"


def build_feas_query(proto: Protocol) -> str:
    key = proto.P or proto.I or ""
    core = f'("{key}"[tiab])' if key else "(clinical[tiab])"
    sr = '("systematic review"[pt] OR "meta-analysis"[pt] OR systematic[tiab] OR meta-analysis[tiab] OR metaanalysis[tiab] OR "network meta-analysis"[tiab] OR "network meta analysis"[tiab])'
    return f"({core}) AND ({sr}) NOT ({proto.NOT})"


# -------------------- Screening fallback (no LLM) --------------------
def simple_rule_screen(title: str, abstract: str, proto: Protocol) -> Tuple[str, float, str, str]:
    t0 = (title or "").lower()
    a0 = (abstract or "").lower()
    blob = t0 + " " + a0

    if any(x in blob for x in ["animal", "mice", "mouse", "rat", "in vitro", "case report"]):
        return "Exclude", 0.9, "Contains animal/in vitro/case report signal.", "NOT-signal"

    q_terms = set(re.findall(r"[a-z0-9\-]+", (proto.P or proto.I or "").lower()))
    q_terms = {x for x in q_terms if len(x) >= 4}

    hit, matched = 0, []
    for w in sorted(q_terms):
        if w in blob:
            hit += 1
            matched.append(w)

    conf = min(0.9, 0.35 + 0.08 * hit)
    if hit >= 2:
        return "Include", conf, "P/I free-text terms matched in title/abstract.", ", ".join(matched[:12])
    if hit == 1:
        return "Unsure", 0.55, "Partial match; needs full-text verification.", ", ".join(matched[:12])
    return "Unsure", 0.45, "No strong match; keep for feasibility or broaden query.", ""


# -------------------- PRISMA (text-based) --------------------
def compute_effective_ta_decision(record_id: str) -> str:
    override = (st.session_state.get("TA_OVERRIDE", {}) or {}).get(record_id)
    if override in ["Include", "Exclude", "Unsure"]:
        return override
    return (st.session_state.get("TA_AI", {}) or {}).get(record_id, "Unsure")


def compute_prisma(df_records: pd.DataFrame) -> Dict[str, Any]:
    total_retrieved = int(st.session_state.get("PUBMED_TOTAL", 0) or 0)
    n_fetched = int(len(df_records)) if df_records is not None else 0

    ta_inc = ta_exc = ta_uns = 0
    rids = df_records["record_id"].tolist() if df_records is not None and not df_records.empty else []
    for rid in rids:
        eff = compute_effective_ta_decision(rid)
        if eff == "Include":
            ta_inc += 1
        elif eff == "Exclude":
            ta_exc += 1
        else:
            ta_uns += 1

    ft = st.session_state.get("FT_DECISIONS", {}) or {}
    ft_inc_meta = 0
    ft_inc_qual = 0
    ft_exc = 0
    ft_assessed = 0
    for rid in rids:
        d = (ft.get(rid) or "Not reviewed yet").strip()
        if d != "Not reviewed yet":
            ft_assessed += 1
        if d == "Include for meta-analysis":
            ft_inc_meta += 1
        elif d == "Include (qualitative only)":
            ft_inc_qual += 1
        elif d == "Exclude after full-text":
            ft_exc += 1

    studies_included = ft_inc_meta + ft_inc_qual
    included_meta = ft_inc_meta

    return {
        "records_identified": total_retrieved,
        "records_fetched": n_fetched,
        "ta_include": ta_inc,
        "ta_exclude": ta_exc,
        "ta_unsure": ta_uns,
        "fulltext_assessed": ft_assessed,
        "fulltext_excluded": ft_exc,
        "studies_included": studies_included,
        "included_meta": included_meta,
    }


def render_prisma_text(pr: Dict[str, Any]):
    st.markdown("#### PRISMA（文字版；可對照 RevMan/PRISMA flow）")
    a = pr
    st.markdown(
        f"""
<div class="flow">
  <div class="flow-box"><div class="t">Records identified</div><div class="n">PubMed count：<b>{a.get('records_identified','—')}</b></div></div>
  <div class="flow-arrow">↓</div>
  <div class="flow-box"><div class="t">Records fetched (details)</div><div class="n">Efetch 解析到：<b>{a.get('records_fetched','—')}</b></div></div>
  <div class="flow-arrow">↓</div>
  <div class="flow-row">
    <div class="flow-box"><div class="t">Title/Abstract: Include</div><div class="n"><b>{a.get('ta_include','—')}</b></div></div>
    <div class="flow-box"><div class="t">Title/Abstract: Exclude</div><div class="n"><b>{a.get('ta_exclude','—')}</b></div></div>
  </div>
  <div class="flow-row">
    <div class="flow-box"><div class="t">Title/Abstract: Unsure</div><div class="n"><b>{a.get('ta_unsure','—')}</b></div></div>
    <div class="flow-box"><div class="t">Full-text assessed</div><div class="n"><b>{a.get('fulltext_assessed','—')}</b></div></div>
  </div>
  <div class="flow-row">
    <div class="flow-box"><div class="t">Full-text excluded (with reasons)</div><div class="n"><b>{a.get('fulltext_excluded','—')}</b></div></div>
    <div class="flow-box"><div class="t">Studies included</div><div class="n">Qualitative：<b>{a.get('studies_included','—')}</b>；Meta-analysis：<b>{a.get('included_meta','—')}</b></div></div>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )


# -------------------- Full text extraction helper --------------------
def extract_text_from_pdf(file) -> str:
    if not HAS_PYPDF2:
        return ""
    try:
        reader = PdfReader(file)
        texts = []
        for page in reader.pages[:80]:
            tx = page.extract_text() or ""
            if tx.strip():
                texts.append(tx)
        return "\n".join(texts).strip()
    except Exception:
        return ""


# -------------------- MA helpers (fixed effect) --------------------
def se_from_ci(effect: float, lcl: float, ucl: float, measure: str) -> Optional[float]:
    """
    Robust SE from 95% CI.
    - OR/RR/HR: use log scale; require effect/lcl/ucl > 0
    - MD/SMD: use linear scale
    """
    measure = (measure or "").upper().strip()
    if measure in ("OR", "RR", "HR"):
        if effect <= 0 or lcl <= 0 or ucl <= 0:
            return None
        return (math.log(ucl) - math.log(lcl)) / 3.92
    # MD/SMD
    return (ucl - lcl) / 3.92


def fixed_effect_meta(effects: List[float], ses: List[float], measure: str) -> Dict[str, Any]:
    """
    Inverse-variance fixed effect.
    Returns pooled effect on natural scale.
    """
    w = []
    for se in ses:
        if se is None or se <= 0:
            w.append(None)
        else:
            w.append(1.0 / (se ** 2))

    pairs = [(e, wi, se) for e, wi, se in zip(effects, w, ses) if (wi is not None and wi > 0 and se is not None)]
    if len(pairs) == 0:
        return {"ok": False, "error": "No valid studies for pooling."}

    measure = (measure or "").upper().strip()
    if measure in ("OR", "RR", "HR"):
        y = [math.log(p[0]) for p in pairs]
    else:
        y = [p[0] for p in pairs]
    wi = [p[1] for p in pairs]

    sw = sum(wi)
    mu = sum(wi_i * yi for wi_i, yi in zip(wi, y)) / sw
    se_mu = math.sqrt(1.0 / sw)
    l = mu - 1.96 * se_mu
    u = mu + 1.96 * se_mu

    # heterogeneity (Q, I^2) for reporting (still fixed-effect pooled)
    Q = sum(wi_i * (yi - mu) ** 2 for wi_i, yi in zip(wi, y))
    df = max(0, len(y) - 1)
    I2 = 0.0
    if Q > 0 and df > 0:
        I2 = max(0.0, (Q - df) / Q) * 100.0

    if measure in ("OR", "RR", "HR"):
        pooled = math.exp(mu)
        lcl = math.exp(l)
        ucl = math.exp(u)
    else:
        pooled = mu
        lcl = l
        ucl = u

    return {
        "ok": True,
        "k": len(y),
        "measure": measure,
        "pooled": pooled,
        "lcl": lcl,
        "ucl": ucl,
        "se_pooled": se_mu,
        "Q": Q,
        "df": df,
        "I2": I2,
    }


def forest_plot(df_plot: pd.DataFrame, pooled: Dict[str, Any], measure: str):
    """
    Draw forest plot using Plotly if available, else Matplotlib, else show table.
    Expects columns: label, effect, lcl, ucl
    """
    if df_plot is None or df_plot.empty:
        st.info("沒有可畫圖的資料。")
        return

    measure = (measure or "").upper().strip()
    # Build labels
    labels = df_plot["label"].tolist()
    eff = df_plot["effect"].tolist()
    lcl = df_plot["lcl"].tolist()
    ucl = df_plot["ucl"].tolist()

    # Append pooled row
    labels2 = labels + ["Pooled (fixed)"]
    eff2 = eff + [pooled.get("pooled")]
    lcl2 = lcl + [pooled.get("lcl")]
    ucl2 = ucl + [pooled.get("ucl")]

    # Plotly
    if HAS_PLOTLY:
        # x-axis: effect scale (log for ratio measures)
        if measure in ("OR", "RR", "HR"):
            x = [math.log(x) if x and x > 0 else None for x in eff2]
            xl = [math.log(x) if x and x > 0 else None for x in lcl2]
            xu = [math.log(x) if x and x > 0 else None for x in ucl2]
            x_title = f"log({measure})"
        else:
            x = eff2
            xl = lcl2
            xu = ucl2
            x_title = measure

        y = list(range(len(labels2)))[::-1]
        fig = go.Figure()
        # CI lines
        for yi, xm, x1, x2, lab in zip(y, x, xl, xu, labels2):
            if xm is None or x1 is None or x2 is None:
                continue
            fig.add_trace(go.Scatter(x=[x1, x2], y=[yi, yi], mode="lines", showlegend=False))
            fig.add_trace(go.Scatter(x=[xm], y=[yi], mode="markers", showlegend=False, marker=dict(size=9)))
        fig.update_yaxes(
            tickmode="array",
            tickvals=y,
            ticktext=labels2,
            autorange=False,
        )
        fig.update_layout(height=max(360, 60 * len(labels2)), xaxis_title=x_title, margin=dict(l=10, r=10, t=20, b=20))
        st.plotly_chart(fig, use_container_width=True)
        return

    # Matplotlib fallback
    if HAS_MPL:
        if measure in ("OR", "RR", "HR"):
            x = [math.log(x) if x and x > 0 else None for x in eff2]
            xl = [math.log(x) if x and x > 0 else None for x in lcl2]
            xu = [math.log(x) if x and x > 0 else None for x in ucl2]
            xlab = f"log({measure})"
        else:
            x, xl, xu = eff2, lcl2, ucl2
            xlab = measure

        fig = plt.figure(figsize=(8, max(3, 0.5 * len(labels2))))
        ax = plt.gca()
        y = list(range(len(labels2)))[::-1]
        for yi, xm, x1, x2 in zip(y, x, xl, xu):
            if xm is None or x1 is None or x2 is None:
                continue
            ax.plot([x1, x2], [yi, yi])
            ax.plot([xm], [yi], marker="s")
        ax.set_yticks(y)
        ax.set_yticklabels(labels2)
        ax.set_xlabel(xlab)
        ax.axvline(0 if measure in ("MD", "SMD") else 0, linestyle="--")  # reference line at 0 on chosen scale
        st.pyplot(fig, clear_figure=True)
        return

    st.info("環境缺少 Plotly/Matplotlib：改以表格顯示森林圖資料。")
    st.dataframe(pd.DataFrame({"label": labels2, "effect": eff2, "lcl": lcl2, "ucl": ucl2}), use_container_width=True)


# -------------------- RoB2 helpers --------------------
ROB_LEVELS = ["Low", "Some concerns", "High", "NA"]
ROB_DOMAINS = [
    ("randomization", "Randomization process"),
    ("deviations", "Deviations from intended interventions"),
    ("missing_data", "Missing outcome data"),
    ("measurement", "Measurement of the outcome"),
    ("selection", "Selection of the reported result"),
    ("overall", "Overall Risk of Bias"),
]


def init_rob2_for_record(rid: str):
    rob = st.session_state.get("ROB2", {}) or {}
    if rid not in rob:
        rob[rid] = {k: {"level": "", "reason": ""} for k, _ in ROB_DOMAINS}
    st.session_state["ROB2"] = rob


# -------------------- Session state (KeyError fix) --------------------
def init_state():
    ss = st.session_state
    # Language
    ss.setdefault("UI_LANG", "zh-TW")

    # BYOK
    ss.setdefault("BYOK_ENABLED", False)
    ss.setdefault("BYOK_KEY", "")
    ss.setdefault("BYOK_BASE_URL", "https://api.openai.com/v1")
    ss.setdefault("BYOK_MODEL", "gpt-4o-mini")
    ss.setdefault("BYOK_TEMP", 0.2)
    ss.setdefault("BYOK_CONSENT", False)

    # Institutional links
    ss.setdefault("RESOLVER_BASE", "")
    ss.setdefault("EZPROXY_PREFIX", "")

    # Core
    ss.setdefault("QUESTION", "")
    ss.setdefault("ARTICLE_TYPE", "不限")
    ss.setdefault("CUSTOM_PUBMED_FILTER", "")
    ss.setdefault("GOAL_MODE", "Fast / feasible (gap-fill)")

    # Protocol
    ss.setdefault("PROTOCOL", guess_pico_from_question(""))
    ss.setdefault("PUBMED_QUERY_AUTO", "")
    ss.setdefault("PUBMED_QUERY_MANUAL", "")
    ss.setdefault("FEAS_QUERY", "")

    # Data
    ss.setdefault("PUBMED_TOTAL", 0)
    ss.setdefault("PUBMED_RECORDS", pd.DataFrame())
    ss.setdefault("SRMA_HITS", pd.DataFrame())
    ss.setdefault("DIAGNOSTICS", {})

    # Screening
    ss.setdefault("TA_AI", {})
    ss.setdefault("TA_AI_REASON", {})
    ss.setdefault("TA_AI_CONF", {})
    ss.setdefault("TA_AI_RULES", {})
    ss.setdefault("TA_OVERRIDE", {})
    ss.setdefault("TA_OVERRIDE_REASON", {})

    # Full text
    ss.setdefault("FT_DECISIONS", {})
    ss.setdefault("FT_REASONS", {})
    ss.setdefault("FT_NOTE", {})
    ss.setdefault("FT_TEXT", {})
    ss.setdefault("AI_EXTRACTION_CACHE", {})

    # Extraction schema + table
    ss.setdefault("SCHEMA_COLS_TEXT", "Outcome_label\nTimepoint\nEffect_measure\nEffect\nLower_CI\nUpper_CI\nEffect_unit\nNotes")
    ss.setdefault("SCHEMA_COLS", ["Outcome_label", "Timepoint", "Effect_measure", "Effect", "Lower_CI", "Upper_CI", "Effect_unit", "Notes"])
    ss.setdefault("EXTRACT_DF", pd.DataFrame(columns=["record_id", "Outcome_label", "Timepoint", "Effect_measure", "Effect", "Lower_CI", "Upper_CI", "Effect_unit", "Notes"]))
    ss.setdefault("EXTRACT_EDITOR_DF", None)  # staging for data_editor form
    ss.setdefault("EXTRACT_COMMITTED_AT", 0.0)

    # MA
    ss.setdefault("MA_RESULT", {})
    ss.setdefault("MA_WARNINGS", [])
    ss.setdefault("MA_OUTCOME_FILTER", "")
    ss.setdefault("MA_MEASURE", "OR")
    ss.setdefault("MA_LAST_RUN", 0.0)

    # ROB2
    ss.setdefault("ROB2", {})

    # Manuscript
    ss.setdefault("MS_DRAFT", {})
    ss.setdefault("MS_LAST_GEN", 0.0)


init_state()


# -------------------- Reset downstream --------------------
def reset_downstream(keep_query: bool = True):
    st.session_state["PUBMED_TOTAL"] = 0
    st.session_state["PUBMED_RECORDS"] = pd.DataFrame()
    st.session_state["SRMA_HITS"] = pd.DataFrame()
    st.session_state["DIAGNOSTICS"] = {}

    st.session_state["TA_AI"] = {}
    st.session_state["TA_AI_REASON"] = {}
    st.session_state["TA_AI_CONF"] = {}
    st.session_state["TA_AI_RULES"] = {}
    st.session_state["TA_OVERRIDE"] = {}
    st.session_state["TA_OVERRIDE_REASON"] = {}

    st.session_state["FT_DECISIONS"] = {}
    st.session_state["FT_REASONS"] = {}
    st.session_state["FT_NOTE"] = {}
    st.session_state["FT_TEXT"] = {}
    st.session_state["AI_EXTRACTION_CACHE"] = {}

    st.session_state["EXTRACT_DF"] = pd.DataFrame(columns=["record_id", "Outcome_label", "Timepoint", "Effect_measure", "Effect", "Lower_CI", "Upper_CI", "Effect_unit", "Notes"])
    st.session_state["EXTRACT_EDITOR_DF"] = None
    st.session_state["EXTRACT_COMMITTED_AT"] = 0.0

    st.session_state["MA_RESULT"] = {}
    st.session_state["MA_WARNINGS"] = []
    st.session_state["MA_LAST_RUN"] = 0.0

    st.session_state["ROB2"] = {}

    st.session_state["MS_DRAFT"] = {}
    st.session_state["MS_LAST_GEN"] = 0.0

    if not keep_query:
        st.session_state["PUBMED_QUERY_AUTO"] = ""
        st.session_state["PUBMED_QUERY_MANUAL"] = ""
        st.session_state["FEAS_QUERY"] = ""


# -------------------- Sidebar --------------------
with st.sidebar:
    st.subheader(t("settings"))

    st.selectbox(
        t("lang_label"),
        options=["zh-TW", "en"],
        format_func=lambda x: "繁體中文" if x == "zh-TW" else "English",
        key="UI_LANG",
    )

    with st.expander(t("tips_title"), expanded=True):
        st.markdown(
            """
- **僅供學術用途、需人工核對**：任何結論、數值、引用與自動萃取都可能有錯，請逐一回到原文核對。
- **勿上傳可識別病人資訊**：姓名、病歷號、影像、日期、或任何可回溯個資。
- **校內訂閱/授權全文上傳風險**：避免把學校資源 PDF 上傳到雲端（尤其公開部署）；遵守圖書館授權條款，避免大量下載/自動化擷取與共享全文。
- **PubMed/eUtils 被擋**：若 count=0 或 Diagnostics 顯示 content-type 非 JSON/XML，可能被擋或回傳 HTML；請開啟 Diagnostics 檢查 esearch_url 與 snippet。
- **未啟用 LLM 時自動降級**：不會卡在 extraction/ROB2；AI 建議改為規則式推論，你仍可人工修正與完成後續步驟。
            """.strip()
        )

    st.markdown("---")
    st.subheader(t("byok_title"))
    st.toggle(t("byok_toggle"), value=bool(st.session_state["BYOK_ENABLED"]), key="BYOK_ENABLED")

    if st.session_state["BYOK_ENABLED"]:
        st.caption(t("byok_notice"))
        st.text_input("Base URL", value=st.session_state["BYOK_BASE_URL"], key="BYOK_BASE_URL")
        st.text_input("Model", value=st.session_state["BYOK_MODEL"], key="BYOK_MODEL")
        st.text_input("API key", value=st.session_state["BYOK_KEY"], key="BYOK_KEY", type="password")
        st.slider("temperature", 0.0, 1.0, float(st.session_state["BYOK_TEMP"]), 0.05, key="BYOK_TEMP")
        st.checkbox(t("byok_consent"), value=bool(st.session_state["BYOK_CONSENT"]), key="BYOK_CONSENT")
        c1, c2 = st.columns(2)
        with c1:
            if st.button(t("byok_clear")):
                st.session_state["BYOK_KEY"] = ""
                st.session_state["BYOK_CONSENT"] = False
                st.rerun()
        with c2:
            st.write("")

    st.markdown("---")
    st.subheader(t("links_title"))
    st.text_input(t("resolver"), value=st.session_state["RESOLVER_BASE"], key="RESOLVER_BASE")
    st.text_input(t("ezproxy"), value=st.session_state["EZPROXY_PREFIX"], key="EZPROXY_PREFIX")

    st.markdown("---")
    st.subheader(t("search_settings"))
    st.selectbox(t("article_type"), list(ARTICLE_TYPE_FILTERS.keys()), key="ARTICLE_TYPE")
    st.text_input(t("custom_filter"), value=st.session_state["CUSTOM_PUBMED_FILTER"], key="CUSTOM_PUBMED_FILTER")
    st.selectbox(
        t("goal_mode"),
        ["Fast / feasible (gap-fill)", "Rigorous (narrow inclusion)"],
        index=0 if str(st.session_state["GOAL_MODE"]).startswith("Fast") else 1,
        key="GOAL_MODE",
    )


# -------------------- Main input --------------------
st.markdown(
    f"<div class='notice'><b>{t('question_notice')}</b><br>{t('question_help')}</div>",
    unsafe_allow_html=True,
)

st.text_input(t("question_label"), value=st.session_state["QUESTION"], key="QUESTION")
run = st.button(t("run"), type="primary", disabled=not bool((st.session_state["QUESTION"] or "").strip()))


# -------------------- Pipeline runners --------------------
def run_pubmed_pipeline(pubmed_query: str):
    total, ids, es_url, es_diag = pubmed_esearch(pubmed_query, retstart=0, retmax=200)
    xml_text, ef_urls = pubmed_efetch_xml(ids[:200])
    df = parse_pubmed_xml(xml_text)

    st.session_state["PUBMED_RECORDS"] = df
    st.session_state["PUBMED_TOTAL"] = total

    diag = st.session_state.get("DIAGNOSTICS", {}) or {}
    diag.update(
        {
            "pubmed_total_count": total,
            "esearch_url": es_url,
            "efetch_urls": ef_urls,
            "esearch_diag": es_diag,
            "warnings": [] if total > 0 else ["PubMed count=0：請把問題寫更具體（縮寫寫全名/加型號/族群/outcome），或看 Diagnostics 是否被阻擋。"],
        }
    )
    st.session_state["DIAGNOSTICS"] = diag

    # feasibility scan
    feas_q = st.session_state.get("FEAS_QUERY") or build_feas_query(st.session_state["PROTOCOL"])
    st.session_state["FEAS_QUERY"] = feas_q
    feas_total, feas_ids, feas_url, feas_diag = pubmed_esearch(feas_q, retstart=0, retmax=20)
    feas_xml, feas_ef = pubmed_efetch_xml(feas_ids[:20])
    df_feas = parse_pubmed_xml(feas_xml)
    st.session_state["SRMA_HITS"] = df_feas

    diag2 = st.session_state.get("DIAGNOSTICS", {}) or {}
    diag2["feasibility"] = {"count": feas_total, "esearch_url": feas_url, "diag": feas_diag, "efetch_urls": feas_ef}
    st.session_state["DIAGNOSTICS"] = diag2

    # title/abstract screening (AI or rules)
    if df is None or df.empty:
        return

    recs = []
    for _, r in df.iterrows():
        recs.append(
            {
                "record_id": r.get("record_id", ""),
                "title": r.get("title", ""),
                "abstract": r.get("abstract", ""),
                "year": r.get("year", ""),
                "doi": r.get("doi", ""),
                "journal": r.get("journal", ""),
                "first_author": r.get("first_author", ""),
            }
        )

    lang = st.session_state.get("UI_LANG", "zh-TW")
    if llm_available():
        out_all: Dict[str, Dict[str, Any]] = {}
        for i in range(0, len(recs), 15):
            batch = recs[i : i + 15]
            try:
                out = llm_screen_title_abstract(batch, st.session_state["PROTOCOL"], lang)
            except Exception as e:
                out = {}
                d3 = st.session_state.get("DIAGNOSTICS", {}) or {}
                d3.setdefault("warnings", []).append(f"LLM screen failed on batch {i//15+1}: {e}")
                st.session_state["DIAGNOSTICS"] = d3
            out_all.update(out)

        for rid, v in out_all.items():
            st.session_state["TA_AI"][rid] = v.get("decision", "Unsure")
            st.session_state["TA_AI_REASON"][rid] = v.get("reason", "")
            st.session_state["TA_AI_CONF"][rid] = float(v.get("confidence", 0.5) or 0.5)
            st.session_state["TA_AI_RULES"][rid] = v.get("matched_rules", "")
    else:
        for r in recs:
            dec, conf, rs, rules = simple_rule_screen(r["title"], r["abstract"], st.session_state["PROTOCOL"])
            rid = r["record_id"]
            st.session_state["TA_AI"][rid] = dec
            st.session_state["TA_AI_REASON"][rid] = rs
            st.session_state["TA_AI_CONF"][rid] = conf
            st.session_state["TA_AI_RULES"][rid] = rules


# -------------------- Run (end-to-end) --------------------
if run:
    reset_downstream(keep_query=False)
    q = (st.session_state["QUESTION"] or "").strip()
    lang = st.session_state.get("UI_LANG", "zh-TW")

    with st.spinner("Step 0：建立 Protocol（PICO/criteria/schema）…"):
        if llm_available():
            try:
                proto = llm_build_protocol(q, lang)
            except Exception as e:
                proto = guess_pico_from_question(q)
                diag = st.session_state.get("DIAGNOSTICS", {}) or {}
                diag.setdefault("warnings", []).append(f"LLM protocol failed, fallback: {e}")
                st.session_state["DIAGNOSTICS"] = diag
        else:
            proto = guess_pico_from_question(q)

        proto.goal_mode = st.session_state["GOAL_MODE"]
        st.session_state["PROTOCOL"] = proto

    with st.spinner("Step 1：自動生成 PubMed 搜尋式（MeSH+free text）…"):
        try:
            proto = st.session_state["PROTOCOL"]
            proto.mesh_P = mesh_suggest(proto.P, 6) if proto.P else []
            proto.mesh_I = mesh_suggest(proto.I, 6) if proto.I else []
            proto.mesh_C = mesh_suggest(proto.C, 6) if proto.C else []
            proto.mesh_O = mesh_suggest(proto.O, 6) if proto.O else []
            st.session_state["PROTOCOL"] = proto
        except Exception:
            pass

        pub_q = build_pubmed_query(st.session_state["PROTOCOL"], st.session_state["ARTICLE_TYPE"], st.session_state["CUSTOM_PUBMED_FILTER"])
        st.session_state["PUBMED_QUERY_AUTO"] = pub_q
        st.session_state["PUBMED_QUERY_MANUAL"] = pub_q
        st.session_state["FEAS_QUERY"] = build_feas_query(st.session_state["PROTOCOL"])

    with st.spinner("Step 1-4：抓文獻、可行性掃描、Title/Abstract 粗篩…"):
        run_pubmed_pipeline(st.session_state["PUBMED_QUERY_MANUAL"])

    st.success("Done。請往下查看 Outputs。")


# =========================================================
# Outputs (Tabs)
# =========================================================
if (st.session_state.get("QUESTION") or "").strip():
    proto: Protocol = st.session_state.get("PROTOCOL", guess_pico_from_question(st.session_state["QUESTION"]))
    df = st.session_state.get("PUBMED_RECORDS", pd.DataFrame())
    df_feas = st.session_state.get("SRMA_HITS", pd.DataFrame())
    diag = st.session_state.get("DIAGNOSTICS", {}) or {}

    tabs = st.tabs(
        [
            t("tabs_overview"),
            t("tabs_step1"),
            t("tabs_step2"),
            t("tabs_step34"),
            t("tabs_ft"),
            t("tabs_extract"),
            t("tabs_ma"),
            t("tabs_rob2"),
            t("tabs_ms"),
            t("tabs_diag"),
        ]
    )

    # -------------------- Tab 0: Overview --------------------
    with tabs[0]:
        pr = compute_prisma(df) if df is not None else {}
        total = int(diag.get("pubmed_total_count", 0) or st.session_state.get("PUBMED_TOTAL", 0) or 0)
        feas_cnt = int((diag.get("feasibility", {}) or {}).get("count", 0) or 0)

        includes = unsure = excluded = 0
        if df is not None and not df.empty:
            for rid in df["record_id"].tolist():
                eff = compute_effective_ta_decision(rid)
                includes += (eff == "Include")
                unsure += (eff == "Unsure")
                excluded += (eff == "Exclude")

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f"<div class='kpi'><div class='label'>PubMed count</div><div class='value'>{total}</div></div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"<div class='kpi'><div class='label'>既有 SR/MA/NMA</div><div class='value'>{feas_cnt}</div></div>", unsafe_allow_html=True)
        with c3:
            st.markdown(f"<div class='kpi'><div class='label'>TA Include / Unsure</div><div class='value'>{includes} / {unsure}</div></div>", unsafe_allow_html=True)
        with c4:
            st.markdown(f"<div class='kpi'><div class='label'>LLM</div><div class='value'>{'ON' if llm_available() else 'OFF'}</div></div>", unsafe_allow_html=True)

        st.markdown("### Protocol（current）")
        st.code(pretty_json(proto.to_dict()), language="json")

        if df is None or df.empty:
            st.info(t("records_none"))
        else:
            render_prisma_text(pr)

    # -------------------- Tab 1: Step 1 (editable query) --------------------
    with tabs[1]:
        st.markdown("### Step 1：PubMed 搜尋式（可手動修改 + 重新抓 PubMed）")

        # Ensure query is initialized
        if not st.session_state.get("PUBMED_QUERY_MANUAL"):
            st.session_state["PUBMED_QUERY_MANUAL"] = st.session_state.get("PUBMED_QUERY_AUTO", "")

        st.text_area(
            t("pubmed_edit"),
            value=st.session_state.get("PUBMED_QUERY_MANUAL", ""),
            key="PUBMED_QUERY_MANUAL",
            height=160,
        )

        colA, colB, colC = st.columns([0.33, 0.33, 0.34])
        with colA:
            if st.button(t("pubmed_refetch")):
                with st.spinner("重新抓取 PubMed…"):
                    reset_downstream(keep_query=True)
                    run_pubmed_pipeline(st.session_state["PUBMED_QUERY_MANUAL"])
                st.success("已更新 records。")
                st.rerun()

        with colB:
            if st.button(t("pubmed_restore")):
                st.session_state["PUBMED_QUERY_MANUAL"] = st.session_state.get("PUBMED_QUERY_AUTO", "")
                st.success("已恢復為自動產生。")
                st.rerun()

        with colC:
            st.download_button(
                t("download_query"),
                data=(st.session_state.get("PUBMED_QUERY_MANUAL", "") or "").encode("utf-8"),
                file_name="pubmed_query.txt",
            )

        st.markdown("#### 自動生成（參考）")
        st.code(st.session_state.get("PUBMED_QUERY_AUTO", ""), language="text")

    # -------------------- Tab 2: Step 2 Feasibility --------------------
    with tabs[2]:
        st.markdown(f"### Step 2：{t('feas_title')}")
        st.markdown("**Feasibility query**")
        st.code(st.session_state.get("FEAS_QUERY", ""), language="text")

        feas = (diag.get("feasibility", {}) or {})
        feas_count = int(feas.get("count", 0) or 0)
        st.markdown(f"- SR/MA/NMA count：**{feas_count}**")

        if df_feas is not None and not df_feas.empty:
            show_cols = ["record_id", "year", "first_author", "journal", "title", "doi", "pmid"]
            df_show = ensure_columns(df_feas.copy(), show_cols, "")
            st.dataframe(df_show[show_cols], use_container_width=True, height=340)
        else:
            st.info("未抓到 SR/MA/NMA 命中（可能題目很窄、或 PubMed 回應受阻）。")

        # Rule-based综合建议（永遠存在）
        st.markdown("#### 綜合建議（自動；無 LLM 也會有）")
        total_count = int(st.session_state.get("PUBMED_TOTAL", 0) or 0)
        rec_lines = []

        if total_count == 0:
            rec_lines.append("目前 PubMed 召回為 0：建議把縮寫寫全名、加入品牌/型號或設計（diffractive/nondiffractive）、補上族群/術式/情境。")
            rec_lines.append("若你要做「不同種類 EDOF 比較」，請明確寫 A vs B（品牌或設計類型），避免 EDOF vs EDOF 這種太抽象的寫法。")
        else:
            if feas_count >= 5:
                rec_lines.append("已有多篇 SR/MA/NMA：建議先確認是否可做更新（update SR/MA）或縮小 PICO 以差異化。")
                rec_lines.append("可行縮題方向：限定族群（例如 OAG/ACG）、限定特定 IOL 型號/設計、限定追蹤時間點、或鎖定特定主要 outcome。")
                rec_lines.append("若要做 NMA：需要多介入節點且 outcome 定義一致；若只有 head-to-head 兩組，通常傳統 MA 足夠。")
            elif 1 <= feas_count < 5:
                rec_lines.append("已有少量 SR/MA/NMA：建議比對其納入研究與 outcome，找 gap（新 RCT、新鏡片、新設計、不同族群/追蹤）。")
                rec_lines.append("若目標偏『快速可行』：挑尚無 MA 的比較或既有 MA 未含最新 RCT 的比較。")
            else:
                rec_lines.append("目前未偵測到明顯的既有 SR/MA/NMA：可往下做，但仍建議先做 full-text feasibility（是否有足夠 RCT 與可統合 outcome）。")
                rec_lines.append("可先在 Step 3+4 看 abstracts 是否多為 trial-like；若多為非比較研究，需調整設計/納入標準。")

        rec_lines.append("學長要求重點：在開始所有步驟前，需先完成可行性掃描，再決定 PICO 劃定（偏可行性快速發 vs 研究嚴謹）。")

        for x in rec_lines:
            st.markdown(f"- {x}")

        # Optional LLM feasibility
        st.markdown(f"#### {t('feas_optional')}")
        if llm_available():
            if st.button("產生可行性報告（BYOK）"):
                with st.spinner("LLM 生成中…"):
                    hits = []
                    if df_feas is not None and not df_feas.empty:
                        for _, r in df_feas.head(15).iterrows():
                            hits.append(
                                {
                                    "pmid": r.get("pmid", ""),
                                    "year": r.get("year", ""),
                                    "first_author": r.get("first_author", ""),
                                    "journal": r.get("journal", ""),
                                    "title": r.get("title", ""),
                                    "doi": r.get("doi", ""),
                                }
                            )
                    rep = llm_feasibility_report(st.session_state["QUESTION"], proto, hits, st.session_state.get("UI_LANG", "zh-TW"))
                    st.code(pretty_json(rep), language="json")
        else:
            st.info("未啟用 LLM：不影響後續流程。")

    # -------------------- Tab 3: Step 3+4 combined screening --------------------
    with tabs[3]:
        st.markdown("### Step 3+4：Records + 粗篩（AI 保留 + 人工修正）")
        if df is None or df.empty:
            st.warning(t("records_none"))
        else:
            df = ensure_columns(df, ["record_id", "pmid", "pmcid", "doi", "year", "journal", "first_author", "title", "abstract", "source"], "")

            # Export table for records
            st.download_button(
                "下載 records（CSV）",
                data=to_csv_bytes(df),
                file_name="records.csv",
            )

            # Counters
            inc = exc = uns = 0
            for rid in df["record_id"].tolist():
                d0 = compute_effective_ta_decision(rid)
                if d0 == "Include":
                    inc += 1
                elif d0 == "Exclude":
                    exc += 1
                else:
                    uns += 1
            st.markdown(f"**Effective（含人工覆寫）**：Include **{inc}** ｜ Exclude **{exc}** ｜ Unsure **{uns}**")

            for _, r in df.iterrows():
                rid = r["record_id"]
                pmid = r.get("pmid", "")
                doi = r.get("doi", "")
                year = r.get("year", "")
                fa = r.get("first_author", "") or "—"
                journal = r.get("journal", "") or "—"
                title = r.get("title", "")
                abstract = r.get("abstract", "")
                pmcid = r.get("pmcid", "")

                pub_link = pubmed_link(pmid)
                doi_u = doi_link(doi)
                pmc_u = pmc_link(pmcid)
                openurl = build_openurl(st.session_state.get("RESOLVER_BASE", ""), doi=doi, pmid=pmid, title=title)
                openurl = apply_ezproxy(st.session_state.get("EZPROXY_PREFIX", ""), openurl) if openurl else ""

                ai_d = st.session_state["TA_AI"].get(rid, "Unsure")
                ai_r = st.session_state["TA_AI_REASON"].get(rid, "")
                ai_c = float(st.session_state["TA_AI_CONF"].get(rid, 0.5) or 0.5)
                ai_rules = st.session_state["TA_AI_RULES"].get(rid, "")
                eff = compute_effective_ta_decision(rid)

                with st.expander(f"{rid}｜{short(title, 110)}", expanded=False):
                    st.markdown(
                        f"<div class='card'>"
                        f"<div><b>{title}</b></div>"
                        f"<div class='small'>ID: PMID:{pmid or '—'}　|　DOI: {doi or '—'}　|　Year: {year or '—'}　|　First author: {fa}　|　Journal: {journal}</div>"
                        f"</div>",
                        unsafe_allow_html=True,
                    )
                    links = []
                    if pub_link:
                        links.append(f"[PubMed]({pub_link})")
                    if doi_u:
                        links.append(f"[DOI]({doi_u})")
                    if pmc_u:
                        links.append(f"[PMC OA]({pmc_u})")
                    if openurl:
                        links.append(f"[學院全文(OpenURL)]({openurl})")
                    if links:
                        st.markdown(" | ".join(links))

                    st.markdown(badge(eff) + "<span class='small'>　（Effective decision）</span>", unsafe_allow_html=True)
                    st.markdown(badge(ai_d) + "<span class='small'>　AI Title/Abstract 建議</span>", unsafe_allow_html=True)
                    st.write(f"理由：{ai_r or '—'}")
                    if ai_rules:
                        st.caption(f"Matched rules/keywords: {ai_rules}")
                    st.caption(f"信心度：{ai_c:.2f}")

                    st.markdown("#### 人工修正（override）")
                    cur = (st.session_state.get("TA_OVERRIDE", {}) or {}).get(rid, "")
                    opts = ["", "Include", "Exclude", "Unsure"]
                    new = st.selectbox("Override decision（留空=不覆蓋）", opts, index=opts.index(cur) if cur in opts else 0, key=f"ov_{rid}")
                    if new == "":
                        (st.session_state.get("TA_OVERRIDE", {}) or {}).pop(rid, None)
                    else:
                        st.session_state["TA_OVERRIDE"][rid] = new
                    st.session_state["TA_OVERRIDE_REASON"][rid] = st.text_area(
                        "Override reason（可留空）",
                        value=(st.session_state.get("TA_OVERRIDE_REASON", {}) or {}).get(rid, ""),
                        key=f"ov_r_{rid}",
                        height=70,
                    )

                    st.markdown("<hr class='hr'/>", unsafe_allow_html=True)
                    st.markdown("#### Abstract")
                    if abstract:
                        st.write(format_abstract(abstract))
                    else:
                        st.caption("_No abstract available._")

    # -------------------- Tab 4: Step 4b Full text review --------------------
    with tabs[4]:
        st.markdown("### Step 4b：Full text review（含排除理由 + PDF upload + 可選抽字）")
        if df is None or df.empty:
            st.info("尚無 records。")
        else:
            df = ensure_columns(df, ["record_id", "pmid", "doi", "year", "journal", "first_author", "title", "pmcid"], "")
            cand_ids = [rid for rid in df["record_id"].tolist() if compute_effective_ta_decision(rid) in ["Include", "Unsure"]]
            if not cand_ids:
                st.warning("沒有 TA Include/Unsure 的研究。")
            else:
                st.caption(
                    "提示：全文排除請填理由（會進 PRISMA/方法段落）；PDF 若為掃描檔請先 OCR。"
                    "若文章來自校內訂閱，避免上傳到公開部署；建議改本機版或僅上傳 OA/PMC。"
                )

                # Bulk upload
                pdfs = st.file_uploader(t("ft_bulk_upload"), type=["pdf"], accept_multiple_files=True, key="bulk_pdfs")
                if pdfs:
                    mapped = 0
                    for f in pdfs:
                        name = f.name or ""
                        # try match PMID-like digits
                        m = re.search(r"(?:PMID[:\-_ ]*)?(\d{7,8})", name)
                        pmid_guess = m.group(1) if m else ""
                        if pmid_guess:
                            rid_guess = f"PMID:{pmid_guess}"
                            if rid_guess in cand_ids and HAS_PYPDF2:
                                txt = extract_text_from_pdf(f)
                                if txt:
                                    st.session_state["FT_TEXT"][rid_guess] = txt
                                    mapped += 1
                    st.success(f"已嘗試自動對應並抽字：{mapped} 篇（若抽不到：可能無文字層，請先 OCR）")

                ft_opts = ["Not reviewed yet", "Include for meta-analysis", "Include (qualitative only)", "Exclude after full-text"]

                for rid in cand_ids:
                    row = df[df["record_id"] == rid].iloc[0]
                    title = row.get("title", "")
                    pmid = row.get("pmid", "")
                    doi = row.get("doi", "")
                    year = row.get("year", "")
                    fa = row.get("first_author", "") or "—"
                    journal = row.get("journal", "") or "—"
                    pmcid = row.get("pmcid", "")

                    pub_link0 = pubmed_link(pmid)
                    doi_u0 = doi_link(doi)
                    pmc_u0 = pmc_link(pmcid)
                    openurl = build_openurl(st.session_state.get("RESOLVER_BASE", ""), doi=doi, pmid=pmid, title=title)
                    openurl = apply_ezproxy(st.session_state.get("EZPROXY_PREFIX", ""), openurl) if openurl else ""

                    cur_ft = (st.session_state.get("FT_DECISIONS", {}) or {}).get(rid, "Not reviewed yet")
                    if cur_ft not in ft_opts:
                        cur_ft = "Not reviewed yet"

                    with st.expander(f"{rid}｜{short(title, 110)}", expanded=False):
                        st.markdown(
                            f"<div class='card'><div><b>{title}</b></div>"
                            f"<div class='small'>PMID:{pmid or '—'}　|　DOI:{doi or '—'}　|　Year:{year or '—'}　|　First author:{fa}　|　Journal:{journal}</div></div>",
                            unsafe_allow_html=True,
                        )
                        links = []
                        if pub_link0:
                            links.append(f"[PubMed]({pub_link0})")
                        if doi_u0:
                            links.append(f"[DOI]({doi_u0})")
                        if pmc_u0:
                            links.append(f"[PMC OA]({pmc_u0})")
                        if openurl:
                            links.append(f"[學院全文(OpenURL)]({openurl})")
                        if links:
                            st.markdown(" | ".join(links))

                        new_ft = st.radio("", ft_opts, index=ft_opts.index(cur_ft), key=f"ft_dec_{rid}", horizontal=True)
                        st.session_state["FT_DECISIONS"][rid] = new_ft

                        st.session_state["FT_REASONS"][rid] = st.text_area(
                            "排除/納入理由（建議必填；PRISMA 會用）",
                            value=(st.session_state.get("FT_REASONS", {}) or {}).get(rid, ""),
                            key=f"ft_reason_{rid}",
                            height=85,
                        )
                        st.session_state["FT_NOTE"][rid] = st.text_input(
                            "若查不到全文：填原因/狀態",
                            value=(st.session_state.get("FT_NOTE", {}) or {}).get(rid, ""),
                            key=f"ft_note_{rid}",
                        )

                        # Single upload
                        up = st.file_uploader(t("ft_single_upload"), type=["pdf"], key=f"ft_pdf_{rid}")
                        colx, coly = st.columns([0.35, 0.65])
                        with colx:
                            if st.button(t("ft_extract_text"), key=f"ft_ocr_{rid}", disabled=(up is None or not HAS_PYPDF2)):
                                txt = extract_text_from_pdf(up) if up is not None else ""
                                if txt:
                                    st.session_state["FT_TEXT"][rid] = txt
                                    st.success(f"已抽取全文文字，長度={len(txt)}")
                                else:
                                    st.warning("抽不到文字：可能是掃描 PDF 或無文字層。建議先 OCR 再上傳。")
                        with coly:
                            st.caption("若是掃描 PDF：請先用外部工具 OCR（Adobe/Drive OCR）後再上傳或貼上文字。")

                        st.text_area(
                            t("ft_text_area"),
                            value=(st.session_state.get("FT_TEXT", {}) or {}).get(rid, ""),
                            key=f"ft_text_{rid}",
                            height=240,
                        )
                        st.session_state["FT_TEXT"][rid] = st.session_state.get(f"ft_text_{rid}", "")

                        # Optional LLM: full-text review + extraction
                        if llm_available():
                            if st.button(t("ft_ai_fill"), key=f"ft_ai_{rid}"):
                                schema = st.session_state.get("SCHEMA_COLS", [])
                                js = llm_fulltext_extract_one(st.session_state["FT_TEXT"].get(rid, ""), proto, schema, st.session_state.get("UI_LANG", "zh-TW"))
                                st.session_state["AI_EXTRACTION_CACHE"][rid] = js

                                fd = (js.get("fulltext_decision") or "").strip()
                                fr = (js.get("fulltext_reason") or "").strip()
                                if fd in ft_opts:
                                    st.session_state["FT_DECISIONS"][rid] = fd
                                if fr:
                                    st.session_state["FT_REASONS"][rid] = fr

                                # Meta extraction auto-append (if provided)
                                meta = js.get("meta") or {}
                                if isinstance(meta, dict) and meta.get("effect_measure"):
                                    row_new = {
                                        "record_id": rid,
                                        "Outcome_label": str(meta.get("outcome_label", "") or ""),
                                        "Timepoint": str(meta.get("timepoint", "") or ""),
                                        "Effect_measure": str(meta.get("effect_measure", "") or ""),
                                        "Effect": meta.get("effect", ""),
                                        "Lower_CI": meta.get("lower_CI", ""),
                                        "Upper_CI": meta.get("upper_CI", ""),
                                        "Effect_unit": str(meta.get("effect_unit", "") or ""),
                                        "Notes": "AI extracted (verify manually)",
                                    }
                                    st.session_state["EXTRACT_DF"] = pd.concat([st.session_state["EXTRACT_DF"], pd.DataFrame([row_new])], ignore_index=True)

                                st.success("已回填（請人工核對）。")

                                # Optional: RoB2 suggestion from full text
                                with st.expander("（可選）AI 建議 ROB2（請人工核對）", expanded=False):
                                    robjs = llm_rob2_suggest(st.session_state["FT_TEXT"].get(rid, ""), st.session_state.get("UI_LANG", "zh-TW"))
                                    st.code(pretty_json(robjs), language="json")
                                    # Auto-fill into ROB2 state (non-destructive: only if empty)
                                    init_rob2_for_record(rid)
                                    rob = st.session_state.get("ROB2", {}) or {}
                                    for k, _ in ROB_DOMAINS:
                                        if k in robjs and isinstance(robjs[k], dict):
                                            if not rob[rid][k].get("level"):
                                                rob[rid][k]["level"] = str(robjs[k].get("level", "") or "")
                                            if not rob[rid][k].get("reason"):
                                                rob[rid][k]["reason"] = str(robjs[k].get("reason", "") or "")
                                    st.session_state["ROB2"] = rob
                        else:
                            st.info("未啟用 LLM：不影響後續流程。")

    # -------------------- Tab 5: Step 5 Extraction --------------------
    with tabs[5]:
        st.markdown("### Step 5：Data extraction（寬表；一次輸入完再寫入）")

        # Schema editor (PICO-level planning, not hard-coded)
        st.markdown(f"#### {t('extract_schema')}")
        st.text_area(
            "",
            value=st.session_state.get("SCHEMA_COLS_TEXT", ""),
            key="SCHEMA_COLS_TEXT",
            height=140,
            help="一行一個欄位名稱。建議包含：研究設計、族群、介入/比較、追蹤、主要/次要 outcomes、安全性、與 effect/CI 欄位。",
        )
        # Parse schema
        schema_lines = [x.strip() for x in (st.session_state.get("SCHEMA_COLS_TEXT") or "").splitlines() if x.strip()]
        # Ensure core columns exist for MA (still keep user-defined)
        core = ["Outcome_label", "Timepoint", "Effect_measure", "Effect", "Lower_CI", "Upper_CI", "Effect_unit", "Notes"]
        for c in core:
            if c not in schema_lines:
                schema_lines.append(c)
        st.session_state["SCHEMA_COLS"] = schema_lines

        st.markdown("#### （提示）Data extraction prompt（含 OCR/figure/table 提示）")
        st.info(
            "若你要用 AI 輔助 extraction：請在提示中明確要求「優先從 Table/Figure 抽取數值」；"
            "若 PDF 為掃描檔，請先 OCR；抽不到可留空，但要在 Notes 註記。"
        )

        # Determine eligible studies for extraction:
        # official: FT Include for meta-analysis; if none, show warning and allow user to still add manually.
        df_records = st.session_state.get("PUBMED_RECORDS", pd.DataFrame())
        df_records = ensure_columns(df_records, ["record_id", "title"], "")

        ft = st.session_state.get("FT_DECISIONS", {}) or {}
        include_meta = [rid for rid, v in ft.items() if v == "Include for meta-analysis"]
        if len(include_meta) == 0:
            st.warning("目前沒有 Full-text = Include for meta-analysis 的研究。你仍可先手動輸入寬表，但建議先在 Step 4b 完成全文決策。")
            include_meta = df_records["record_id"].tolist()[:50]  # allow selection anyway

        # Quick add form (prevents per-cell rerun annoyance)
        st.markdown(f"#### {t('extract_quick_add')}")
        with st.form("quick_add_row", clear_on_submit=True):
            rid = st.selectbox("record_id", options=include_meta)
            oc = st.text_input("Outcome_label（必填）")
            tp = st.text_input("Timepoint（可空）")
            meas = st.selectbox("Effect_measure（必填）", options=["OR", "RR", "HR", "MD", "SMD"])
            eff = st.text_input("Effect（必填；數值）")
            lcl = st.text_input("Lower CI（必填；數值）")
            ucl = st.text_input("Upper CI（必填；數值）")
            unit = st.text_input("Effect_unit（可空）")
            notes = st.text_area("Notes（可空）", height=70)
            ok = st.form_submit_button("新增到寬表")
            if ok:
                new = {
                    "record_id": rid,
                    "Outcome_label": oc.strip(),
                    "Timepoint": tp.strip(),
                    "Effect_measure": meas.strip(),
                    "Effect": eff.strip(),
                    "Lower_CI": lcl.strip(),
                    "Upper_CI": ucl.strip(),
                    "Effect_unit": unit.strip(),
                    "Notes": notes.strip(),
                }
                st.session_state["EXTRACT_DF"] = pd.concat([st.session_state["EXTRACT_DF"], pd.DataFrame([new])], ignore_index=True)
                st.success("已新增。")

        st.markdown("#### 寬表（目前）")
        st.dataframe(st.session_state["EXTRACT_DF"], use_container_width=True, height=320)
        st.download_button("下載 extraction（CSV）", data=to_csv_bytes(st.session_state["EXTRACT_DF"]), file_name="extraction_wide.csv")

        # Advanced editor: commit on save (avoids rerun confusion + KeyError)
        st.markdown(f"#### {t('extract_editor')}")
        if st.session_state["EXTRACT_EDITOR_DF"] is None:
            st.session_state["EXTRACT_EDITOR_DF"] = st.session_state["EXTRACT_DF"].copy()

        with st.form("extract_editor_form", clear_on_submit=False):
            edited = st.data_editor(
                st.session_state["EXTRACT_EDITOR_DF"],
                use_container_width=True,
                height=360,
                num_rows="dynamic",
            )
            save = st.form_submit_button(t("extract_save"))
            if save:
                # Commit safely; ensure columns exist
                edited = ensure_columns(edited, ["record_id", "Outcome_label", "Timepoint", "Effect_measure", "Effect", "Lower_CI", "Upper_CI", "Effect_unit", "Notes"], "")
                st.session_state["EXTRACT_DF"] = edited.copy()
                st.session_state["EXTRACT_EDITOR_DF"] = edited.copy()
                st.session_state["EXTRACT_COMMITTED_AT"] = time.time()
                st.success("已儲存/commit。")

        # Validation summary (red warnings; do not block)
        st.markdown("#### 輸入檢核（不會阻擋下一步；異常以紅字提示）")
        v = ensure_columns(st.session_state["EXTRACT_DF"].copy(), ["Effect", "Lower_CI", "Upper_CI", "Effect_measure", "Outcome_label"], "")
        issues = []
        for i, row in v.iterrows():
            meas = str(row.get("Effect_measure", "") or "").upper().strip()
            eff0 = safe_float(row.get("Effect"))
            l0 = safe_float(row.get("Lower_CI"))
            u0 = safe_float(row.get("Upper_CI"))
            if not (row.get("Outcome_label") or "").strip():
                issues.append((i, "Outcome_label", "缺 Outcome_label"))
            if meas not in ("OR", "RR", "HR", "MD", "SMD"):
                issues.append((i, "Effect_measure", "Effect_measure 必須是 OR/RR/HR/MD/SMD"))
            if eff0 is None or l0 is None or u0 is None:
                issues.append((i, "Effect/CI", "Effect/CI 必須是數值"))
            else:
                if l0 > u0:
                    issues.append((i, "CI", "Lower_CI 大於 Upper_CI"))
                if meas in ("OR", "RR", "HR") and (eff0 <= 0 or l0 <= 0 or u0 <= 0):
                    issues.append((i, "Effect/CI", "OR/RR/HR 需要 > 0（否則無法取 log）"))
        if issues:
            st.markdown("<span class='red'>發現以下問題（請修正後再跑 MA）：</span>", unsafe_allow_html=True)
            st.dataframe(pd.DataFrame(issues, columns=["row", "field", "issue"]), use_container_width=True, height=220)
        else:
            st.markdown("<span class='green'>未偵測到明顯輸入問題。</span>", unsafe_allow_html=True)

    # -------------------- Tab 6: Step 6 MA + Forest --------------------
    with tabs[6]:
        st.markdown("### Step 6：MA + 森林圖（fixed effect；按鈕執行）")
        st.caption("下方 outcome 與 measure 改成手動輸入；避免下拉造成 rerun 跳動。若資料不足，會顯示缺口與警告，不會 crash。")

        col1, col2, col3 = st.columns([0.45, 0.25, 0.30])
        with col1:
            st.text_input(t("ma_outcome_label"), value=st.session_state.get("MA_OUTCOME_FILTER", ""), key="MA_OUTCOME_FILTER")
        with col2:
            st.text_input(t("ma_measure"), value=st.session_state.get("MA_MEASURE", "OR"), key="MA_MEASURE")
        with col3:
            run_ma = st.button(t("ma_run"), type="primary")

        if run_ma:
            st.session_state["MA_WARNINGS"] = []
            df_ex = ensure_columns(st.session_state.get("EXTRACT_DF", pd.DataFrame()).copy(),
                                   ["record_id", "Outcome_label", "Effect_measure", "Effect", "Lower_CI", "Upper_CI", "Timepoint", "Notes"], "")

            outcome_filter = (st.session_state.get("MA_OUTCOME_FILTER") or "").strip()
            measure = (st.session_state.get("MA_MEASURE") or "").upper().strip()
            if measure not in ("OR", "RR", "HR", "MD", "SMD"):
                st.session_state["MA_WARNINGS"].append("Effect measure 不合法：請輸入 OR/RR/HR/MD/SMD")
                measure = "OR"

            # Filter rows
            if outcome_filter:
                df_use = df_ex[df_ex["Outcome_label"].astype(str).str.contains(re.escape(outcome_filter), case=False, na=False)].copy()
            else:
                df_use = df_ex.copy()

            if df_use.empty:
                st.session_state["MA_RESULT"] = {"ok": False, "error": "找不到符合 outcome 的列（或 extraction 空白）。"}
                st.session_state["MA_LAST_RUN"] = time.time()
            else:
                # If user measure conflicts with row measure, we still use user measure but warn.
                # Better: use row measure when it matches; else skip
                effects, ses, labels, lcls, ucls = [], [], [], [], []
                for _, r in df_use.iterrows():
                    row_meas = str(r.get("Effect_measure", "") or "").upper().strip()
                    eff0 = safe_float(r.get("Effect"))
                    l0 = safe_float(r.get("Lower_CI"))
                    u0 = safe_float(r.get("Upper_CI"))
                    if eff0 is None or l0 is None or u0 is None:
                        st.session_state["MA_WARNINGS"].append(f"{r.get('record_id','?')}: Effect/CI 非數值，已跳過。")
                        continue
                    if row_meas and row_meas != measure:
                        st.session_state["MA_WARNINGS"].append(f"{r.get('record_id','?')}: measure={row_meas} 與你指定的 {measure} 不一致，已跳過。")
                        continue
                    se = se_from_ci(eff0, l0, u0, measure)
                    if se is None or se <= 0:
                        st.session_state["MA_WARNINGS"].append(f"{r.get('record_id','?')}: 由 CI 推算 SE 失敗（可能 OR/RR/HR ≤ 0 或 CI 不合理），已跳過。")
                        continue
                    effects.append(float(eff0))
                    ses.append(float(se))
                    # label: author-year
                    rid = str(r.get("record_id", "") or "")
                    label = rid
                    # try map to author/year from records
                    try:
                        rec = st.session_state.get("PUBMED_RECORDS", pd.DataFrame())
                        if rec is not None and not rec.empty:
                            rr = rec[rec["record_id"] == rid]
                            if not rr.empty:
                                fa = rr.iloc[0].get("first_author", "") or rid
                                yy = rr.iloc[0].get("year", "") or ""
                                label = f"{fa} ({yy})"
                    except Exception:
                        pass
                    labels.append(label)
                    lcls.append(float(l0))
                    ucls.append(float(u0))

                res = fixed_effect_meta(effects, ses, measure) if len(effects) > 0 else {"ok": False, "error": "No valid rows after validation."}
                st.session_state["MA_RESULT"] = res
                st.session_state["MA_LAST_RUN"] = time.time()

                if res.get("ok"):
                    st.success(f"Pooled ({measure}, fixed): {res['pooled']:.4g}  (95% CI {res['lcl']:.4g} to {res['ucl']:.4g});  I²={res['I2']:.1f}% ; k={res['k']}")
                else:
                    st.error(res.get("error", "MA failed."))

                # Plot
                if res.get("ok"):
                    df_plot = pd.DataFrame({"label": labels, "effect": effects, "lcl": lcls, "ucl": ucls})
                    forest_plot(df_plot, res, measure)

        # Warnings
        warns = st.session_state.get("MA_WARNINGS", []) or []
        if warns:
            st.markdown("#### 警告/跳過原因（不會阻擋）")
            for w in warns:
                st.markdown(f"- <span class='red'>{html.escape(str(w))}</span>", unsafe_allow_html=True)

        # Show current MA result
        if st.session_state.get("MA_RESULT"):
            with st.expander("MA 結果（JSON）", expanded=False):
                st.code(pretty_json(st.session_state["MA_RESULT"]), language="json")

    # -------------------- Tab 7: Step 6b ROB2 --------------------
    with tabs[7]:
        st.markdown("### Step 6b：ROB 2.0（手動評分 + 理由；可選 AI 建議）")
        st.caption("ROB 2.0 通常在納入後做。本頁以 Full-text = Include for meta-analysis 為評估對象；若沒有，會提示先完成 Step 4b。")

        df_records = ensure_columns(st.session_state.get("PUBMED_RECORDS", pd.DataFrame()).copy(),
                                   ["record_id", "title", "first_author", "year"], "")
        ft = st.session_state.get("FT_DECISIONS", {}) or {}
        include_meta = [rid for rid, v in ft.items() if v == "Include for meta-analysis"]

        if len(include_meta) == 0:
            st.warning("目前沒有 Full-text = Include for meta-analysis 的研究；請先到 Step 4b 完成全文決策。")
        else:
            for rid in include_meta:
                init_rob2_for_record(rid)
                rr = df_records[df_records["record_id"] == rid]
                title = rr.iloc[0]["title"] if not rr.empty else rid
                fa = rr.iloc[0].get("first_author", "") if not rr.empty else ""
                yy = rr.iloc[0].get("year", "") if not rr.empty else ""
                header = f"{rid}｜{short(title, 90)}"
                if fa or yy:
                    header += f" ({fa} {yy})"

                with st.expander(header, expanded=False):
                    rob = st.session_state.get("ROB2", {}) or {}
                    for k, label in ROB_DOMAINS:
                        c1, c2 = st.columns([0.28, 0.72])
                        with c1:
                            level_key = f"rob_{rid}_{k}_level"
                            st.selectbox(label, ROB_LEVELS, index=ROB_LEVELS.index(rob[rid][k].get("level") or "NA") if (rob[rid][k].get("level") in ROB_LEVELS) else 0, key=level_key)
                            rob[rid][k]["level"] = st.session_state[level_key]
                        with c2:
                            reason_key = f"rob_{rid}_{k}_reason"
                            st.text_area("理由（建議必填）", value=rob[rid][k].get("reason", ""), key=reason_key, height=80)
                            rob[rid][k]["reason"] = st.session_state[reason_key]
                        st.markdown("<hr class='hr'/>", unsafe_allow_html=True)
                    st.session_state["ROB2"] = rob

            # summary table
            st.markdown("#### ROB2 Summary")
            rows = []
            rob = st.session_state.get("ROB2", {}) or {}
            for rid in include_meta:
                row = {"record_id": rid}
                for k, label in ROB_DOMAINS:
                    row[label] = (rob.get(rid, {}).get(k, {}) or {}).get("level", "")
                rows.append(row)
            st.dataframe(pd.DataFrame(rows), use_container_width=True, height=240)

    # -------------------- Tab 8: Step 7 Manuscript --------------------
    with tabs[8]:
        st.markdown("### Step 7：自動書寫稿件（分段呈現；缺失以『』占位）")
        st.caption("本草稿僅供學術用途，請逐句核對。若啟用 BYOK，可生成更完整版本，但仍需人工確認。")

        pr = compute_prisma(df) if df is not None else {}
        ma = st.session_state.get("MA_RESULT", {}) or {}

        # Build default draft (always available)
        PICO = proto.to_dict().get("pico", {})
        pooled_str = "『尚未完成 MA 或資料不足』"
        if ma.get("ok"):
            pooled_str = f"{ma.get('measure')} fixed pooled={ma.get('pooled'):.4g} (95% CI {ma.get('lcl'):.4g}–{ma.get('ucl'):.4g}), I²={ma.get('I2'):.1f}%, k={ma.get('k')}"

        intro = (
            "【Introduction】\n"
            "『研究背景：疾病/族群負擔』『介入/比較的臨床意義』『目前證據缺口（是否已有 SR/MA/NMA）』。\n"
            f"本研究旨在評估：P『{PICO.get('P','')}』；I『{PICO.get('I','')}』；"
            f"C『{PICO.get('C','')}』；O『{PICO.get('O','')}』。\n"
        )
        methods = (
            "【Methods】\n"
            "本研究依 PRISMA 指引進行系統性回顧與統合分析。\n"
            "資料庫：PubMed（必要時可增加其他資料庫）。\n"
            f"搜尋式（可人工修正）：『{(st.session_state.get('PUBMED_QUERY_MANUAL','') or '')[:300]}…』\n"
            "納入標準（PICO 層級）：『研究設計』『族群』『介入/比較』『追蹤時間』『Outcome』；"
            "排除：動物/體外/病例報告等。\n"
            "粗篩（Title/Abstract）後進行 Full-text review，排除研究需附理由。\n"
            "資料抽取：Extraction sheet 依題目與既有文獻（SR/MA/NMA）自行規劃；"
            "若全文為掃描 PDF，先 OCR；抽取 figure/table 時需明確提示。\n"
            "偏倚風險：以 RoB 2.0 逐 domain 評估並提供理由。\n"
            "統計：使用 inverse-variance fixed-effect；視需要可擴充 random-effects/敏感度分析。\n"
        )
        results = (
            "【Results】\n"
            f"PRISMA：Records identified={pr.get('records_identified','『』')}；"
            f"Full-text assessed={pr.get('fulltext_assessed','『』')}；"
            f"Studies included={pr.get('studies_included','『』')}；"
            f"Meta-analysis included={pr.get('included_meta','『』')}。\n"
            f"統合結果：{pooled_str}\n"
            "『次要結局』『安全性事件』『亞組/敏感度分析（若有）』。\n"
        )
        discussion = (
            "【Discussion】\n"
            "『主要發現與臨床意義』\n"
            "『與既有文獻一致/不一致』\n"
            "『可能機制』\n"
            "限制：『異質性』『偏倚』『資料缺失/定義不一致』『出版偏倚』。\n"
            "結論：『以主要 outcome 總結』；未來研究方向：『需要更多 RCT/一致的 outcome 報告』。\n"
        )

        st.text_area("Introduction", value=intro, height=180)
        st.text_area("Methods", value=methods, height=260)
        st.text_area("Results", value=results, height=200)
        st.text_area("Discussion", value=discussion, height=220)

        combined = "\n\n".join([intro, methods, results, discussion])
        st.text_area("整份草稿（可直接複製）", value=combined, height=520)

        # Optional LLM manuscript generation
        st.markdown(f"#### {t('ms_generate')}")
        if llm_available():
            if st.button("生成稿件（BYOK；請人工核對）"):
                with st.spinner("LLM 生成中…"):
                    lang = st.session_state.get("UI_LANG", "zh-TW")
                    sys = "You are a medical academic writer. Output in the requested language. Use placeholders 『』 if missing."
                    user = f"""
Language: {"Traditional Chinese" if lang=="zh-TW" else "English"}.
Write a structured SR/MA manuscript draft (Introduction, Methods, Results, Discussion).
Use the following protocol, PRISMA numbers, and MA results. If any detail is missing, use 『』 placeholders.

Protocol:
{pretty_json(proto.to_dict())}

PRISMA:
{pretty_json(pr)}

MA:
{pretty_json(ma)}

RoB2 summary keys available: {list((st.session_state.get("ROB2") or {}).keys())}

Also include:
- clear statement that AI-assisted steps were verified by humans (placeholder if unknown)
- do NOT fabricate citations; instead, output 『PMID:...』 placeholders if needed.

Return plain text with section headers.
""".strip()
                    out = llm_chat([{"role": "system", "content": sys}, {"role": "user", "content": user}], temperature=float(st.session_state.get("BYOK_TEMP", 0.2) or 0.2), timeout=240)
                    st.text_area("LLM 草稿（請人工核對）", value=out, height=700)
        else:
            st.info("未啟用 LLM：仍可使用上方模板草稿。")

        # Word export
        st.markdown("#### 匯出")
        if HAS_DOCX:
            if st.button(t("export_docx")):
                doc = Document()
                style = doc.styles["Normal"]
                style.font.name = "Times New Roman"
                style.font.size = Pt(11)

                doc.add_heading("Systematic Review / Meta-analysis Draft", level=1)
                doc.add_paragraph(f"Author: Ya Hsin Yao")
                doc.add_paragraph("Disclaimer: Academic use only. Verify all results and citations.")
                doc.add_paragraph("")

                doc.add_heading("Introduction", level=2)
                doc.add_paragraph(intro)
                doc.add_heading("Methods", level=2)
                doc.add_paragraph(methods)
                doc.add_heading("Results", level=2)
                doc.add_paragraph(results)
                doc.add_heading("Discussion", level=2)
                doc.add_paragraph(discussion)

                doc.add_heading("Protocol (JSON)", level=2)
                doc.add_paragraph(pretty_json(proto.to_dict()))
                doc.add_heading("PRISMA (JSON)", level=2)
                doc.add_paragraph(pretty_json(pr))
                doc.add_heading("MA (JSON)", level=2)
                doc.add_paragraph(pretty_json(ma))

                bio = io.BytesIO()
                doc.save(bio)
                st.download_button("下載 DOCX", data=bio.getvalue(), file_name="srma_draft.docx")
        else:
            st.info("環境未安裝 python-docx：無法匯出 Word。你可在 requirements.txt 加入 python-docx。")

    # -------------------- Tab 9: Diagnostics --------------------
    with tabs[9]:
        st.markdown("### Diagnostics（PubMed 被擋時非常重要）")
        st.code(pretty_json(st.session_state.get("DIAGNOSTICS", {})), language="json")
        st.markdown("#### 系統能力")
        st.write(
            {
                "PyPDF2": HAS_PYPDF2,
                "Plotly": HAS_PLOTLY,
                "Matplotlib": HAS_MPL,
                "python-docx": HAS_DOCX,
            }
        )

