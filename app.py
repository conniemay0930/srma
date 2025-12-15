
# app.py
# =========================================================
# 一句話帶你完成 Meta-analysis（繁體中文）
# Author: Ya Hsin Yao
#
# 免責聲明：本工具僅供學術研究/教學用途，不構成醫療建議或法律意見；
# 使用者須自行驗證所有結果、引用與全文內容；請勿上傳可識別之病人資訊。
#
# 校內資源/授權提醒（重要）：
# - 若文章來自校內訂閱（付費期刊/EZproxy/館藏系統），請勿將受版權保護之全文
#   上傳至任何第三方服務或公開部署之網站（包含本 app 的雲端部署）。
# - 請遵守圖書館授權條款：避免大量下載/自動化批次擷取、避免共享全文給未授權者。
# - 若不確定是否可上傳：建議改用「本機版」或僅上傳你有權分享的開放取用全文（OA/PMC）。
#
# Privacy notice (BYOK):
# - Key only used for this session; do not use on untrusted deployments.
# =========================================================

from __future__ import annotations

import io
import re
import math
import json
import time
import html
import uuid
import textwrap
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

import requests
import pandas as pd
import streamlit as st

# Optional plotting
try:
    import matplotlib.pyplot as plt
    from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
    HAS_MPL = True
except Exception:
    HAS_MPL = False

try:
    import plotly.graph_objects as go
    HAS_PLOTLY = True
except Exception:
    HAS_PLOTLY = False

# Optional docx export
try:
    from docx import Document
    from docx.shared import Pt
    HAS_DOCX = True
except Exception:
    HAS_DOCX = False


# =========================
# Page config + styles
# =========================
st.set_page_config(page_title="一句話帶你完成 Meta-analysis（繁體中文）", layout="wide")

CSS = """
<style>
:root{
  --bg:#ffffff;
  --muted:#6b7280;
  --line:#e5e7eb;
  --soft:#f7f7fb;
  --warn-bg:#fff7ed;
  --warn-line:#f59e0b;
  --ok-bg:#ecfdf5;
  --ok-line:#10b981;
  --bad-bg:#fef2f2;
  --bad-line:#ef4444;
  --brand:#0f172a;
}
.block-container { padding-top: 2rem; padding-bottom: 3rem; }
h1,h2,h3 { letter-spacing: -0.02em; color: var(--brand); }
.card { border:1px solid var(--line); border-radius:16px; padding: 0.95rem 1.05rem; background: var(--bg); margin-bottom: 0.9rem; box-shadow: 0 1px 0 rgba(0,0,0,0.03);}
.notice { border-left:5px solid var(--warn-line); background: var(--warn-bg); padding: 0.95rem 1.0rem; border-radius: 12px; margin-bottom: 0.9rem;}
.ok { border-left:5px solid var(--ok-line); background: var(--ok-bg); padding: 0.95rem 1.0rem; border-radius: 12px; margin-bottom: 0.9rem;}
.bad { border-left:5px solid var(--bad-line); background: var(--bad-bg); padding: 0.95rem 1.0rem; border-radius: 12px; margin-bottom: 0.9rem;}
.small { font-size: 0.9rem; color: var(--muted); }
.muted { color: var(--muted); }
.kpiwrap { display:flex; gap: 0.75rem; flex-wrap: wrap; }
.kpi { border:1px solid var(--line); border-radius:12px; padding:0.75rem 0.85rem; background: #f9fafb; min-width: 180px;}
.kpi .label { font-size: 0.82rem; color: var(--muted);}
.kpi .value { font-size: 1.25rem; font-weight: 800; color: #111827; }
.badge { display:inline-block; padding:0.15rem 0.55rem; border-radius:999px; font-size:0.78rem; margin-right: 0.35rem; border:1px solid rgba(0,0,0,0.06); background:#f3f4f6; }
.badge-include{ background:#d1fae5; color:#065f46; }
.badge-exclude{ background:#fee2e2; color:#991b1b; }
.badge-unsure{ background:#e0f2fe; color:#075985; }
.badge-warn{ background:#fff7ed; color:#92400e; }
.red { color: var(--bad-line); font-weight: 700; }
pre code { white-space: pre-wrap !important; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


# =========================
# i18n (minimal)
# =========================
LANGS = {
    "繁體中文": "zh",
    "English": "en",
}

T = {
    "zh": {
        "title": "一句話帶你完成 Meta-analysis",
        "author": "作者：Ya Hsin Yao",
        "disclaimer": "免責聲明：僅供學術用途；請自行驗證所有結果與引用；請勿上傳可識別病人資訊；使用校內資源下載/全文請遵守授權（避免濫用）。",
        "byok_title": "LLM（使用者自備 key）",
        "byok_enable": "啟用 LLM（BYOK）",
        "byok_hint": "Key only used for this session；不寫入 secrets、不落盤。不要在不可信部署使用；不要上傳可識別病人資訊。",
        "byok_consent": "我理解並同意：僅供學術用途；輸出需人工核對；不輸入病人資訊；不違反校內授權。",
        "byok_clear": "清除本次 session 的 key",
        "run": "Run（one-click pipeline）",
        "question": "研究問題（可一句話；例：FLACS 是否比傳統 phaco 好？）",
        "studytype": "文章類型 filter",
        "goal": "目標模式",
        "max_records": "每個資料庫抓取文獻數量上限",
        "delay": "Polite delay（秒）",
        "outputs": "輸出",
        "diag": "Diagnostics（很重要：PubMed 被擋時用）",
        "no_llm": "未設定 LLM：此區將用規則法/模板法，並自動降級（不會卡住）。",
        "pubmed_query": "PubMed 搜尋式（可直接手動改）",
        "rebuild_query": "套用：重建自動搜尋式",
        "feasibility": "可行性掃描（既有 SR/MA/NMA）",
        "records": "文獻列表（可展開看摘要；可 Override）",
        "screening": "粗篩 + Full text review（合併）",
        "extraction": "資料萃取（寬表）",
        "ma": "MA + 森林圖（RevMan-like）",
        "manuscript": "稿件草稿（分段呈現；可匯出）",
        "prisma": "PRISMA 流程（圖/文字）",
    },
    "en": {
        "title": "One-sentence to Meta-analysis",
        "author": "Author: Ya Hsin Yao",
        "disclaimer": "Academic use only; verify all results and citations; do not upload identifiable patient info; respect institutional license/subscription terms.",
        "byok_title": "LLM (Bring Your Own Key)",
        "byok_enable": "Enable LLM (BYOK)",
        "byok_hint": "Key only used for this session; not stored. Do not use on untrusted deployments; do not upload identifiable patient info.",
        "byok_consent": "I understand: academic use only; outputs need human verification; no patient identifiers; comply with institutional license.",
        "byok_clear": "Clear session key",
        "run": "Run (one-click pipeline)",
        "question": "Research question (one sentence)",
        "studytype": "Study-type filter",
        "goal": "Goal mode",
        "max_records": "Max records per database",
        "delay": "Polite delay (sec)",
        "outputs": "Outputs",
        "diag": "Diagnostics (important if PubMed blocked)",
        "no_llm": "LLM not configured: using rule/template fallback; auto-degrade (won't get stuck).",
        "pubmed_query": "PubMed query (editable)",
        "rebuild_query": "Apply: rebuild auto query",
        "feasibility": "Feasibility scan (existing SR/MA/NMA)",
        "records": "Records (expand for abstract; override allowed)",
        "screening": "Screening + full-text review (merged)",
        "extraction": "Data extraction (wide table)",
        "ma": "Meta-analysis + forest plot",
        "manuscript": "Manuscript draft (sections shown; export optional)",
        "prisma": "PRISMA flow",
    },
}


# =========================
# Helpers
# =========================
def norm_text(x: Any) -> str:
    if x is None:
        return ""
    x = html.unescape(str(x))
    x = re.sub(r"\s+", " ", x).strip()
    return x

def short(s: str, n: int = 140) -> str:
    s = s or ""
    return (s[:n] + "…") if len(s) > n else s

def ensure_columns(df: pd.DataFrame, cols: List[str], default: Any = "") -> pd.DataFrame:
    for c in cols:
        if c not in df.columns:
            df[c] = default
    return df


def contains_cjk(text: str) -> bool:
    if not text:
        return False
    return any("\u4e00" <= ch <= "\u9fff" or "\u3400" <= ch <= "\u4dbf" for ch in text)

def warn_if_non_english(label: str, text: str):
    # PubMed is primarily indexed in English; Chinese terms in query typically reduce recall.
    if contains_cjk(text):
        st.warning(f"偵測到「{label}」包含中文/非英文字元；PubMed 命中率可能很低。建議在下方『人工修正 PICO（英文檢索用）』改成英文或加入英文同義詞。")

def maybe_translate_to_english(text: str) -> str:
    """Translate to English if CJK is detected and BYOK is enabled; otherwise return original."""
    text = (text or "").strip()
    if not text:
        return ""
    if not contains_cjk(text):
        return text
    if not llm_available():
        return text
    try:
        msg = [
            {"role": "system", "content": "You are a professional medical translator. Output ONLY the translated English text, no quotes."},
            {"role": "user", "content": text},
        ]
        out = call_llm(msg, max_tokens=220)
        return (out or "").strip()
    except Exception:
        return text

def safe_float(x: Any) -> Optional[float]:
    try:
        if x is None or (isinstance(x, float) and math.isnan(x)):
            return None
        s = str(x).strip()
        if s == "" or s.lower() in {"nan", "none"}:
            return None
        return float(s)
    except Exception:
        return None

def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")

def doi_to_url(doi: str) -> str:
    doi = (doi or "").strip()
    return f"https://doi.org/{doi}" if doi else ""

def pubmed_url(pmid: str) -> str:
    pmid = (pmid or "").strip()
    return f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/" if pmid else ""


# =========================
# Session state init
# =========================
def init_state() -> None:
    ss = st.session_state
    ss.setdefault("lang", "繁體中文")
    ss.setdefault("byok_enabled", False)
    ss.setdefault("byok_key", "")
    ss.setdefault("byok_base_url", "https://api.openai.com/v1")
    ss.setdefault("byok_model", "gpt-4o-mini")
    ss.setdefault("byok_temp", 0.2)
    ss.setdefault("byok_consent", False)

    ss.setdefault("question", "")
    ss.setdefault("study_type", "RCT")
    ss.setdefault("goal_mode", "Rigorous / comprehensive（嚴謹）")
    ss.setdefault("max_records", 300)
    ss.setdefault("polite_delay", 0.25)

    ss.setdefault("protocol", {})
    ss.setdefault("pubmed_query", "")
    ss.setdefault("records_df", pd.DataFrame())
    ss.setdefault("ta_ai", {})              # record_id -> dict (decision/reason/conf)
    ss.setdefault("ta_override", {})        # record_id -> decision
    ss.setdefault("ta_override_reason", {}) # record_id -> reason text

    ss.setdefault("ft_text", {})            # record_id -> pasted full text (or extracted)
    ss.setdefault("ft_decisions", {})       # record_id -> decision
    ss.setdefault("ft_reasons", {})         # record_id -> reason

    ss.setdefault("schema_lines", "Primary outcome\nSecondary outcome 1\nSecondary outcome 2")
    ss.setdefault("criteria_text", "")
    ss.setdefault("extract_df", pd.DataFrame())
    ss.setdefault("extract_editor_buf", pd.DataFrame())
    ss.setdefault("extract_last_saved", 0.0)

    ss.setdefault("ma_outcome_input", "")
    ss.setdefault("ma_measure_choice", "")
    ss.setdefault("ma_model_choice", "Fixed effect")
    ss.setdefault("ma_last_run", 0.0)
    ss.setdefault("ma_result", {})
    ss.setdefault("ms_sections", {})        # manuscript sections strings

init_state()


def tr(key: str) -> str:
    lang_code = LANGS.get(st.session_state.get("lang", "繁體中文"), "zh")
    return T.get(lang_code, T["zh"]).get(key, key)


# =========================
# BYOK LLM
# =========================
def llm_available() -> bool:
    return bool(st.session_state.get("byok_enabled")) and bool(st.session_state.get("byok_key", "").strip()) and bool(st.session_state.get("byok_consent"))

def call_llm(messages: List[Dict[str, str]], max_tokens: int = 1400, timeout: int = 75) -> str:
    """
    OpenAI-compatible: POST {base}/chat/completions  (for compatibility with many gateways)
    - Streamlit Cloud users: do NOT store keys in code.
    """
    base_url = (st.session_state.get("byok_base_url") or "").strip().rstrip("/")
    api_key = (st.session_state.get("byok_key") or "").strip()
    model = (st.session_state.get("byok_model") or "").strip()
    temperature = float(st.session_state.get("byok_temp") or 0.2)

    if not (base_url and api_key and model):
        raise RuntimeError("LLM 未設定完成（base_url / key / model）。")

    # try both paths for compatibility
    url_candidates = [
        f"{base_url}/chat/completions",
        f"{base_url}/v1/chat/completions",
    ]
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {"model": model, "messages": messages, "temperature": temperature, "max_tokens": max_tokens}

    last_err = None
    for url in url_candidates:
        try:
            r = requests.post(url, headers=headers, json=payload, timeout=timeout)
            if r.status_code != 200:
                last_err = f"HTTP {r.status_code}: {r.text[:400]}"
                continue
            data = r.json()
            return data["choices"][0]["message"]["content"]
        except Exception as e:
            last_err = str(e)
            continue
    raise RuntimeError(f"LLM 呼叫失敗：{last_err}")


# =========================
# Question parsing (no-LLM fallback)
# =========================
def parse_question_rule(q: str) -> Dict[str, str]:
    """
    Heuristic: if "A vs B" → I=A, C=B, P empty unless "in/among" tail suggests population.
    Else: set I=q and leave others blank.
    """
    q0 = norm_text(q)
    pico = {"P": "", "I": "", "C": "", "O": "", "NOT": "animal OR mice OR rat OR in vitro OR case report"}
    if not q0:
        return pico

    q_low = q0.lower()

    # try to split population: "... in cataract patients"
    pop = ""
    m = re.search(r"\b(in|among|within)\b\s+(.+)$", q_low)
    if m:
        # keep original casing from q0 by slicing
        idx = m.start()
        pop = q0[idx + len(m.group(1)):].strip()
        # the head part is before "in/among/within"
        q_head = q0[:idx].strip()
    else:
        q_head = q0

    # compare forms
    vs_pat = re.compile(r"\b(vs\.?|versus|compare(d)?\s+with)\b", re.I)
    if vs_pat.search(q_head):
        parts = vs_pat.split(q_head)
        left = parts[0].strip(" :,-")
        # right side is after the matched token; attempt to take last chunk
        right = parts[-1].strip(" :,-")
        if left and right:
            pico["I"] = left
            pico["C"] = right
        else:
            pico["I"] = q0
    else:
        pico["I"] = q_head

    if pop:
        pico["P"] = pop

    return pico


# =========================
# PubMed E-utilities
# =========================
NCBI_ESEARCH = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
NCBI_EFETCH = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"

def pubmed_esearch(term: str, retmax: int = 200, retstart: int = 0) -> Tuple[int, List[str], str, List[str]]:
    """
    returns: (count, idlist, url, warnings)
    warnings: list of human-readable messages if response seems blocked/HTML
    """
    params = {"db": "pubmed", "term": term, "retmode": "json", "retmax": retmax, "retstart": retstart}
    r = requests.get(NCBI_ESEARCH, params=params, timeout=60)
    txt = r.text or ""
    warnings: List[str] = []
    if "<html" in txt.lower():
        warnings.append("PubMed 回傳 HTML（可能被網路環境擋住）")
        return 0, [], r.url, warnings
    js = r.json()
    count = int(js["esearchresult"]["count"])
    ids = js["esearchresult"].get("idlist", []) or []
    return count, ids, r.url, warnings

def _etree_from_xml(xml_text: str):
    import xml.etree.ElementTree as ET
    return ET.fromstring(xml_text)

def pubmed_efetch(pmids: List[str]) -> Tuple[pd.DataFrame, List[str]]:
    """
    Fetch PubMed records for a list of PMIDs, parse minimal fields.
    """
    if not pmids:
        return pd.DataFrame(), []
    params = {"db": "pubmed", "retmode": "xml", "id": ",".join(pmids)}
    r = requests.get(NCBI_EFETCH, params=params, timeout=90)
    txt = r.text or ""
    warnings: List[str] = []
    if "<html" in txt.lower():
        warnings.append("efetch 回傳 HTML（可能被擋住）")
        return pd.DataFrame(), warnings

    try:
        root = _etree_from_xml(txt)
    except Exception as e:
        warnings.append(f"XML 解析失敗：{e}")
        return pd.DataFrame(), warnings

    rows: List[Dict[str, Any]] = []
    for art in root.findall(".//PubmedArticle"):
        try:
            pmid = (art.findtext(".//MedlineCitation/PMID") or "").strip()
            title = norm_text(art.findtext(".//Article/ArticleTitle") or "")
            # abstract: multiple AbstractText nodes; keep line breaks
            abs_nodes = art.findall(".//Article/Abstract/AbstractText")
            abs_parts = []
            for n in abs_nodes:
                label = n.attrib.get("Label") or n.attrib.get("NlmCategory") or ""
                seg = norm_text("".join(n.itertext()))
                if seg:
                    abs_parts.append((f"{label}: {seg}" if label else seg))
            abstract = "\n\n".join(abs_parts).strip()

            journal = norm_text(art.findtext(".//Article/Journal/Title") or art.findtext(".//Article/Journal/ISOAbbreviation") or "")
            year = ""
            for path in [".//Article/Journal/JournalIssue/PubDate/Year",
                         ".//Article/Journal/JournalIssue/PubDate/MedlineDate",
                         ".//ArticleDate/Year"]:
                v = (art.findtext(path) or "").strip()
                if v:
                    year = re.findall(r"\d{4}", v)[0] if re.findall(r"\d{4}", v) else v[:4]
                    break

            # authors
            authors = []
            for au in art.findall(".//Article/AuthorList/Author"):
                ln = (au.findtext("LastName") or "").strip()
                inits = (au.findtext("Initials") or "").strip()
                coll = (au.findtext("CollectiveName") or "").strip()
                if coll:
                    authors.append(coll)
                elif ln:
                    authors.append(f"{ln} {inits}".strip())
            first_author = authors[0].split()[0] if authors else ""

            # DOI
            doi = ""
            for aid in art.findall(".//ArticleIdList/ArticleId"):
                if aid.attrib.get("IdType") == "doi":
                    doi = (aid.text or "").strip()
                    break

            # PMC
            pmc = ""
            for aid in art.findall(".//ArticleIdList/ArticleId"):
                if aid.attrib.get("IdType") == "pmc":
                    pmc = (aid.text or "").strip()
                    break

            rid = f"PMID:{pmid}" if pmid else f"RID:{uuid.uuid4().hex[:10]}"
            rows.append({
                "record_id": rid,
                "pmid": pmid,
                "doi": doi,
                "pmc": pmc,
                "year": year,
                "journal": journal,
                "first_author": first_author,
                "title": title,
                "abstract": abstract,
                "source": "PubMed",
                "pubmed_url": pubmed_url(pmid),
                "doi_url": doi_to_url(doi),
                "pmc_url": f"https://pmc.ncbi.nlm.nih.gov/articles/{pmc}/" if pmc else "",
            })
        except Exception:
            continue

    return pd.DataFrame(rows), warnings


# =========================
# Query builder
# =========================
def build_pubmed_block(free_text: str, mesh: str = "", synonyms: Optional[List[str]] = None) -> str:
    terms = []
    ft = (free_text or "").strip()
    if ft:
        # if user provides multiple tokens separated by comma/semicolon
        toks = [t.strip() for t in re.split(r"[;,/]+", ft) if t.strip()]
        if len(toks) == 1:
            terms.append(f'("{toks[0]}"[tiab])')
        else:
            # OR list
            terms.append("(" + " OR ".join([f'("{t}"[tiab])' for t in toks]) + ")")
    if synonyms:
        syn = [s.strip() for s in synonyms if s and s.strip()]
        if syn:
            terms.append("(" + " OR ".join([f'("{s}"[tiab])' for s in syn]) + ")")
    if mesh.strip():
        terms.append(f'("{mesh}"[MeSH Terms])')
    if not terms:
        return ""
    if len(terms) == 1:
        return terms[0]
    return "(" + " OR ".join(terms) + ")"

def build_rct_filter() -> str:
    # broad but common
    return '(randomized controlled trial[tiab] OR randomised[tiab] OR randomized[tiab] OR trial[tiab] OR "Randomized Controlled Trial"[Publication Type])'

def build_srma_filter() -> str:
    return '("systematic review"[Publication Type] OR "meta-analysis"[Publication Type] OR systematic[tiab] OR meta-analysis[tiab] OR "network meta-analysis"[tiab] OR NMA[tiab])'

def build_pubmed_query(pico: Dict[str, str], include_CO: bool, add_rct: bool, extra: str, not_terms: str) -> str:
    P = pico.get("P", "")
    I = pico.get("I", "")
    C = pico.get("C", "")
    O = pico.get("O", "")

    parts = []
    if P.strip(): parts.append(build_pubmed_block(P))
    if I.strip(): parts.append(build_pubmed_block(I))
    if include_CO:
        if C.strip(): parts.append(build_pubmed_block(C))
        if O.strip(): parts.append(build_pubmed_block(O))
    if extra.strip(): parts.append(build_pubmed_block(extra))
    if add_rct:
        parts.append(build_rct_filter())

    base = " AND ".join([p for p in parts if p]).strip()
    if base and not_terms.strip():
        return f"({base}) NOT ({not_terms})"
    if base:
        return base
    if not_terms.strip():
        return f"NOT ({not_terms})"
    return ""


# =========================
# Rule-based screening fallback
# =========================
def _count_hits(text_low: str, term: str) -> int:
    term = (term or "").strip()
    if not term:
        return 0
    toks = [t.strip().lower() for t in re.split(r"[,\s;/]+", term) if t.strip()]
    return sum(1 for t in toks if t and t in text_low)

def ai_screen_rule_based(row: pd.Series, pico: Dict[str, str], study_filter: str) -> Dict[str, Any]:
    title = row.get("title", "") or ""
    abstract = row.get("abstract", "") or ""
    text = (title + " " + abstract).lower()

    P = pico.get("P", "") or ""
    I = pico.get("I", "") or ""
    C = pico.get("C", "") or ""
    O = pico.get("O", "") or ""
    NOT = pico.get("NOT", "") or ""

    # filter NOT terms
    not_hit = _count_hits(text, NOT)

    p_hit = _count_hits(text, P)
    i_hit = _count_hits(text, I)
    c_hit = _count_hits(text, C)
    o_hit = _count_hits(text, O)

    is_trial_like = any(w in text for w in ["randomized", "randomised", "trial", "controlled", "prospective"])
    is_srma_like = any(w in text for w in ["systematic", "meta-analysis", "network meta-analysis", "nma"])

    # study type preference
    prefer_rct = (study_filter or "").upper().startswith("RCT")

    # score
    keys = [k for k, v in [("P", P), ("I", I), ("C", C), ("O", O)] if v.strip()]
    denom = max(len(keys), 1)
    score = sum(1 for k in keys if {"P": p_hit, "I": i_hit, "C": c_hit, "O": o_hit}[k] > 0) / denom
    conf = min(0.95, 0.35 + 0.55 * score + (0.1 if is_trial_like else 0.0))

    reason_bits = []
    if p_hit: reason_bits.append(f"P hit×{p_hit}")
    if i_hit: reason_bits.append(f"I hit×{i_hit}")
    if c_hit: reason_bits.append(f"C hit×{c_hit}")
    if o_hit: reason_bits.append(f"O hit×{o_hit}")
    if prefer_rct and is_trial_like: reason_bits.append("Trial-like")
    if not_hit: reason_bits.append("NOT-term hit")

    # decision
    if not_hit:
        dec = "Exclude"
    elif prefer_rct and not is_trial_like and not is_srma_like:
        # not obviously trial-like: keep as unsure rather than hard exclude
        dec = "Unsure" if score >= 0.33 else "Exclude"
    else:
        # inclusive rule
        if i_hit or c_hit:
            dec = "Include" if (score >= 0.34 or is_trial_like) else "Unsure"
        else:
            dec = "Unsure" if score >= 0.5 else "Exclude"

    return {"decision": dec, "confidence": float(conf), "reason": "; ".join(reason_bits) or "Rule-based: limited evidence in title/abstract."}


def ai_screen_llm(row: pd.Series, proto: Dict[str, Any], study_filter: str) -> Dict[str, Any]:
    """
    LLM-based title/abstract screening. Returns decision/reason/confidence.
    """
    title = row.get("title", "") or ""
    abstract = row.get("abstract", "") or ""
    pico = proto.get("pico", {}) if isinstance(proto, dict) else {}
    goal = proto.get("goal_mode", "")
    msg = [
        {"role": "system", "content": "You are a meticulous evidence-synthesis assistant. Output JSON only."},
        {"role": "user", "content": f"""
Task: title/abstract screening suggestion for a systematic review / meta-analysis.

Constraints:
- Follow PICO as best as possible; if unclear, return Unsure.
- If study-type filter suggests RCT, prioritize RCT/controlled trials; observational may be Unsure.
- Provide: decision in [Include, Exclude, Unsure], reason (one paragraph), confidence 0-1.

Goal mode: {goal}
Study-type filter: {study_filter}

PICO:
P: {pico.get("P","")}
I: {pico.get("I","")}
C: {pico.get("C","")}
O: {pico.get("O","")}
NOT: {pico.get("NOT","")}

Title: {title}

Abstract:
{abstract[:4000]}
"""}
    ]
    out = call_llm(msg, max_tokens=450)
    # strict JSON parse with fallback
    try:
        js = json.loads(out)
        dec = str(js.get("decision", "Unsure")).strip()
        if dec not in ["Include", "Exclude", "Unsure"]:
            dec = "Unsure"
        conf = float(js.get("confidence", 0.5))
        conf = max(0.0, min(1.0, conf))
        reason = str(js.get("reason", "")).strip()
        return {"decision": dec, "confidence": conf, "reason": reason}
    except Exception:
        return {"decision": "Unsure", "confidence": 0.45, "reason": "LLM 回覆非 JSON；已降級為 Unsure。"}

# =========================
# Feasibility scan
# =========================
def feasibility_query(base_query: str, not_terms: str) -> str:
    # Use the base (P+I+extra) but swap filter to SR/MA/NMA.
    parts = []
    if base_query.strip():
        parts.append(f"({base_query})")
    parts.append(build_srma_filter())
    q = " AND ".join(parts)
    if not_terms.strip():
        q = f"({q}) NOT ({not_terms})"
    return q


# =========================
# Meta-analysis
# =========================
def se_from_ci(effect: float, lcl: float, ucl: float, measure: str) -> Optional[float]:
    """
    For ratio measures, CI assumed on ratio scale → log-transform.
    For MD/SMD, CI on linear scale.
    """
    measure = (measure or "").upper().strip()
    if any(v is None for v in [effect, lcl, ucl]):
        return None
    if measure in {"OR", "RR", "HR"}:
        if effect <= 0 or lcl <= 0 or ucl <= 0:
            return None
        return (math.log(ucl) - math.log(lcl)) / (2 * 1.96)
    # MD/SMD
    return (ucl - lcl) / (2 * 1.96)

def fixed_effect_pool(effects: List[float], ses: List[float]) -> Dict[str, float]:
    w = []
    for s in ses:
        if s is None or s <= 0:
            w.append(0.0)
        else:
            w.append(1.0 / (s * s))
    sw = sum(w)
    if sw <= 0:
        raise ValueError("No valid weights.")
    pooled = sum(wi * xi for wi, xi in zip(w, effects)) / sw
    se_pool = math.sqrt(1.0 / sw)
    return {"pooled": pooled, "se": se_pool, "weights": w}

def dersimonian_laird_tau2(effects: List[float], ses: List[float]) -> float:
    w = [0.0 if (s is None or s <= 0) else 1.0/(s*s) for s in ses]
    sw = sum(w)
    if sw <= 0:
        return 0.0
    mu = sum(wi*xi for wi, xi in zip(w, effects))/sw
    q = sum(wi*((xi-mu)**2) for wi, xi in zip(w, effects))
    df = max(len([wi for wi in w if wi>0]) - 1, 1)
    c = sw - (sum(wi*wi for wi in w)/sw if sw>0 else 0)
    if c <= 0:
        return 0.0
    tau2 = max(0.0, (q - df)/c)
    return tau2

def random_effect_pool(effects: List[float], ses: List[float]) -> Dict[str, float]:
    tau2 = dersimonian_laird_tau2(effects, ses)
    w = []
    for s in ses:
        if s is None or s <= 0:
            w.append(0.0)
        else:
            w.append(1.0 / (s*s + tau2))
    sw = sum(w)
    if sw <= 0:
        raise ValueError("No valid weights (random).")
    pooled = sum(wi*xi for wi, xi in zip(w, effects))/sw
    se_pool = math.sqrt(1.0/sw)
    return {"pooled": pooled, "se": se_pool, "weights": w, "tau2": tau2}

def i2_stat(effects: List[float], ses: List[float]) -> float:
    w = [0.0 if (s is None or s <= 0) else 1.0/(s*s) for s in ses]
    sw = sum(w)
    if sw <= 0:
        return 0.0
    mu = sum(wi*xi for wi, xi in zip(w, effects))/sw
    q = sum(wi*((xi-mu)**2) for wi, xi in zip(w, effects))
    df = max(len([wi for wi in w if wi>0]) - 1, 1)
    if q <= df:
        return 0.0
    return max(0.0, (q - df)/q) * 100.0

def transform_effect(effect: float, lcl: float, ucl: float, measure: str) -> Tuple[float, float, float]:
    """
    Convert to analysis scale:
    - OR/RR/HR -> log scale
    - MD/SMD -> identity
    """
    measure = (measure or "").upper().strip()
    if measure in {"OR","RR","HR"}:
        return math.log(effect), math.log(lcl), math.log(ucl)
    return effect, lcl, ucl

def back_transform(x: float, measure: str) -> float:
    measure = (measure or "").upper().strip()
    if measure in {"OR","RR","HR"}:
        return math.exp(x)
    return x


# =========================
# PRISMA (matplotlib)
# =========================
def draw_prisma(counts: Dict[str, int]):
    if not HAS_MPL:
        return None
    # Simple PRISMA-like boxes
    fig, ax = plt.subplots(figsize=(8.5, 6.2))
    ax.axis("off")

    def box(x, y, w, h, text):
        p = FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.02",
                           linewidth=1, edgecolor="#cbd5e1", facecolor="#ffffff")
        ax.add_patch(p)
        ax.text(x + w/2, y + h/2, text, ha="center", va="center", fontsize=10, wrap=True)

    def arrow(x1, y1, x2, y2):
        a = FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="-|>", mutation_scale=12, linewidth=1, color="#64748b")
        ax.add_patch(a)

    # Coordinates (0..1)
    box(0.12, 0.78, 0.76, 0.14, f"識別\n資料庫檢索得到：n = {counts.get('identified',0)}")
    box(0.12, 0.58, 0.76, 0.14, f"篩選\nTitle/Abstract 篩選：n = {counts.get('screened',0)}\n排除：n = {counts.get('ta_excluded',0)}")
    box(0.12, 0.38, 0.76, 0.14, f"合格性\nFull text 評估：n = {counts.get('ft_assessed',0)}\n排除（附理由）：n = {counts.get('ft_excluded',0)}")
    box(0.12, 0.18, 0.76, 0.14, f"納入\nMeta-analysis 納入：n = {counts.get('included',0)}")

    arrow(0.5, 0.78, 0.5, 0.72)
    arrow(0.5, 0.58, 0.5, 0.52)
    arrow(0.5, 0.38, 0.5, 0.32)

    fig.tight_layout()
    return fig


# =========================
# Manuscript template
# =========================
def manuscript_template(proto: Dict[str, Any], ma_result: Dict[str, Any], lang_code: str) -> Dict[str, str]:
    pico = proto.get("pico", {}) if isinstance(proto, dict) else {}
    P, I, C, O = pico.get("P",""), pico.get("I",""), pico.get("C",""), pico.get("O","")
    study = st.session_state.get("study_type", "")
    goal = proto.get("goal_mode","")
    outcome = ma_result.get("outcome_label","")
    measure = ma_result.get("measure","")
    model = ma_result.get("model","")
    pooled = ma_result.get("pooled_disp","")
    i2 = ma_result.get("I2","")
    n = ma_result.get("k","")

    if lang_code == "en":
        return {
            "Title": f"『』: A systematic review and meta-analysis of {I} versus {C or '『comparator』'} in {P or '『population』'}",
            "Abstract": (
                "Background: 『』\n\n"
                "Methods: We conducted a systematic search in PubMed and other sources. "
                f"Study-type focus: {study}. Eligibility criteria were defined by PICO. 『』\n\n"
                f"Results: 『Number of studies』 studies were included. For {outcome or '『primary outcome』'}, "
                f"the pooled {model} {measure} was {pooled or '『pooled effect』'} (heterogeneity I²={i2 or '『』'}). 『』\n\n"
                "Conclusions: 『』"
            ),
            "Introduction": (
                f"『Clinical context』\n\n"
                f"This review addresses: {I} versus {C or '『comparator』'} in {P or '『population』'}. "
                "Prior evidence (including SR/MA/NMA) suggests 『』; however, gaps remain. 『』"
            ),
            "Methods": (
                "Protocol and reporting: 『PRISMA statement』.\n"
                f"Eligibility criteria: P={P or '『』'}; I={I or '『』'}; C={C or '『』'}; O={O or '『』'}. "
                f"Study design: {study}. 『』\n\n"
                "Search strategy: We used a PubMed query (see below) and screened titles/abstracts, then full texts. 『』\n\n"
                "Data extraction: Two reviewers 『』; extracted outcomes and effect measures; resolved discrepancies by 『』.\n\n"
                "Risk of bias: ROB 2.0 was used; reasons recorded per domain. 『』\n\n"
                f"Statistical analysis: {model} model; effect measure={measure}. Heterogeneity assessed with I². 『』"
            ),
            "Results": (
                f"Study selection: PRISMA flow summarized. 『』\n\n"
                f"Included studies: k={n or '『』'}. Study characteristics are summarized in Table 『』.\n\n"
                f"Meta-analysis: For {outcome or '『outcome』'}, pooled {model} {measure} = {pooled or '『』'}; I²={i2 or '『』'}. "
                "Forest plot shown in Figure 『』.\n\n"
                "Additional outcomes: 『』"
            ),
            "Discussion": (
                "Principal findings: 『』\n\n"
                "Comparison with prior evidence (SR/MA/NMA): 『』\n\n"
                "Clinical implications: 『』\n\n"
                "Limitations: 『』\n\n"
                "Future research: 『』"
            ),
            "Conclusions": "『』",
        }

    # zh default
    return {
        "標題": f"『』：{I} 與 {C or '『比較組』'} 在 {P or '『族群』'} 之系統性回顧與統合分析",
        "摘要": (
            "背景：『』\n\n"
            "方法：本研究依 PRISMA 精神進行資料庫檢索與篩選。"
            f"文章類型重點：{study}。以 PICO 擬定納入/排除標準並進行全文審查。『』\n\n"
            f"結果：共納入『研究數』篇研究。以 {outcome or '『主要 outcome』'} 為例，"
            f"統合 {model} 模型之 {measure} 為 {pooled or '『統合效應值』'}（異質性 I²={i2 or '『』'}）。『』\n\n"
            "結論：『』"
        ),
        "前言": (
            "臨床背景：『』\n\n"
            f"本研究欲回答之問題為：{I} 相較於 {C or '『比較組』'}，在 {P or '『族群』'} 的效果與安全性。"
            "既有 SR/MA/NMA 證據顯示『』，但仍存在研究缺口（如族群/介入細分/結局一致性/新研究加入等）。『』"
        ),
        "方法": (
            "研究設計與報告：依 PRISMA 準則撰寫。『』\n\n"
            f"納入/排除標準（PICO）：P={P or '『』'}；I={I or '『』'}；C={C or '『』'}；O={O or '『』'}；"
            f"研究設計偏好：{study}；目標模式：{goal}。『』\n\n"
            "搜尋策略：使用自動產生之 PubMed 搜尋式（可手動微調），並可加上排除關鍵字（NOT）。『』\n\n"
            "篩選流程：先進行 Title/Abstract 粗篩，再進入 Full text review；全文排除需附理由。『』\n\n"
            "資料萃取：建立寬表（可自訂 schema），並納入既有 RCT 之 primary/secondary outcomes 以減少遺漏。『』\n\n"
            "偏倚風險：採 ROB 2.0 五大 domain + overall，並要求每一 domain 提供判斷理由。『』\n\n"
            f"統計方法：以 {model} 模型進行統合；主要效應量={measure}；以 I² 評估異質性。『』"
        ),
        "結果": (
            "研究篩選：PRISMA 流程圖/文字摘要如下。『』\n\n"
            f"納入研究特徵：k={n or '『』'}；研究特徵表見『』。『』\n\n"
            f"統合分析：以 {outcome or '『outcome』'} 為例，統合 {model} 模型之 {measure}={pooled or '『』'}；I²={i2 or '『』'}。"
            "森林圖見下方。『』\n\n"
            "其他 outcomes：『』"
        ),
        "討論": (
            "主要發現：『』\n\n"
            "與既有證據（SR/MA/NMA）比較：『』\n\n"
            "臨床意涵：『』\n\n"
            "限制：『』\n\n"
            "未來研究方向：『』"
        ),
        "結論": "『』",
    }


# =========================
# Sidebar (settings)
# =========================
with st.sidebar:
    st.header("設定")

    st.selectbox("Language / 語言", options=list(LANGS.keys()), key="lang")
    lang_code = LANGS.get(st.session_state["lang"], "zh")

    st.subheader(tr("byok_title"))
    st.checkbox(tr("byok_enable"), key="byok_enabled", help="關閉時：所有需要 LLM 的功能自動降級（不會卡住）。")
    st.caption(tr("byok_hint"))
    st.session_state["byok_consent"] = st.checkbox(tr("byok_consent"), value=bool(st.session_state.get("byok_consent", False)))

    if st.session_state.get("byok_enabled"):
        st.text_input("Base URL", key="byok_base_url", help="OpenAI: https://api.openai.com/v1")
        st.text_input("Model", key="byok_model")
        st.number_input("Temperature", min_value=0.0, max_value=1.0, value=float(st.session_state.get("byok_temp", 0.2)), step=0.05, key="byok_temp")
        st.text_input("API key", key="byok_key", type="password")
        if st.button(tr("byok_clear")):
            st.session_state["byok_key"] = ""
            st.success("已清除（僅限此 session）。")

    st.markdown("---")
    st.selectbox(tr("studytype"), options=["RCT", "Any (broader)", "SR/MA/NMA (feasibility only)"], key="study_type")
    st.selectbox(tr("goal"), options=["Rigorous / comprehensive（嚴謹）", "Fast / feasible（可行性優先）"], key="goal_mode")
    st.number_input(tr("max_records"), min_value=50, max_value=2000, value=int(st.session_state.get("max_records", 300)), step=50, key="max_records")
    st.slider(tr("delay"), min_value=0.0, max_value=2.0, value=float(st.session_state.get("polite_delay", 0.25)), step=0.05, key="polite_delay")

    st.markdown(
        f"<div class='notice'><b>{tr('disclaimer')}</b></div>",
        unsafe_allow_html=True
    )


# =========================
# Header
# =========================
st.title(tr("title"))
st.caption(tr("author"))

# =========================
# Input
# =========================
q = st.text_area(tr("question"), value=st.session_state.get("question", ""), height=140, key="question")

run = st.button(tr("run"), type="primary")

if not run:
    st.stop()

# =========================
# Pipeline starts
# =========================
st.markdown("<div class='ok'><b>Done.</b> 請往下看輸出。</div>", unsafe_allow_html=True)

# ---- Build protocol (PICO, criteria, schema) ----
q_for_parse = maybe_translate_to_english(q) if contains_cjk(q) else q
proto: Dict[str, Any] = {
    "question_original": q,
    "question_en": q_for_parse if q_for_parse != q else "",
    "pico": parse_question_rule(q_for_parse),
    "search_expansion": {"P_synonyms": [], "I_synonyms": [], "C_synonyms": [], "O_synonyms": [], "NOT": []},
    "mesh_candidates": {"P": [], "I": [], "C": [], "O": []},
    "goal_mode": st.session_state.get("goal_mode", ""),
    "study_type": st.session_state.get("study_type", ""),
}
proto["search_expansion"]["NOT"] = [t.strip() for t in re.split(r"\s+OR\s+|,|;|/|\n", proto["pico"].get("NOT","")) if t.strip()]

# LLM upgrade for protocol (optional)
if llm_available():
    try:
        msg = [
            {"role": "system", "content": "You are a rigorous evidence-synthesis assistant. Output JSON only."},
            {"role": "user", "content": f"""
Given ONE research question, infer a usable SR/MA protocol.

Return JSON with keys:
- pico: {{P,I,C,O,NOT}} (strings; NOT should be OR-joined terms)
- search_expansion: synonyms lists for P/I/C/O (free-text terms; include common brand/device names if applicable)
- mesh_candidates: suggested MeSH terms for P/I/C/O (strings; can be empty)
- criteria: inclusion/exclusion bullets (short)
- schema: recommended extraction schema (list of outcome labels + key baseline covariates)
- feasibility_plan: how to check existing SR/MA/NMA and how to adjust PICO for novelty/feasibility
- effect_plan: likely effect measures and how to compute (e.g., OR/RR/HR/MD/SMD)

Constraints:
- Write in Traditional Chinese.
- Do not fabricate results; this is only protocol planning.
- Be explicit about trade-off between (gap-fill fast) vs (rigorous scope).

Question: {q}
Study-type filter: {st.session_state.get('study_type')}
Goal mode: {st.session_state.get('goal_mode')}
"""}
        ]
        out = call_llm(msg, max_tokens=1100)
        js = json.loads(out)
        # merge safe
        if isinstance(js, dict):
            proto.update({k: js.get(k, proto.get(k)) for k in js.keys()})
            if "pico" in js and isinstance(js["pico"], dict):
                # keep NOT default if missing
                js["pico"].setdefault("NOT", proto["pico"].get("NOT",""))
            proto["pico"] = js.get("pico", proto["pico"])
            if "schema" in js and isinstance(js["schema"], dict):
                # allow schema dict
                pass
    except Exception as e:
        st.warning(f"LLM 建議 protocol 失敗（將降級）：{e}")
else:
    st.info(tr("no_llm"))

st.session_state["protocol"] = proto

# =========================
# Outputs
# =========================
st.header(tr("outputs"))

with st.expander("Protocol（current）", expanded=True):
    st.json(proto)

# ---- Build PubMed query UI blocks ----
pico_raw = proto.get("pico", {}) if isinstance(proto, dict) else {}
# ------- Manual PICO correction (for PubMed, preferably English) -------
if "pico_search" not in st.session_state or not isinstance(st.session_state.get("pico_search"), dict):
    # Initialize from protocol; if question contains Chinese and LLM is enabled, attempt translation to English.
    pico_init = {}
    for k in ["P","I","C","O","NOT"]:
        v = str(pico_raw.get(k, "") or "")
        pico_init[k] = maybe_translate_to_english(v) if k in ["P","I","C","O"] else v
    pico_init["NOT"] = pico_init.get("NOT") or "animal OR mice OR rat OR in vitro OR case report"
    st.session_state["pico_search"] = pico_init

with st.expander("人工修正 PICO（英文檢索用；會影響 Step 1/2 的 PubMed 查詢）", expanded=False):
    st.caption("若你的問句是中文，本區建議改成英文關鍵詞（可含品牌/裝置名），以提升 PubMed 召回。")
    # If I and C are identical, prompt user to refine
    if (st.session_state["pico_search"].get("I","").strip().lower() and
        st.session_state["pico_search"].get("I","").strip().lower() == st.session_state["pico_search"].get("C","").strip().lower()):
        st.warning("偵測到 I 與 C 相同（例如『EDOF vs EDOF』）。請在此處把 I/C 改成不同類型、品牌或設計（例如 diffractive vs nondiffractive 或特定 IOL 名稱），否則 PubMed 會抓不到。")
    c1, c2 = st.columns(2)
    with c1:
        st.text_input("P（Population / Problem；英文）", key="picoP_en", value=st.session_state["pico_search"].get("P",""))
        st.text_input("I（Intervention；英文）", key="picoI_en", value=st.session_state["pico_search"].get("I",""))
    with c2:
        st.text_input("C（Comparator；英文；可留空）", key="picoC_en", value=st.session_state["pico_search"].get("C",""))
        st.text_input("O（Outcome；英文；可留空）", key="picoO_en", value=st.session_state["pico_search"].get("O",""))
    st.text_input("NOT（排除；用 OR 串）", key="picoNOT", value=st.session_state["pico_search"].get("NOT",""))
    colx, coly = st.columns([1,1])
    with colx:
        if st.button("套用此 PICO（影響 PubMed query）"):
            st.session_state["pico_search"] = {
                "P": st.session_state.get("picoP_en","").strip(),
                "I": st.session_state.get("picoI_en","").strip(),
                "C": st.session_state.get("picoC_en","").strip(),
                "O": st.session_state.get("picoO_en","").strip(),
                "NOT": st.session_state.get("picoNOT","").strip() or "animal OR mice OR rat OR in vitro OR case report",
            }
            # Rebuild query immediately
            st.session_state["pubmed_query"] = ""
            st.success("已套用。請在 Step 1 檢查 PubMed query，必要時再手改。")
    with coly:
        if st.button("重設為 AI 推論（並嘗試自動英文化）"):
            pico_reset = {}
            for k in ["P","I","C","O","NOT"]:
                v = str(pico_raw.get(k, "") or "")
                pico_reset[k] = maybe_translate_to_english(v) if k in ["P","I","C","O"] else v
            pico_reset["NOT"] = pico_reset.get("NOT") or "animal OR mice OR rat OR in vitro OR case report"
            st.session_state["pico_search"] = pico_reset
            st.session_state["pubmed_query"] = ""
            st.success("已重設。")

pico = st.session_state.get("pico_search", pico_raw)  # use this for PubMed
P, I, C, O = pico.get("P",""), pico.get("I",""), pico.get("C",""), pico.get("O","")
NOT = pico.get("NOT","animal OR mice OR rat OR in vitro OR case report")

warn_if_non_english("P", P)
warn_if_non_english("I", I)
extra_kw = st.text_input("擴充關鍵字（可留空；例如 device/brand 名稱，多個用逗號）", value="")

# Rebuild query button
if st.button(tr("rebuild_query")):
    st.session_state["pubmed_query"] = build_pubmed_query(pico, include_CO, add_rct, extra_kw, NOT)

# always show current query editor
if not st.session_state.get("pubmed_query"):
    st.session_state["pubmed_query"] = build_pubmed_query(pico, include_CO, add_rct, extra_kw, NOT)

pub_q = st.text_area(tr("pubmed_query"), value=st.session_state["pubmed_query"], height=120)
st.session_state["pubmed_query"] = pub_q

# ---- Feasibility scan (SR/MA/NMA) ----
st.subheader(tr("feasibility"))
base_for_feas = build_pubmed_query(pico, include_CO=False, add_rct=False, extra=extra_kw, not_terms=NOT)
feas_q = feasibility_query(base_for_feas, NOT)

with st.expander("可行性掃描：PubMed（SR/MA/NMA）", expanded=False):
    st.code(feas_q, language="text")
    do_feas = st.button("執行可行性掃描（可能較慢）")
    if do_feas:
        try:
            cnt, ids, url, warns = pubmed_esearch(feas_q, retmax=50, retstart=0)
            st.markdown(f"<div class='kpiwrap'><div class='kpi'><div class='label'>SR/MA/NMA 可能數量</div><div class='value'>{cnt}</div></div></div>", unsafe_allow_html=True)
            if warns:
                st.warning("；".join(warns))
            # fetch top few
            df_feas, warns2 = pubmed_efetch(ids[:20])
            if warns2:
                st.warning("；".join(warns2))
            if df_feas.empty:
                st.info("未抓到摘要。")
            else:
                for _, r in df_feas.iterrows():
                    st.markdown(f"- {r.get('year','')} {r.get('first_author','')} — {r.get('title','')}")
        except Exception as e:
            st.error(f"可行性掃描失敗：{e}")

# ---- PubMed retrieve ----
st.markdown("---")
st.subheader("Step 2：抓 PubMed 文獻（含摘要）")

diag = {}
df_records = pd.DataFrame()
warnings_all: List[str] = []

try:
    total, ids, es_url, warns = pubmed_esearch(pub_q, retmax=min(int(st.session_state["max_records"]), 500), retstart=0)
    warnings_all += warns
    diag["pubmed_total_count"] = total
    diag["esearch_url"] = es_url

    # fetch in batches
    pmids = ids[: int(st.session_state["max_records"])]
    batch_size = 100
    dfs = []
    efetch_urls = []
    for i in range(0, len(pmids), batch_size):
        if st.session_state["polite_delay"] > 0:
            time.sleep(float(st.session_state["polite_delay"]))
        sub = pmids[i:i+batch_size]
        params = {"db": "pubmed", "retmode": "xml", "id": ",".join(sub)}
        efetch_urls.append(requests.Request("GET", NCBI_EFETCH, params=params).prepare().url)
        dfi, warns2 = pubmed_efetch(sub)
        warnings_all += warns2
        if not dfi.empty:
            dfs.append(dfi)
    diag["efetch_urls"] = efetch_urls

    df_records = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
    if not df_records.empty:
        df_records = df_records.drop_duplicates(subset=["pmid"], keep="first")
except Exception as e:
    warnings_all.append(str(e))

diag["warnings"] = warnings_all
st.session_state["records_df"] = df_records

if df_records.empty:
    st.markdown("<div class='bad'><b>沒有抓到 records。</b> 若你覺得不合理，請打開 Diagnostics：PubMed 可能被擋住或回傳 HTML。</div>", unsafe_allow_html=True)
else:
    st.markdown(f"<div class='ok'><b>PubMed：</b>抓到 {len(df_records)} 篇（已去除重複）。</div>", unsafe_allow_html=True)

with st.expander(tr("diag"), expanded=False):
    st.json(diag)

# ---- Title/abstract screening (Step 3+4 merged) ----
st.markdown("---")
st.subheader(tr("screening"))

if df_records.empty:
    st.stop()

study_filter = st.session_state.get("study_type", "RCT")

# Precompute AI suggestions (cached per record_id)
for _, row in df_records.iterrows():
    rid = row["record_id"]
    if rid in st.session_state["ta_ai"]:
        continue
    if llm_available():
        try:
            st.session_state["ta_ai"][rid] = ai_screen_llm(row, proto, study_filter)
        except Exception as e:
            st.session_state["ta_ai"][rid] = ai_screen_fallback(row, proto, study_filter, err=str(e))
    else:
        st.session_state["ta_ai"][rid] = ai_screen_fallback(row, proto, study_filter)

st.caption("以下為 Title/Abstract 粗篩（Step 3+4 合併）。你可保留 AI 判讀與理由，同時用下拉選單人工覆寫。")

# Optional filters
fcol1, fcol2, fcol3 = st.columns([1,1,1])
with fcol1:
    show_only = st.selectbox("顯示篩選", options=["全部", "Effective=Include", "Effective=Exclude", "Effective=Unsure"], index=0)
with fcol2:
    show_abs = st.checkbox("顯示摘要", value=True)
with fcol3:
    compact = st.checkbox("精簡顯示（較少文字）", value=False)

def _effective_bucket(rid: str) -> str:
    eff = compute_effective_ta(rid)
    return "Include" if eff == "Include" else ("Exclude" if eff == "Exclude" else "Unsure")

for _, r in df_records.iterrows():
    rid = r["record_id"]
    eff = compute_effective_ta(rid)
    if show_only == "Effective=Include" and eff != "Include":
        continue
    if show_only == "Effective=Exclude" and eff != "Exclude":
        continue
    if show_only == "Effective=Unsure" and eff != "Unsure":
        continue

    title = str(r.get("title","") or "")
    ai = st.session_state["ta_ai"].get(rid, {}) or {}
    ai_dec = ai.get("decision","Unsure")
    ai_conf = float(ai.get("confidence", 0.5) or 0.5)
    ai_reason = ai.get("reason","") or ""

    badge = "badge-unsure"
    if eff == "Include":
        badge = "badge-include"
    elif eff == "Exclude":
        badge = "badge-exclude"

    meta_line = f"PMID: {r.get('pmid','') or '—'}    DOI: {r.get('doi','') or '—'}    Year: {r.get('year','') or '—'}    First author: {r.get('first_author','') or '—'}    Journal: {r.get('journal','') or '—'}"

    with st.expander(f"{rid}｜{short(title, 110)}", expanded=False):
        st.markdown(
            f"<span class='badge {badge}'>Effective：{eff}</span>"
            f"<span class='badge'>AI：{ai_dec}</span>"
            f"<span class='badge'>信心度：{ai_conf:.2f}</span>",
            unsafe_allow_html=True
        )
        st.markdown(f"<div class='small'>{meta_line}</div>", unsafe_allow_html=True)

        links = []
        if r.get("pubmed_url"):
            links.append(f"[PubMed]({r.get('pubmed_url')})")
        if r.get("doi_url"):
            links.append(f"[DOI]({r.get('doi_url')})")
        if r.get("pmc_url"):
            links.append(f"[PMC]({r.get('pmc_url')})")
        if links:
            st.markdown(" | ".join(links))

        if not compact:
            st.markdown("**AI 建議（Title/Abstract）**")
            st.write(f"- 建議：**{ai_dec}**")
            if ai_reason.strip():
                st.write(f"- 理由：{ai_reason}")
        # Human override
        ov = st.selectbox("人工覆寫（Title/Abstract）", options=["(不覆寫)", "Include", "Exclude", "Unsure"], index=0, key=f"ov_ta_{rid}")
        reason = st.text_input("人工理由（可留空；建議寫排除/保留原因）", value=st.session_state['ta_reason'].get(rid, ''), key=f"ov_ta_reason_{rid}")
        if st.button("儲存此筆粗篩決策", key=f"save_ta_{rid}"):
            if ov and ov != "(不覆寫)":
                st.session_state["ta_override"][rid] = ov
            elif rid in st.session_state["ta_override"]:
                st.session_state["ta_override"].pop(rid, None)
            st.session_state["ta_reason"][rid] = reason
            st.success("已儲存。")

        if show_abs:
            st.markdown("**摘要（分段）**")
            abs_txt = str(r.get("abstract","") or "")
            if abs_txt.strip():
                paras = split_abstract_paras(abs_txt)
                for p in paras:
                    st.markdown(p)
            else:
                st.write("（無摘要或未擷取到）")

# ---- Step 4b: Full text review (separate section) ----
st.markdown("---")
st.subheader("Step 4b. Full text review（全文審查：含排除理由 + PDF 上傳 + 可選抽字）")

ft_pool = df_records[df_records["record_id"].apply(lambda x: compute_effective_ta(str(x)) != "Exclude")].copy()
if ft_pool.empty:
    st.info("目前沒有可做全文審查的文章（粗篩全為 Exclude）。")
else:
    st.caption("建議做法：先在上方粗篩至少保留 Include/Unsure，再在此步驟進入全文審查。全文排除務必寫理由。")
    st.caption("重要提醒：若全文來自校內訂閱/授權，避免把受版權保護 PDF 上傳至雲端部署。建議改用本機版或僅上傳 OA/PMC。")

    for _, r in ft_pool.iterrows():
        rid = r["record_id"]
        title = str(r.get("title","") or "")
        eff = compute_effective_ta(rid)
        ft_dec = st.session_state["ft_decisions"].get(rid, "Not reviewed")
        ft_reason = st.session_state["ft_reasons"].get(rid, "")

        with st.expander(f"{rid}｜{short(title, 110)}", expanded=False):
            st.markdown(f"<span class='badge'>粗篩 Effective：{eff}</span>", unsafe_allow_html=True)
            st.markdown(f"<div class='small'>PMID: {r.get('pmid','') or '—'} | {r.get('first_author','') or '—'} {r.get('year','') or '—'} | {r.get('journal','') or '—'}</div>", unsafe_allow_html=True)

            # Full text decision
            ft_dec2 = st.selectbox(
                "全文決策（Full text）",
                options=["Not reviewed", "Include for full-text", "Exclude", "Include for meta-analysis"],
                index=["Not reviewed", "Include for full-text", "Exclude", "Include for meta-analysis"].index(ft_dec) if ft_dec in ["Not reviewed","Include for full-text","Exclude","Include for meta-analysis"] else 0,
                key=f"ft_dec_{rid}",
            )
            ft_reason2 = st.text_input("全文排除理由（若 Exclude 必填）", value=ft_reason, key=f"ft_reason_{rid}")

            # PDF upload (optional)
            up = st.file_uploader("上傳全文 PDF（可選；建議僅上傳你有權分享的 OA/PMC）", type=["pdf"], key=f"ft_pdf_{rid}")
            extracted = ""
            if up is not None:
                try:
                    extracted = extract_text_from_pdf_bytes(up.getvalue(), max_pages=6)
                    st.success("已嘗試抽字（前 6 頁）。你也可在下方自行補貼關鍵段落。")
                    if extracted.strip():
                        with st.expander("抽字結果（可複製/修正）", expanded=False):
                            st.text_area("Extracted text", value=extracted, height=220, key=f"ft_extracted_{rid}")
                except Exception as e:
                    st.warning(f"PDF 抽字失敗：{e}")

            ft_text = st.text_area(
                "Full text（可貼全文或關鍵段落；可留空）",
                value=st.session_state["ft_text"].get(rid, "") or extracted,
                height=180,
                key=f"ft_text_{rid}"
            )

            if st.button("儲存 Full text review", key=f"save_ft_{rid}"):
                st.session_state["ft_decisions"][rid] = ft_dec2
                st.session_state["ft_reasons"][rid] = ft_reason2
                st.session_state["ft_text"][rid] = ft_text
                st.success("已儲存。")
# ---- PRISMA ----
st.markdown("---")
st.subheader(tr("prisma"))

identified = len(df_records)
screened = identified
ta_excluded = sum(1 for rid in df_records["record_id"].tolist() if compute_effective_ta(rid) == "Exclude")
ft_assessed = sum(1 for rid in df_records["record_id"].tolist() if compute_effective_ta(rid) in ["Include","Unsure"] or st.session_state["ft_decisions"].get(rid) != "Not reviewed yet")
ft_excluded = sum(1 for rid, d in st.session_state["ft_decisions"].items() if d == "Exclude after full-text")
included = sum(1 for rid, d in st.session_state["ft_decisions"].items() if d == "Include for meta-analysis")

counts = {
    "identified": identified,
    "screened": screened,
    "ta_excluded": ta_excluded,
    "ft_assessed": ft_assessed,
    "ft_excluded": ft_excluded,
    "included": included,
}

col1, col2 = st.columns([1.1, 1])
with col1:
    if HAS_MPL:
        fig = draw_prisma(counts)
        if fig is not None:
            st.pyplot(fig, clear_figure=True)
    else:
        st.info("環境未安裝 matplotlib：改用文字版 PRISMA。")
with col2:
    st.markdown(
        f"""
- 識別（identified）：{identified}
- Title/Abstract 排除：{ta_excluded}
- Full text 評估：{ft_assessed}
- Full text 排除（附理由）：{ft_excluded}
- 最終納入 MA：{included}
"""
    )

# =========================
# Step 1-1 schema + criteria
# =========================
st.markdown("---")
st.subheader("Step 1-1：extraction schema（欄位不寫死，可自訂；也可由 AI 建議）")

schema_lines = st.text_area("Outcome / 欄位名稱（每行一個，可自訂）", value=st.session_state["schema_lines"], height=120)
st.session_state["schema_lines"] = schema_lines
schema = [x.strip() for x in schema_lines.splitlines() if x.strip()]

if llm_available() and st.button("用 AI 建議 outcomes/schema（會考量既有 RCT primary/secondary outcomes）"):
    try:
        msg = [
            {"role": "system", "content": "You are a systematic review methodologist. Output JSON only."},
            {"role": "user", "content": f"""
基於此 PICO 與目標，請提出建議的 extraction schema。
要求：
- outcomes：列出主要/次要 outcome（可多個）；要考量既有 RCT 常用 primary/secondary outcomes，避免遺漏。
- base_cols：研究特徵/族群/介入細節/追蹤時間等欄位建議。
輸出 JSON：{{"outcomes":[...], "base_cols":[...]}}
PICO: {json.dumps(pico, ensure_ascii=False)}
"""}
        ]
        out = call_llm(msg, max_tokens=550)
        js = json.loads(out)
        outs = js.get("outcomes", [])
        if isinstance(outs, list) and outs:
            st.session_state["schema_lines"] = "\n".join([str(x) for x in outs if str(x).strip()])
            schema = [x.strip() for x in st.session_state["schema_lines"].splitlines() if x.strip()]
        st.success("已更新 outcomes/schema。")
    except Exception as e:
        st.error(f"AI 建議 schema 失敗：{e}")

st.subheader("Step 1-2：inclusion/exclusion criteria（PICO 層級）")
default_criteria = proto.get("criteria", "")
if isinstance(default_criteria, list):
    default_criteria = "\n".join([f"- {x}" for x in default_criteria])
criteria_text = st.text_area("Criteria（可人工修正；全文排除理由可引用此標準）", value=(st.session_state.get("criteria_text") or default_criteria or ""), height=160)
st.session_state["criteria_text"] = criteria_text

# =========================
# =========================
# Step 5 Extraction
# =========================
st.markdown("---")
st.subheader(tr("extraction"))

# Candidate set for MA: prefer FT=Include for meta-analysis; else allow FT include; else TA include
ft_meta = [rid for rid, d in st.session_state["ft_decisions"].items() if d == "Include for meta-analysis"]
ft_full = [rid for rid, d in st.session_state["ft_decisions"].items() if d == "Include for full-text"]
ta_inc = [rid for rid in df_records["record_id"].tolist() if compute_effective_ta(rid) == "Include"]

eligible = ft_meta or ft_full or ta_inc

if not eligible:
    st.warning("目前沒有可供 extraction 的文章（建議先在粗篩或全文決策至少選 Include）。")
else:
    st.caption("建議先完成 Full text review 並標記「Include for meta-analysis」，再進入 extraction。若尚未做全文，本步驟仍允許先用粗篩 Include 建立寬表。")

# Metadata map
meta_cols = ["record_id","pmid","year","doi","journal","first_author","title","pubmed_url","doi_url","pmc_url"]
meta_map = {}
for _, r in df_records.iterrows():
    rid = str(r.get("record_id",""))
    meta_map[rid] = {c: r.get(c, "") for c in meta_cols}

# Load existing extraction table (long format; one row per study-outcome-timepoint)
ex = st.session_state.get("extract_df", pd.DataFrame())
if not isinstance(ex, pd.DataFrame):
    try:
        ex = pd.DataFrame(ex)
    except Exception:
        ex = pd.DataFrame()

ensure_columns(
    ex,
    meta_cols + ["Outcome_label","Timepoint","Effect_measure","Effect","Lower_CI","Upper_CI","Effect_unit","Notes"],
    default=""
)

# If empty, offer to seed one empty row per eligible record
if ex.empty:
    if st.button("建立初始寬表（每篇 1 列；可再用『快速新增』加入更多 outcome）"):
        seed_rows = []
        for rid in eligible:
            m = meta_map.get(rid, {})
            seed_rows.append({
                **{c: m.get(c,"") for c in meta_cols},
                "Outcome_label": "",
                "Timepoint": "",
                "Effect_measure": "",
                "Effect": "",
                "Lower_CI": "",
                "Upper_CI": "",
                "Effect_unit": "",
                "Notes": "",
            })
        ex = pd.DataFrame(seed_rows)
        st.session_state["extract_df"] = ex

# -------- Quick add row (reduces rerun pain) --------
with st.expander("快速新增一筆（建議：每個 outcome / timepoint 各一列）", expanded=not ex.empty):
    # Build label list
    opts = []
    for rid in eligible:
        m = meta_map.get(rid, {})
        lab = f"PMID:{m.get('pmid','') or '—'} | {m.get('first_author','') or '—'} {m.get('year','') or ''} | {short(str(m.get('title','')), 70)}"
        opts.append((lab, rid))
    if not opts:
        st.info("（沒有 eligible 文章）")
    else:
        labels = [o[0] for o in opts]
        rid_by_label = {o[0]: o[1] for o in opts}
        with st.form("quick_add_form", clear_on_submit=False):
            sel_lab = st.selectbox("選擇文章", options=labels, index=0)
            rid = rid_by_label.get(sel_lab, eligible[0])
            c1, c2, c3 = st.columns(3)
            with c1:
                outc = st.text_input("Outcome_label", value="")
                tpt = st.text_input("Timepoint", value="")
            with c2:
                meas = st.selectbox("Effect_measure", options=["", "OR","RR","HR","MD","SMD"], index=0)
                unit = st.text_input("Unit/Scale（可留空）", value="")
            with c3:
                eff = st.text_input("Effect（數值）", value="")
                lcl = st.text_input("Lower CI", value="")
                ucl = st.text_input("Upper CI", value="")
            notes = st.text_area("Notes（可留空）", value="", height=90)
            add_btn = st.form_submit_button("新增一列到寬表")
        if add_btn:
            m = meta_map.get(rid, {})
            new_row = {
                **{c: m.get(c,"") for c in meta_cols},
                "Outcome_label": outc.strip(),
                "Timepoint": tpt.strip(),
                "Effect_measure": meas.strip().upper(),
                "Effect": eff.strip(),
                "Lower_CI": lcl.strip(),
                "Upper_CI": ucl.strip(),
                "Effect_unit": unit.strip(),
                "Notes": notes.strip(),
            }
            ex = pd.concat([ex, pd.DataFrame([new_row])], ignore_index=True)
            st.session_state["extract_df"] = ex
            st.success("已新增。請在下方『進階：一次編輯整張表』檢查/修正後再儲存。")

# -------- Advanced edit (commit on Save) --------
st.markdown("**進階：一次編輯整張表（按儲存才 commit）**")
st.caption("為避免輸入時一直 rerun：此表格在你按下『儲存/更新寬表』前不會覆寫正式資料。")

# Buffer for editor
if "extract_editor_buf" not in st.session_state or not isinstance(st.session_state.get("extract_editor_buf"), pd.DataFrame):
    st.session_state["extract_editor_buf"] = ex.copy()

# Keep buffer synced when ex changes significantly (e.g., quick-add)
if len(st.session_state["extract_editor_buf"]) != len(ex):
    st.session_state["extract_editor_buf"] = ex.copy()

with st.form("extract_editor_form"):
    buf = st.data_editor(
        st.session_state["extract_editor_buf"],
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "record_id": st.column_config.TextColumn("record_id", disabled=True),
            "pmid": st.column_config.TextColumn("PMID", disabled=True, width="small"),
            "first_author": st.column_config.TextColumn("First author", disabled=True, width="small"),
            "year": st.column_config.TextColumn("Year", disabled=True, width="small"),
            "journal": st.column_config.TextColumn("Journal", disabled=True, width="small"),
            "title": st.column_config.TextColumn("Title", disabled=True, width="large"),
            "Effect_measure": st.column_config.SelectboxColumn("Effect measure", options=["", "OR","RR","HR","MD","SMD"]),
            "Effect": st.column_config.NumberColumn("Effect", format="%.6f"),
            "Lower_CI": st.column_config.NumberColumn("Lower CI", format="%.6f"),
            "Upper_CI": st.column_config.NumberColumn("Upper CI", format="%.6f"),
        },
    )
    saved = st.form_submit_button("儲存/更新寬表（不會立即跑森林圖）")

# Normalize buf to DataFrame, ensure required columns (Fix KeyError)
if not isinstance(buf, pd.DataFrame):
    try:
        buf = pd.DataFrame(buf)
    except Exception:
        buf = ex.copy()

ensure_columns(buf, meta_cols + ["Outcome_label","Timepoint","Effect_measure","Effect","Lower_CI","Upper_CI","Effect_unit","Notes"], default="")

# Commit on save
if saved:
    # back-fill metadata for any rows that came from manual copy/paste
    for i, rr in buf.iterrows():
        rid = str(rr.get("record_id","") or "")
        if rid and rid in meta_map:
            for c in meta_cols:
                if (not str(rr.get(c,"")).strip()) and str(meta_map[rid].get(c,"")).strip():
                    buf.at[i, c] = meta_map[rid].get(c,"")
    st.session_state["extract_df"] = buf
    st.session_state["extract_editor_buf"] = buf.copy()
    st.session_state["extract_last_saved"] = time.time()
    st.success("已儲存寬表。下一步請到 Step 6 按『更新森林圖/MA』。")

# Validation (non-blocking; red hints without crashing)
v = st.session_state.get("extract_df", pd.DataFrame()).copy()
if not isinstance(v, pd.DataFrame):
    try:
        v = pd.DataFrame(v)
    except Exception:
        v = pd.DataFrame()

for c in ["Effect","Lower_CI","Upper_CI"]:
    if c not in v.columns:
        v[c] = pd.NA
    v[c] = pd.to_numeric(v[c], errors="coerce")

issues = []
if not v.empty:
    for idx, row in v.iterrows():
        rid = str(row.get("record_id",""))
        title = short(str(row.get("title","")), 60)
        outcome = str(row.get("Outcome_label","") or "").strip()
        meas = str(row.get("Effect_measure","") or "").strip().upper()
        eff = row.get("Effect", pd.NA)
        lcl = row.get("Lower_CI", pd.NA)
        ucl = row.get("Upper_CI", pd.NA)

        # allow blanks; only flag when partially filled
        if any(pd.notna(x) for x in [eff,lcl,ucl]) and not all(pd.notna(x) for x in [eff,lcl,ucl]):
            issues.append(f"{rid} | {title}: Effect/CI 未填完整（允許留空，但同一列若要用於 MA 必須填滿）")
        if all(pd.notna(x) for x in [eff,lcl,ucl]):
            if ucl <= lcl:
                issues.append(f"{rid} | {title}: CI 上界 ≤ 下界（請修正）")
            if meas in {"OR","RR","HR"} and (eff <= 0 or lcl <= 0 or ucl <= 0):
                issues.append(f"{rid} | {title}: {meas} 需要正值（Effect/CI > 0），否則無法取 log")
            if not outcome:
                issues.append(f"{rid} | {title}: 缺 Outcome_label（MA 需要）")

if issues:
    with st.expander("資料完整性提醒（不會阻擋下一步）", expanded=False):
        for it in issues[:80]:
            st.write(f"- {it}")
        if len(issues) > 80:
            st.write(f"... 另有 {len(issues)-80} 筆提醒")
# Step 6 MA + Forest
# =========================
st.markdown("---")
st.subheader(tr("ma"))

ex = st.session_state.get("extract_df", pd.DataFrame())
if not isinstance(ex, pd.DataFrame):
    try:
        ex = pd.DataFrame(ex)
    except Exception:
        ex = pd.DataFrame()
if ex is None or ex.empty:
    st.info("尚未建立 extraction 寬表。")
else:
    dfm = ex.copy()
    ensure_columns(dfm, ["Outcome_label","Effect_measure","Effect","Lower_CI","Upper_CI","Timepoint","first_author","year","record_id"], default="")
    for c in ["Effect","Lower_CI","Upper_CI"]:
        if c not in dfm.columns:
            dfm[c] = pd.NA
        dfm[c] = pd.to_numeric(dfm[c], errors="coerce")

    dfm["Outcome_label"] = dfm["Outcome_label"].astype(str).str.strip()
    dfm["Effect_measure"] = dfm["Effect_measure"].astype(str).str.strip().str.upper()

    available_outcomes = sorted([x for x in dfm["Outcome_label"].unique().tolist() if x])
    if not available_outcomes:
        st.warning("你尚未在寬表填入 Outcome_label。仍可先繼續，但 MA 需要至少一個 outcome。")
        available_outcomes = ["(未命名 outcome)"]
        dfm.loc[dfm["Outcome_label"]=="", "Outcome_label"] = "(未命名 outcome)"

    st.caption("為避免輸入時一直 rerun：請先編輯寬表，再按下方按鈕『更新森林圖/MA』。")
    default_outcome = st.session_state.get("ma_outcome_input") or available_outcomes[0]
    chosen_outcome = st.text_input("Outcome_label（手動輸入/可貼上）", value=default_outcome, key="ma_outcome_input").strip()
    if not chosen_outcome:
        chosen_outcome = available_outcomes[0]

    # measure choice
    measures_avail = sorted([m for m in dfm["Effect_measure"].unique().tolist() if m])
    if not measures_avail:
        measures_avail = ["OR","RR","HR","MD","SMD"]
    default_meas = st.session_state.get("ma_measure_choice") or (measures_avail[0] if measures_avail else "MD")
    chosen_measure = st.selectbox("Effect measure", options=measures_avail, index=measures_avail.index(default_meas) if default_meas in measures_avail else 0, key="ma_measure_choice")
    model = st.selectbox("Model", options=["Fixed effect", "Random effects (DL)"], index=0 if st.session_state.get("ma_model_choice","Fixed effect")=="Fixed effect" else 1, key="ma_model_choice")

    if st.button("更新森林圖/MA（Run）", type="primary"):
        sub = dfm[(dfm["Outcome_label"]==chosen_outcome) & (dfm["Effect_measure"]==chosen_measure)].copy()

        # keep only valid numeric
        sub = sub.dropna(subset=["Effect","Lower_CI","Upper_CI"])
        ratio_measures = {"OR","RR","HR"}
        if chosen_measure in ratio_measures:
            sub = sub[(sub["Effect"]>0) & (sub["Lower_CI"]>0) & (sub["Upper_CI"]>0)]
        # CI order
        sub = sub[sub["Upper_CI"]>sub["Lower_CI"]]

        if sub.empty:
            st.error("沒有可用資料：請確認寬表已填 Effect/CI 且 Outcome_label 與 Effect_measure 一致。")
        else:
            # compute analysis scale effects + SE
            eff_a = []
            se_a = []
            labels = []
            disp_eff = []
            disp_l = []
            disp_u = []

            for _, rr in sub.iterrows():
                eff = float(rr["Effect"])
                lcl = float(rr["Lower_CI"])
                ucl = float(rr["Upper_CI"])
                a_eff, a_l, a_u = transform_effect(eff, lcl, ucl, chosen_measure)
                se = se_from_ci(eff, lcl, ucl, chosen_measure)
                if se is None or se <= 0:
                    continue
                eff_a.append(a_eff)
                se_a.append(se)
                disp_eff.append(eff)
                disp_l.append(lcl)
                disp_u.append(ucl)
                lab = f"{rr.get('first_author','') or ''} {rr.get('year','')}".strip()
                labels.append(lab if lab else str(rr.get("record_id","")))

            if len(eff_a) == 0:
                st.error("所有資料的 SE 計算失敗（常見原因：OR/RR/HR 有 0/負值或 CI 不合理）。")
            else:
                # pool
                if model.startswith("Random"):
                    pooled = random_effect_pool(eff_a, se_a)
                    w = pooled["weights"]
                    tau2 = pooled.get("tau2", 0.0)
                else:
                    pooled = fixed_effect_pool(eff_a, se_a)
                    w = pooled["weights"]
                    tau2 = 0.0

                mu = pooled["pooled"]
                se_mu = pooled["se"]
                ci_l = mu - 1.96*se_mu
                ci_u = mu + 1.96*se_mu

                pooled_disp = back_transform(mu, chosen_measure)
                pooled_l = back_transform(ci_l, chosen_measure)
                pooled_u = back_transform(ci_u, chosen_measure)

                i2 = i2_stat(eff_a, se_a)

                st.session_state["ma_last_run"] = time.time()
                st.session_state["ma_result"] = {
                    "outcome_label": chosen_outcome,
                    "measure": chosen_measure,
                    "model": model,
                    "k": len(eff_a),
                    "pooled": float(pooled_disp),
                    "pooled_l": float(pooled_l),
                    "pooled_u": float(pooled_u),
                    "I2": float(i2),
                    "tau2": float(tau2),
                    "pooled_disp": f"{pooled_disp:.4f} (95% CI {pooled_l:.4f}–{pooled_u:.4f})",
                    "table": {
                        "label": labels,
                        "effect": disp_eff[:len(labels)],
                        "lcl": disp_l[:len(labels)],
                        "ucl": disp_u[:len(labels)],
                        "weight": [float(x) for x in w[:len(labels)]],
                    },
                }

                st.markdown(
                    f"<div class='ok'><b>統合結果</b><br>"
                    f"{model} pooled {chosen_measure} = <b>{pooled_disp:.4f}</b> "
                    f"(95% CI {pooled_l:.4f}–{pooled_u:.4f}); "
                    f"I² = {i2:.1f}%"
                    + (f"; τ² = {tau2:.4f}" if model.startswith("Random") else "")
                    + "</div>",
                    unsafe_allow_html=True
                )

                # Forest plot
                st.markdown("#### 森林圖 / Forest plot")

                tab = st.session_state["ma_result"]["table"]
                plot_df = pd.DataFrame(tab)
                # compute CI for each study for plotting
                if HAS_PLOTLY:
                    # RevMan-like: horizontal CI with marker; y reversed
                    y = list(range(len(plot_df)))[::-1]
                    fig = go.Figure()
                    fig.add_trace(go.Scatter(
                        x=plot_df["effect"], y=y,
                        mode="markers",
                        name="Effect",
                        error_x=dict(
                            type="data",
                            symmetric=False,
                            array=plot_df["ucl"]-plot_df["effect"],
                            arrayminus=plot_df["effect"]-plot_df["lcl"]
                        ),
                        text=plot_df["label"],
                        hovertemplate="%{text}<br>Effect=%{x}<extra></extra>"
                    ))
                    # null line
                    null = 1.0 if chosen_measure in {"OR","RR","HR"} else 0.0
                    fig.add_vline(x=null, line_width=1, line_dash="dash")
                    fig.update_yaxes(
                        tickmode="array",
                        tickvals=y,
                        ticktext=plot_df["label"].tolist()[::-1],
                        automargin=True
                    )
                    fig.update_layout(height=450, margin=dict(l=10, r=10, t=30, b=10))
                    st.plotly_chart(fig, use_container_width=True)
                elif HAS_MPL:
                    fig, ax = plt.subplots(figsize=(9, max(3.8, 0.42*len(plot_df)+1.8)))
                    ax.axvline(1.0 if chosen_measure in {"OR","RR","HR"} else 0.0, linestyle="--", linewidth=1)
                    y = list(range(len(plot_df)))[::-1]
                    ax.errorbar(plot_df["effect"], y,
                                xerr=[plot_df["effect"]-plot_df["lcl"], plot_df["ucl"]-plot_df["effect"]],
                                fmt="o")
                    ax.set_yticks(y)
                    ax.set_yticklabels(plot_df["label"].tolist()[::-1])
                    ax.set_xlabel(chosen_measure)
                    ax.set_title(f"{chosen_outcome} — {model}")
                    ax.grid(True, axis="x", linestyle=":", linewidth=0.6)
                    fig.tight_layout()
                    st.pyplot(fig, clear_figure=True)
                else:
                    st.warning("環境缺少 Plotly/Matplotlib：改以表格顯示森林圖資料。")
                    st.dataframe(plot_df, use_container_width=True)

# =========================
# Step 7 Manuscript draft
# =========================
st.markdown("---")
st.subheader(tr("manuscript"))

ma_result = st.session_state.get("ma_result", {}) or {}
if not ma_result:
    st.info("尚未執行 MA。你仍可先用模板產生稿件骨架。")

# Optional: style templates upload for LLM
with st.expander("（可選）上傳書寫範本（DOCX）以引導語氣", expanded=False):
    st.write("若啟用 LLM，可把範本文字作為 style guide；若未啟用，仍會用模板（以『』留空）。")
    tmpl_files = st.file_uploader("上傳 1–3 份 DOCX（範本）", type=["docx"], accept_multiple_files=True)
    style_text = ""
    if tmpl_files and HAS_DOCX:
        for f in tmpl_files[:3]:
            try:
                d = Document(io.BytesIO(f.getvalue()))
                paras = [p.text.strip() for p in d.paragraphs if p.text.strip()]
                style_text += "\n".join(paras[:120]) + "\n\n"
            except Exception:
                continue
        style_text = style_text.strip()

# Generate manuscript
if st.button("產生/更新稿件草稿（分段呈現）"):
    if llm_available():
        try:
            msg = [
                {"role":"system", "content":"You are a senior academic writer for SR/MA. Output JSON only."},
                {"role":"user", "content": f"""
Write a systematic review / meta-analysis manuscript draft in academic English.
Emulate the style guide (if provided) in tone/structure, but do not copy verbatim.

Requirements:
- Return a JSON object where each key is a section title and each value is the section text.
- Required sections: Title, Abstract (structured: Background/Methods/Results/Conclusions), Introduction, Methods, Results, Discussion, Conclusions.
- Optional sections: Limitations, Funding/Conflicts.
- Use full-width brackets 『』 to mark any missing/uncertain content that requires human verification or manual insertion.
- Do NOT fabricate data, effect sizes, ROB 2.0 judgments, or references. If unknown, use 『』.
- Use only the information present in the Inputs; if something is missing, keep 『』 placeholders.
- Methods should describe: PubMed query, feasibility scan (existing SR/MA/NMA), screening (title/abstract + full text), data extraction, ROB 2.0, and synthesis (fixed-effect).
- Results should report: PRISMA counts, study characteristics, ROB 2.0 summary, and pooled estimates (with CI) if available.

Inputs:
{json.dumps({
  "question_original": proto.get("question_original",""),
  "question_en": proto.get("question_en",""),
  "pico": proto.get("pico", {}),
  "criteria": proto.get("criteria", []),
  "schema": proto.get("schema", {}),
  "pubmed_query": st.session_state.get("pubmed_query",""),
  "feasibility": st.session_state.get("feasibility_summary", {}),
  "prisma": st.session_state.get("prisma", {}),
  "fulltext_reasons": st.session_state.get("ft_reasons", {}),
  "rob2_table": st.session_state.get("rob2_table", []),
  "meta_analysis": ma_result,
  "style_guide_excerpt": style_text[:6000] if style_text else ""
}, ensure_ascii=False, indent=2)}
"""}
            ]
            out = call_llm(msg, max_tokens=1600, timeout=90)
            js = json.loads(out)
            if isinstance(js, dict) and js:
                st.session_state["ms_sections"] = js
            else:
                raise ValueError("LLM 回傳非 dict JSON")
        except Exception as e:
            st.warning(f"LLM 寫作失敗，改用模板：{e}")
            st.session_state["ms_sections"] = manuscript_template(proto, ma_result, "en")
    else:
        st.session_state["ms_sections"] = manuscript_template(proto, ma_result, "en")

sections = st.session_state.get("ms_sections", {}) or {}
if sections:
    # show sections as expanders
    for k, v in sections.items():
        with st.expander(k, expanded=(k in ["標題","摘要","Title","Abstract"])):
            st.text_area(f"{k}", value=str(v), height=260, key=f"ms_{k}")
else:
    st.info("尚未產生稿件草稿。")

# Word export
if HAS_DOCX and sections:
    if st.button("匯出 Word（DOCX）"):
        doc = Document()
        doc.add_heading(tr("title"), level=1)
        doc.add_paragraph(tr("author"))
        doc.add_paragraph(tr("disclaimer"))

        doc.add_heading("Protocol / PICO", level=2)
        doc.add_paragraph(json.dumps(proto.get("pico", {}), ensure_ascii=False, indent=2))
        doc.add_paragraph("PubMed query:")
        doc.add_paragraph(st.session_state.get("pubmed_query",""))

        doc.add_heading("PRISMA", level=2)
        doc.add_paragraph(json.dumps(counts, ensure_ascii=False, indent=2))

        doc.add_heading("Manuscript draft", level=2)
        for k, v in sections.items():
            doc.add_heading(str(k), level=3)
            doc.add_paragraph(str(v))

        # add extraction table if exists
        ex = st.session_state.get("extract_df", pd.DataFrame())
        if isinstance(ex, pd.DataFrame) and not ex.empty:
            doc.add_heading("Extraction table (wide)", level=2)
            # only first 30 rows to keep doc reasonable
            ex2 = ex.head(30)
            t = doc.add_table(rows=1, cols=len(ex2.columns))
            hdr = t.rows[0].cells
            for j, col in enumerate(ex2.columns):
                hdr[j].text = str(col)
            for _, rr in ex2.iterrows():
                row_cells = t.add_row().cells
                for j, col in enumerate(ex2.columns):
                    row_cells[j].text = str(rr.get(col, ""))

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.download_button("下載 DOCX", data=buf.getvalue(), file_name="srma_draft.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
else:
    st.caption("若要匯出 Word：請在 requirements.txt 加上 python-docx。")
