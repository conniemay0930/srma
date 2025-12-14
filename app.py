# app.py
# =========================================================
# 一句話帶你完成 MA（BYOK）
# Author: Ya Hsin Yao
#
# 免責聲明：本工具僅供學術研究/教學用途，不構成醫療建議或法律意見；
# 使用者須自行驗證所有結果、引用與全文內容；請勿上傳可識別之病人資訊。
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
import time
import html
import hashlib
from dataclasses import dataclass
from typing import Dict, List, Tuple, Optional, Any

import requests
import pandas as pd
import streamlit as st

# Optional: PDF extraction
try:
    from PyPDF2 import PdfReader
    HAS_PYPDF2 = True
except Exception:
    HAS_PYPDF2 = False

# Optional: DOCX for template reading / export
try:
    import docx  # python-docx
    HAS_DOCX = True
except Exception:
    HAS_DOCX = False

# Optional: Plotly for forest plot (preferred)
try:
    import plotly.graph_objects as go
    HAS_PLOTLY = True
except Exception:
    HAS_PLOTLY = False


# -------------------- Page config --------------------
st.set_page_config(page_title="一句話帶你完成 MA（繁體中文）", layout="wide")


# -------------------- UI helpers --------------------
CSS = """
<style>
.small { font-size: 0.88rem; color: #555; }
.kpi { border: 1px solid #e5e7eb; border-radius: 12px; padding: 0.75rem 0.9rem; background: #fafafa; }
.kpi .label { font-size: 0.82rem; color: #6b7280; }
.kpi .value { font-size: 1.25rem; font-weight: 700; color: #111827; }
hr.soft { border: none; border-top: 1px solid #eef2f7; margin: 0.8rem 0; }
.badge { display:inline-block; padding:0.15rem 0.55rem; border-radius: 999px; font-size:0.78rem; margin-right:0.35rem; border:1px solid rgba(0,0,0,0.06); }
.badge-ok { background:#d1fae5; color:#065f46; }
.badge-warn { background:#fef3c7; color:#92400e; }
.badge-bad { background:#fee2e2; color:#991b1b; }
.card { border: 1px solid #dde2eb; border-radius: 12px; padding: 0.9rem 1rem; background:#fff; margin-bottom: 0.9rem; }
code.smallcode { font-size: 0.84rem; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

st.title("一句話帶你完成 MA")
st.caption("作者：Ya Hsin Yao　|　Language：繁體中文　|　免責聲明：僅供學術用途；請自行驗證所有結果與引用。")


# -------------------- Core helpers --------------------
def norm_text(x: str) -> str:
    if x is None:
        return ""
    x = html.unescape(str(x))
    x = re.sub(r"\s+", " ", x).strip()
    return x

def short(s: str, n: int = 120) -> str:
    s = s or ""
    return (s[:n] + "…") if len(s) > n else s

def ensure_columns(df: pd.DataFrame, cols: List[str], default: Any = "") -> pd.DataFrame:
    for c in cols:
        if c not in df.columns:
            df[c] = default
    return df

def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")

def stable_hash(text: str) -> str:
    return hashlib.sha256((text or "").encode("utf-8")).hexdigest()[:12]


# -------------------- Data models --------------------
@dataclass
class Protocol:
    P: str = ""
    I: str = ""
    C: str = ""
    O: str = ""
    NOT: str = "animal OR mice OR rat OR in vitro OR case report"

    goal_mode: str = "Fast / feasible (gap-fill)"  # or "Rigorous / narrow scope"

    # expansions
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

    def to_dict(self) -> Dict[str, Any]:
        return {
            "pico": {
                "P": self.P, "I": self.I, "C": self.C, "O": self.O, "NOT": self.NOT
            },
            "search_expansion": {
                "P_synonyms": self.P_syn or [],
                "I_synonyms": self.I_syn or [],
                "C_synonyms": self.C_syn or [],
                "O_synonyms": self.O_syn or [],
                "NOT": [x.strip() for x in (self.NOT or "").split(" OR ") if x.strip()],
            },
            "mesh_candidates": {
                "P": self.mesh_P or [],
                "I": self.mesh_I or [],
                "C": self.mesh_C or [],
                "O": self.mesh_O or [],
            },
            "goal_mode": self.goal_mode,
            "criteria": {
                "inclusion": self.inclusion,
                "exclusion": self.exclusion,
            },
            "plans": {
                "outcomes_plan": self.outcomes_plan,
                "extraction_plan": self.extraction_plan,
            }
        }


# -------------------- Session state --------------------
def init_state():
    ss = st.session_state
    ss.setdefault("lang", "繁體中文")

    ss.setdefault("byok_enabled", False)
    ss.setdefault("byok_key", "")
    ss.setdefault("byok_base_url", "https://api.openai.com/v1")
    ss.setdefault("byok_model", "gpt-4o-mini")  # user can change
    ss.setdefault("byok_temp", 0.2)

    ss.setdefault("question", "")
    ss.setdefault("protocol", Protocol(P_syn=[], I_syn=[], C_syn=[], O_syn=[], mesh_P=[], mesh_I=[], mesh_C=[], mesh_O=[]))

    ss.setdefault("pubmed_query", "")
    ss.setdefault("pubmed_records", pd.DataFrame())
    ss.setdefault("diagnostics", {})

    # screening decisions
    ss.setdefault("ta_ai", {})      # record_id -> Include/Exclude/Unsure
    ss.setdefault("ta_reason", {})  # record_id -> text
    ss.setdefault("ft_decision", {})  # record_id -> Include for meta-analysis / Exclude / Not reviewed yet
    ss.setdefault("ft_reason", {})
    ss.setdefault("ft_text", {})      # record_id -> extracted text (from PDF or paste)
    ss.setdefault("ft_note", {})      # cannot obtain note

    # extraction table
    ss.setdefault("extract_df", pd.DataFrame())

    # RoB 2.0
    ss.setdefault("rob2", {})  # record_id -> dict

    # manuscript
    ss.setdefault("manuscript_md", "")

init_state()


# -------------------- LLM client (BYOK, no secrets) --------------------
def llm_available() -> bool:
    return bool(st.session_state.get("byok_enabled")) and bool(st.session_state.get("byok_key", "").strip())

def call_openai_compatible(messages: List[Dict[str, str]], max_tokens: int = 1200) -> str:
    """
    OpenAI-compatible chat.completions REST call.
    No SDK; BYOK only; stored in session state only.
    """
    base_url = (st.session_state.get("byok_base_url") or "").strip().rstrip("/")
    api_key = (st.session_state.get("byok_key") or "").strip()
    model = (st.session_state.get("byok_model") or "").strip()
    temperature = float(st.session_state.get("byok_temp") or 0.2)

    if not (base_url and api_key and model):
        raise RuntimeError("LLM 未設定完成（base_url / key / model）。")

    url = f"{base_url}/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
    }
    r = requests.post(url, headers=headers, json=payload, timeout=60)
    if r.status_code != 200:
        raise RuntimeError(f"LLM 呼叫失敗：HTTP {r.status_code} / {r.text[:300]}")
    data = r.json()
    return data["choices"][0]["message"]["content"]


# -------------------- Sidebar (BYOK + language) --------------------
with st.sidebar:
    st.header("設定")

    st.selectbox("Language（顯示）", options=["繁體中文", "English"], key="lang")

    st.markdown("---")
    st.subheader("LLM（使用者自備 key）")

    st.checkbox("啟用 LLM（BYOK）", key="byok_enabled", help="預設關閉；關閉時流程自動降級，不會卡住。")

    st.caption("Key only used for this session; do not use on untrusted deployments; do not upload identifiable patient info.")

    st.text_input("Base URL（OpenAI-compatible）", key="byok_base_url", help="例如 https://api.openai.com/v1")
    st.text_input("Model", key="byok_model")
    st.text_input("API Key（只在本次 session）", type="password", key="byok_key")

    c1, c2 = st.columns(2)
    with c1:
        st.button("Clear key", on_click=lambda: st.session_state.update({"byok_key": ""}))
    with c2:
        st.slider("Temperature", 0.0, 1.0, 0.2, 0.05, key="byok_temp")

    st.markdown("---")
    st.subheader("寫作範本（可選，DOCX）")
    st.caption("可上傳你自己的稿件範本用於「語調提示」。請確認你擁有使用權。")
    tmpl_files = st.file_uploader("上傳 DOCX（可多份）", type=["docx"], accept_multiple_files=True)

    st.markdown("---")
    st.subheader("使用提醒")
    st.markdown(
        "- 建議輸入一句清楚問題（含介入/比較與族群）\n"
        "- 若只有縮寫（如 EDOF），請盡量補上完整詞或 lens 名稱\n"
        "- 沒有 key 也可跑到搜尋式/抓文獻/PRISMA/寬表/MA（前提：寬表有 Effect/CI）\n"
        "- ROB 2.0 與納入標準仍需人工把關"
    )


# -------------------- Template extraction --------------------
def read_docx_text(file_bytes: bytes, max_chars: int = 20000) -> str:
    if not HAS_DOCX:
        return ""
    doc = docx.Document(io.BytesIO(file_bytes))
    paras = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
    text = "\n".join(paras)
    return text[:max_chars]

def build_style_guide_from_templates(files) -> str:
    """
    Do not quote long passages; summarize style as guidance.
    """
    if not files:
        return ""
    if not HAS_DOCX:
        return ""
    combined = ""
    for f in files:
        try:
            combined += "\n" + read_docx_text(f.getvalue())
        except Exception:
            continue

    # Heuristic: extract headings presence (do not reproduce content)
    has_intro = bool(re.search(r"\bIntroduction\b|前言|背景", combined, re.IGNORECASE))
    has_methods = bool(re.search(r"\bMethods\b|Materials and Methods|方法", combined, re.IGNORECASE))
    has_results = bool(re.search(r"\bResults\b|結果", combined, re.IGNORECASE))
    has_disc = bool(re.search(r"\bDiscussion\b|討論", combined, re.IGNORECASE))

    guide = []
    guide.append("以醫學期刊系統性回顧/統合分析的正式學術口吻撰寫。")
    guide.append("結構採：摘要（可選）/Introduction/Methods/Results/Discussion/Conclusion。")
    guide.append("Methods 明確描述：資料庫、搜尋日期、納入排除、篩選流程（PRISMA）、資料萃取、風險偏倚、統計方法（fixed/random、I²）。")
    guide.append("Results 報告：檢索量與納入數、研究特徵、主要 outcome 合併效應與異質性。")
    guide.append("Discussion 以臨床意義、可能機轉、異質性來源、限制、未來研究收尾。")
    guide.append(f"範本章節偵測：Intro={has_intro}, Methods={has_methods}, Results={has_results}, Discussion={has_disc}（僅做語調參考，不複製內容）。")
    return "\n".join(guide)


STYLE_GUIDE = build_style_guide_from_templates(tmpl_files)


# -------------------- Question -> PICO parsing + expansions --------------------
ABBR_MAP = {
    # keep it small and generic; helpful but not "hard-coded example paragraphs"
    "EDOF": ["extended depth of focus", "extended depth-of-focus", "extended range of vision", "extended range-of-vision"],
    "IOL": ["intraocular lens", "intra-ocular lens"],
    "RCT": ["randomized controlled trial", "randomised controlled trial"],
}

def split_vs(question: str) -> Tuple[str, str]:
    q = question or ""
    m = re.split(r"\s+vs\.?\s+|\s+VS\.?\s+|\s+versus\s+", q, flags=re.IGNORECASE)
    if len(m) >= 2:
        left = m[0].strip()
        right = " vs ".join([x.strip() for x in m[1:]]).strip()
        return left, right
    return q.strip(), ""

def expand_terms(text: str) -> List[str]:
    """
    Rule-based expansion:
    - preserve phrases
    - add common abbreviation expansions if token matches
    - add quoted tiab variants later in query builder
    """
    text = norm_text(text)
    if not text:
        return []
    syn = []

    # split by commas / semicolons / slash
    parts = re.split(r"[;,/]+", text)
    for p in parts:
        p = p.strip()
        if not p:
            continue
        syn.append(p)

        # if exactly an abbreviation
        key = p.upper()
        if key in ABBR_MAP:
            syn.extend(ABBR_MAP[key])

        # if contains abbreviation tokens
        toks = re.findall(r"[A-Za-z]{2,10}", p)
        for t in toks:
            tu = t.upper()
            if tu in ABBR_MAP:
                syn.extend(ABBR_MAP[tu])

    # dedupe while keeping order
    out = []
    seen = set()
    for s in syn:
        s2 = s.strip()
        if not s2:
            continue
        k = s2.lower()
        if k not in seen:
            seen.add(k)
            out.append(s2)
    return out

def propose_mesh_candidates(terms: List[str]) -> List[str]:
    """
    Heuristic MeSH candidates: for stable operation without external lookups.
    If LLM enabled, we can ask LLM to propose MeSH terms explicitly.
    """
    # basic mapping heuristics
    mesh = []
    for t in terms or []:
        tl = t.lower()
        if "cataract" in tl:
            mesh.append("Cataract")
            mesh.append("Cataract Extraction")
        if "glaucoma" in tl:
            mesh.append("Glaucoma")
        if "intraocular lens" in tl or "iol" in tl or "lens" in tl:
            mesh.append("Lenses, Intraocular")
            mesh.append("Lens Implantation, Intraocular")
    # dedupe
    out = []
    seen = set()
    for m in mesh:
        k = m.lower()
        if k not in seen:
            seen.add(k)
            out.append(m)
    return out

def question_to_protocol(question: str) -> Protocol:
    q = norm_text(question)
    left, right = split_vs(q)

    # minimal inference:
    # If "vs" present, treat left as I and right as C unless the user wrote a single concept.
    P = ""
    I = left
    C = right
    O = ""

    # If both sides identical, broaden to "different types/models" rather than literal same
    if I and C and I.strip().lower() == C.strip().lower():
        C = "其他比較組（例如不同型號/設計）"

    # If question extremely short, keep P blank and force placeholders later
    proto = Protocol(P=P, I=I, C=C, O=O)
    proto.P_syn = expand_terms(proto.P)
    proto.I_syn = expand_terms(proto.I)
    proto.C_syn = expand_terms(proto.C)
    proto.O_syn = expand_terms(proto.O)

    proto.mesh_P = propose_mesh_candidates(proto.P_syn)
    proto.mesh_I = propose_mesh_candidates(proto.I_syn)
    proto.mesh_C = propose_mesh_candidates(proto.C_syn)
    proto.mesh_O = propose_mesh_candidates(proto.O_syn)

    # Plans placeholders (auto-filled later, possibly by LLM)
    proto.inclusion = ""
    proto.exclusion = ""
    proto.outcomes_plan = ""
    proto.extraction_plan = ""
    return proto


# -------------------- PubMed query builder --------------------
def quote_tiab(term: str) -> str:
    term = term.strip()
    if not term:
        return ""
    # if it already looks like a fielded query, keep it
    if "[" in term and "]" in term:
        return term
    # phrase with spaces
    return f"\"{term}\"[tiab]" if " " in term else f"{term}[tiab]"

def mesh_clause(mesh_terms: List[str]) -> str:
    items = []
    for m in mesh_terms or []:
        m = m.strip()
        if m:
            items.append(f"\"{m}\"[MeSH Terms]")
    if not items:
        return ""
    return "(" + " OR ".join(items) + ")"

def tiab_clause(syn: List[str]) -> str:
    items = []
    for s in syn or []:
        s = s.strip()
        if not s:
            continue
        items.append(quote_tiab(s))
    if not items:
        return ""
    return "(" + " OR ".join(items) + ")"

def build_pubmed_query(proto: Protocol) -> str:
    """
    Generic (not ophthalmology-locked):
    (P terms) AND (I terms) AND (C terms optional) AND (O terms optional) NOT (...)
    Each block uses: (MeSH OR tiab)
    """
    blocks = []
    def block(mesh_terms, syn_terms):
        a = mesh_clause(mesh_terms)
        b = tiab_clause(syn_terms)
        if a and b:
            return f"({a} OR {b})"
        return a or b

    P_block = block(proto.mesh_P, proto.P_syn if proto.P_syn else expand_terms(proto.P))
    I_block = block(proto.mesh_I, proto.I_syn if proto.I_syn else expand_terms(proto.I))
    C_block = block(proto.mesh_C, proto.C_syn if proto.C_syn else expand_terms(proto.C))
    O_block = block(proto.mesh_O, proto.O_syn if proto.O_syn else expand_terms(proto.O))

    # If P is empty, don't force it; but keep query meaningful.
    if P_block:
        blocks.append(P_block)
    if I_block:
        blocks.append(I_block)
    if C_block:
        blocks.append(C_block)
    if O_block:
        blocks.append(O_block)

    core = " AND ".join(blocks) if blocks else quote_tiab(proto.I or proto.P or "systematic review")
    not_block = proto.NOT.strip()
    if not_block:
        return f"({core}) NOT ({not_block})"
    return core

def build_srma_feasibility_query(pubmed_query: str) -> str:
    # find existing SR/MA on the same topic
    sr_filter = '(systematic review[pt] OR meta-analysis[pt] OR "systematic review"[tiab] OR "meta-analysis"[tiab])'
    return f"({pubmed_query}) AND {sr_filter}"


# -------------------- PubMed E-utilities --------------------
EUTILS = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils"

def pubmed_esearch(term: str, retmax: int = 200, retstart: int = 0) -> Tuple[int, List[str], str, Dict[str, Any]]:
    params = {
        "db": "pubmed",
        "term": term,
        "retmode": "json",
        "retmax": retmax,
        "retstart": retstart,
    }
    url = f"{EUTILS}/esearch.fcgi"
    r = requests.get(url, params=params, timeout=30)
    text = r.text or ""
    diag = {"status_code": r.status_code, "content_type": r.headers.get("content-type", ""), "body_head": text[:200]}
    if r.status_code != 200:
        return 0, [], r.url, diag
    try:
        data = r.json()
    except Exception:
        # sometimes returns HTML if blocked
        return 0, [], r.url, {**diag, "warning": "Non-JSON response; PubMed may be blocked or rate-limited."}
    es = data.get("esearchresult", {})
    count = int(es.get("count", 0) or 0)
    ids = es.get("idlist", []) or []
    return count, ids, r.url, diag

def pubmed_efetch_xml(pmids: List[str]) -> Tuple[str, List[str]]:
    if not pmids:
        return "", []
    chunks = []
    urls = []
    # efetch has practical length limit
    for i in range(0, len(pmids), 200):
        sub = pmids[i:i+200]
        params = {"db": "pubmed", "id": ",".join(sub), "retmode": "xml"}
        url = f"{EUTILS}/efetch.fcgi"
        r = requests.get(url, params=params, timeout=60)
        urls.append(r.url)
        if r.status_code != 200:
            continue
        chunks.append(r.text or "")
    return "\n".join(chunks), urls

def parse_pubmed_xml_minimal(xml_text: str) -> pd.DataFrame:
    """
    Minimal robust parsing via regex (no lxml dependency).
    Extract PMID, ArticleTitle, AbstractText, Year, DOI.
    """
    xml_text = xml_text or ""
    if "<PubmedArticle" not in xml_text:
        return pd.DataFrame()

    articles = re.split(r"<PubmedArticle\b", xml_text)[1:]
    rows = []
    for a in articles:
        chunk = "<PubmedArticle" + a
        pmid = re.search(r"<PMID[^>]*>(\d+)</PMID>", chunk)
        pmid = pmid.group(1) if pmid else ""
        title = re.search(r"<ArticleTitle>(.*?)</ArticleTitle>", chunk, flags=re.DOTALL)
        title = norm_text(re.sub(r"<.*?>", "", title.group(1))) if title else ""
        abst_parts = re.findall(r"<AbstractText[^>]*>(.*?)</AbstractText>", chunk, flags=re.DOTALL)
        abstract = norm_text(" ".join([re.sub(r"<.*?>", "", x) for x in abst_parts])) if abst_parts else ""
        year = ""
        y = re.search(r"<PubDate>.*?<Year>(\d{4})</Year>.*?</PubDate>", chunk, flags=re.DOTALL)
        if y:
            year = y.group(1)
        else:
            y2 = re.search(r"<PubDate>.*?<MedlineDate>(\d{4})", chunk, flags=re.DOTALL)
            year = y2.group(1) if y2 else ""

        doi = ""
        dois = re.findall(r'<ArticleId IdType="doi">(.*?)</ArticleId>', chunk, flags=re.DOTALL)
        if dois:
            doi = norm_text(dois[0])

        rows.append({
            "pmid": pmid,
            "year": year,
            "title": title,
            "abstract": abstract,
            "doi": doi,
        })
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["pmid"] = df["pmid"].astype(str)
    df["record_id"] = df["pmid"].apply(lambda x: f"PMID:{x}")
    return df


# -------------------- AI: criteria / feasibility / screening / extraction / writing --------------------
def build_protocol_llm_prompt(question: str, proto: Protocol) -> List[Dict[str, str]]:
    sys = (
        "你是資深系統性回顧/統合分析（SR/MA）研究助理。"
        "請用繁體中文輸出 JSON，協助把一句研究問題整理成可執行的 protocol。"
        "你必須：\n"
        "1) 以臨床情境推導 PICO（若不確定用『』保留），\n"
        "2) 產出可行性建議（Fast/feasible vs Rigorous），\n"
        "3) 擬定 inclusion/exclusion（寫在 PICO 層級，並說明需要人工判斷的點），\n"
        "4) 提出 outcomes 規劃（primary/secondary；若可能參考過去 RCT 常見 outcomes），\n"
        "5) 產出 extraction sheet 建議（不要寫死欄位，提出『應包含哪些欄位類別』與可能的 outcome 欄）。\n"
        "注意：不得編造不存在的研究結果；不得要求使用者提供帳密；不得輸出任何病人可識別資訊。"
    )
    user = {
        "question": question,
        "current_guess_protocol": proto.to_dict(),
    }
    return [
        {"role": "system", "content": sys},
        {"role": "user", "content": json.dumps(user, ensure_ascii=False)}
    ]

def try_llm_fill_protocol(question: str, proto: Protocol) -> Protocol:
    if not llm_available():
        return proto
    try:
        content = call_openai_compatible(build_protocol_llm_prompt(question, proto), max_tokens=1400)
        # Expect JSON
        js = json.loads(content)
        pico = js.get("pico", {}) or js.get("pico", js.get("protocol", {}).get("pico", {})) or {}
        proto.P = norm_text(pico.get("P", proto.P))
        proto.I = norm_text(pico.get("I", proto.I))
        proto.C = norm_text(pico.get("C", proto.C))
        proto.O = norm_text(pico.get("O", proto.O))
        proto.NOT = norm_text(pico.get("NOT", proto.NOT)) or proto.NOT

        proto.goal_mode = norm_text(js.get("goal_mode", proto.goal_mode)) or proto.goal_mode

        crit = js.get("criteria", {}) or {}
        proto.inclusion = norm_text(crit.get("inclusion", proto.inclusion))
        proto.exclusion = norm_text(crit.get("exclusion", proto.exclusion))

        plans = js.get("plans", {}) or {}
        proto.outcomes_plan = norm_text(plans.get("outcomes_plan", proto.outcomes_plan))
        proto.extraction_plan = norm_text(plans.get("extraction_plan", proto.extraction_plan))

        # expansions if provided
        exp = js.get("search_expansion", {}) or {}
        proto.P_syn = exp.get("P_synonyms", proto.P_syn) or proto.P_syn
        proto.I_syn = exp.get("I_synonyms", proto.I_syn) or proto.I_syn
        proto.C_syn = exp.get("C_synonyms", proto.C_syn) or proto.C_syn
        proto.O_syn = exp.get("O_synonyms", proto.O_syn) or proto.O_syn

        mesh = js.get("mesh_candidates", {}) or {}
        proto.mesh_P = mesh.get("P", proto.mesh_P) or proto.mesh_P
        proto.mesh_I = mesh.get("I", proto.mesh_I) or proto.mesh_I
        proto.mesh_C = mesh.get("C", proto.mesh_C) or proto.mesh_C
        proto.mesh_O = mesh.get("O", proto.mesh_O) or proto.mesh_O

        # ensure list types
        proto.P_syn = [norm_text(x) for x in (proto.P_syn or []) if norm_text(x)]
        proto.I_syn = [norm_text(x) for x in (proto.I_syn or []) if norm_text(x)]
        proto.C_syn = [norm_text(x) for x in (proto.C_syn or []) if norm_text(x)]
        proto.O_syn = [norm_text(x) for x in (proto.O_syn or []) if norm_text(x)]

        proto.mesh_P = [norm_text(x) for x in (proto.mesh_P or []) if norm_text(x)]
        proto.mesh_I = [norm_text(x) for x in (proto.mesh_I or []) if norm_text(x)]
        proto.mesh_C = [norm_text(x) for x in (proto.mesh_C or []) if norm_text(x)]
        proto.mesh_O = [norm_text(x) for x in (proto.mesh_O or []) if norm_text(x)]
        return proto
    except Exception:
        return proto

def rule_based_ta_screen(title: str, abstract: str, proto: Protocol) -> Tuple[str, str]:
    """
    Fast heuristic when no LLM:
    - must include at least one I term token
    - if C specified, prefer mention but not mandatory
    """
    t = (title or "").lower()
    a = (abstract or "").lower()
    blob = t + " " + a
    # exclude obvious animals
    if re.search(r"\b(mice|mouse|rat|porcine|rabbit|canine)\b", blob):
        return "Exclude", "疑似動物/非人體研究（rule-based）"
    # require I signal
    i_terms = proto.I_syn or expand_terms(proto.I)
    sig = 0
    for term in i_terms[:20]:
        if term and term.lower() in blob:
            sig += 1
            break
    if sig == 0 and proto.I:
        # maybe acronym mismatch; allow Unsure
        return "Unsure", "未偵測到明顯介入關鍵詞（rule-based）"
    return "Include", "符合基本關鍵詞訊號（rule-based）"

def ta_screen_with_llm(df: pd.DataFrame, proto: Protocol) -> Dict[str, Dict[str, str]]:
    """
    Returns mapping record_id -> {decision, reason}
    """
    out = {}
    if df.empty:
        return out
    if not llm_available():
        for _, r in df.iterrows():
            rid = r["record_id"]
            d, rs = rule_based_ta_screen(r.get("title",""), r.get("abstract",""), proto)
            out[rid] = {"decision": d, "reason": rs}
        return out

    sys = (
        "你是系統性回顧的兩位評讀者合一（先粗篩 title/abstract）。"
        "請用繁體中文輸出 JSON。"
        "規則：\n"
        "- decision 只能是 Include / Exclude / Unsure\n"
        "- reason 要簡短且可核對（例如：族群不符、介入不符、非臨床研究、不是比較研究等）\n"
        "- 不得編造全文內容\n"
    )
    records = []
    for _, r in df.iterrows():
        records.append({
            "record_id": r["record_id"],
            "title": r.get("title",""),
            "abstract": r.get("abstract",""),
            "year": r.get("year",""),
            "doi": r.get("doi",""),
        })

    user = {
        "protocol": proto.to_dict(),
        "records": records[:80],  # keep bounded
    }
    try:
        content = call_openai_compatible(
            [{"role":"system","content":sys},{"role":"user","content":json.dumps(user, ensure_ascii=False)}],
            max_tokens=1600
        )
        js = json.loads(content)
        items = js.get("decisions", js)  # allow direct dict
        if isinstance(items, dict):
            for rid, v in items.items():
                if isinstance(v, dict):
                    out[rid] = {
                        "decision": v.get("decision","Unsure"),
                        "reason": v.get("reason",""),
                    }
        elif isinstance(items, list):
            for v in items:
                rid = v.get("record_id")
                if rid:
                    out[rid] = {"decision": v.get("decision","Unsure"), "reason": v.get("reason","")}
        return out
    except Exception:
        # fallback
        for _, r in df.iterrows():
            rid = r["record_id"]
            d, rs = rule_based_ta_screen(r.get("title",""), r.get("abstract",""), proto)
            out[rid] = {"decision": d, "reason": rs}
        return out

def extraction_prompt(proto: Protocol) -> str:
    """
    强化：包含 OCR/figure/table 提示 + 動態 extraction sheet 規劃 + 過去 SR/MA 與 RCT outcomes 考量
    """
    return f"""
你是系統性回顧與統合分析的資料萃取助理。請依照下列 Protocol 抽取資料，輸出 JSON。

【Protocol（PICO + criteria）】
P: {proto.P or "『』"}
I: {proto.I or "『』"}
C: {proto.C or "『』"}
O: {proto.O or "『』"}

Inclusion（PICO 層級，若不確定用『』）：
{proto.inclusion or "『請根據現有證據/臨床情境擬定，可行性優先或嚴謹度優先都需標示』"}

Exclusion（PICO 層級，若不確定用『』）：
{proto.exclusion or "『』"}

【重要：OCR / Figure / Table】
- 若全文中 outcome 數據只在 Figure/Table：請主動提示「需要 OCR」並嘗試從表格/圖說擷取數值。
- 若 OCR 仍無法讀取，請回傳空白並在 notes 說明缺漏位置（例如：Table 2 無法辨識）。

【Extraction sheet（不要寫死欄位）】
請先規劃「本題目應有的 extraction sheet 欄位類別」，至少包含：
1) Study characteristics（作者、年份、設計、國家、樣本數、追蹤期）
2) Population baseline（年齡、性別、納入條件關鍵、重要共病）
3) Intervention/Comparator details（劑量/器材/術式/型號/時間點）
4) Outcomes：
   - 參考既有 SR/MA 可能關注的 outcomes（若有）
   - 也要考量過去 RCT 常見 primary/secondary outcomes（若有）
   - 清楚列出 primary 與 secondary outcomes 建議
5) Effect size for MA（若可）：
   - effect_measure（OR/RR/HR/MD/SMD）
   - effect（點估）
   - lower_CI / upper_CI
   - timepoint
   - unit
6) notes（缺資料、需要人工確認、表格位置）

【輸出格式】
{{
  "sheet_plan": {{
     "sections": [ ... ],
     "suggested_outcomes": {{
        "primary": [...],
        "secondary": [...]
     }}
  }},
  "extracted_fields": {{
     "...": "..."
  }},
  "meta": {{
     "effect_measure": "OR/RR/HR/MD/SMD/''",
     "effect": 0.0,
     "lower_CI": 0.0,
     "upper_CI": 0.0,
     "timepoint": "",
     "unit": ""
  }},
  "needs_ocr": true/false,
  "notes": "..."
}}
""".strip()

def rob2_ai_prompt(fulltext: str) -> str:
    return f"""
你是 Cochrane RoB 2.0 的輔助評讀者。請根據提供的全文內容，對五大 domain 與 overall 給出「建議等級」與簡短理由。
注意：
- 只能輸出建議，不得宣稱百分之百正確
- 若資訊不足，請回報 Some concerns 並說明缺口
- 以繁體中文輸出 JSON
- 等級只能是：Low risk / Some concerns / High risk

全文：
{fulltext[:18000]}
""".strip()


# -------------------- Meta-analysis + forest plot --------------------
def se_from_ci(effect: float, lcl: float, ucl: float, measure: str) -> float:
    measure = (measure or "").upper().strip()
    if measure in {"OR","RR","HR"}:
        # on log scale
        return (math.log(ucl) - math.log(lcl)) / 3.92
    # MD/SMD on linear scale
    return (ucl - lcl) / 3.92

def transform_effect(effect: float, measure: str) -> float:
    measure = (measure or "").upper().strip()
    if measure in {"OR","RR","HR"}:
        return math.log(effect)
    return effect

def inverse_transform(theta: float, measure: str) -> float:
    measure = (measure or "").upper().strip()
    if measure in {"OR","RR","HR"}:
        return math.exp(theta)
    return theta

def pool_fixed_random(effects: List[float], ses: List[float], measure: str) -> Dict[str, Any]:
    """
    Returns both fixed and DerSimonian-Laird random
    """
    k = len(effects)
    w = [1.0/(se*se) for se in ses]
    sumw = sum(w)
    theta_fixed = sum(w[i]*effects[i] for i in range(k)) / sumw
    se_fixed = math.sqrt(1.0/sumw)

    # Q, I2
    Q = sum(w[i] * (effects[i]-theta_fixed)**2 for i in range(k))
    df = max(k-1, 1)
    C = sumw - (sum(wi*wi for wi in w) / sumw)
    tau2 = max(0.0, (Q - (k-1)) / C) if C > 0 else 0.0
    w_re = [1.0/(ses[i]**2 + tau2) for i in range(k)]
    sumw_re = sum(w_re)
    theta_re = sum(w_re[i]*effects[i] for i in range(k)) / sumw_re
    se_re = math.sqrt(1.0/sumw_re)

    I2 = max(0.0, (Q - (k-1)) / Q) * 100.0 if Q > 0 else 0.0

    def ci(theta, se):
        l = theta - 1.96*se
        u = theta + 1.96*se
        return l, u

    lf, uf = ci(theta_fixed, se_fixed)
    lr, ur = ci(theta_re, se_re)

    return {
        "k": k,
        "fixed": {"theta": theta_fixed, "se": se_fixed, "lcl": lf, "ucl": uf},
        "random": {"theta": theta_re, "se": se_re, "lcl": lr, "ucl": ur, "tau2": tau2},
        "heterogeneity": {"Q": Q, "df": k-1, "I2": I2},
        "measure": measure
    }

def forest_plot_plotly(studies: List[str], eff: List[float], lcl: List[float], ucl: List[float], pooled: Tuple[float,float,float], measure: str, model_label: str):
    """
    pooled: (effect, lcl, ucl) in original scale
    """
    if not HAS_PLOTLY:
        return None

    y = list(range(len(studies)))[::-1]
    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=eff, y=y,
        mode="markers",
        name="Study",
        error_x=dict(
            type="data",
            symmetric=False,
            array=[ucl[i]-eff[i] for i in range(len(eff))],
            arrayminus=[eff[i]-lcl[i] for i in range(len(eff))],
        ),
        hovertext=studies,
    ))

    # pooled at bottom
    pooled_y = -1
    pe, pl, pu = pooled
    fig.add_trace(go.Scatter(
        x=[pe], y=[pooled_y],
        mode="markers",
        name=f"Pooled ({model_label})",
        marker=dict(symbol="diamond", size=12),
        error_x=dict(type="data", symmetric=False, array=[pu-pe], arrayminus=[pe-pl]),
        hovertext=[f"Pooled {model_label}"],
    ))

    # y labels
    ytickvals = y + [pooled_y]
    yticktext = studies[::-1] + [f"Pooled ({model_label})"]

    fig.update_layout(
        height=380 + 18*len(studies),
        xaxis_title=f"Effect ({measure})",
        yaxis=dict(
            tickmode="array",
            tickvals=ytickvals,
            ticktext=yticktext,
        ),
        margin=dict(l=10, r=10, t=35, b=10),
        showlegend=False,
    )

    # For ratio measures, add reference line at 1
    if (measure or "").upper().strip() in {"OR","RR","HR"}:
        fig.add_vline(x=1.0, line_width=1, line_dash="dash")
    else:
        fig.add_vline(x=0.0, line_width=1, line_dash="dash")
    return fig


# -------------------- Manuscript drafting --------------------
def build_manuscript_skeleton(proto: Protocol, ma_summary: Optional[Dict[str,Any]], prisma: Dict[str,Any]) -> str:
    """
    Traditional Chinese, placeholders 『』 if missing.
    """
    P = proto.P or "『』"
    I = proto.I or "『』"
    C = proto.C or "『』"
    O = proto.O or "『』"

    n_records = prisma.get("records", "『』")
    n_dups = prisma.get("duplicates_removed", "『』")
    n_screened = prisma.get("screened", "『』")
    n_fulltext = prisma.get("fulltext_assessed", "『』")
    n_included = prisma.get("included", "『』")

    # MA numbers
    if ma_summary:
        meas = ma_summary.get("measure","")
        model = ma_summary.get("model","")
        eff = ma_summary.get("effect","")
        lcl = ma_summary.get("lcl","")
        ucl = ma_summary.get("ucl","")
        I2 = ma_summary.get("I2","")
        k = ma_summary.get("k","")
    else:
        meas = model = eff = lcl = ucl = I2 = k = "『』"

    md = f"""
# 標題（請依投稿期刊格式調整）
『{I}』相較於『{C}』於『{P}』之系統性回顧與統合分析

## 免責聲明
本稿件由工具自動產生初稿，僅供學術研究與教學用途；不構成醫療建議或法律意見。所有資料、引用、數值與結論需由作者團隊逐一核對後方可使用。請勿輸入或上傳任何可識別之病人資訊。

## Introduction
白內障/相關臨床領域中，『{I}』與『{C}』常被用於『{P}』以改善『{O}』。然而，目前關於兩者在臨床成效與安全性上的相對優劣仍存在不確定性，且既有研究之設計、族群與評估時間點可能存在差異，導致結論不一致。基於上述背景，本研究旨在以系統性回顧與統合分析方式，整合現有證據以比較『{I}』與『{C}』於『{P}』之臨床結果，並評估研究間異質性及潛在限制。

## Methods
### Study design
本研究依循 PRISMA 指南進行系統性回顧與統合分析，並於研究開始前擬定 PICO、納入排除標準及統計分析計畫（若有註冊請填：『PROSPERO/registration』）。

### Eligibility criteria
- Population：{P}
- Intervention：{I}
- Comparator：{C}
- Outcomes：{O}

納入標準（PICO 層級）：
{proto.inclusion or "『請補上具體 inclusion criteria（含研究設計/族群/介入/追蹤期等）』"}

排除標準（PICO 層級）：
{proto.exclusion or "『請補上具體 exclusion criteria（含動物研究、病例報告、非比較研究等）』"}

### Information sources and search strategy
本研究以 PubMed/Medline（及其他資料庫：『EMBASE/CENTRAL/Scopus/Web of Science』）進行檢索，最後檢索日期為『YYYY-MM-DD』。完整搜尋式見附錄（Appendix）。另外，亦檢索既有系統性回顧/統合分析以評估可行性與既有證據缺口。

### Study selection
由兩位評讀者獨立進行 title/abstract 粗篩及全文審閱；分歧以討論達成共識，必要時由第三位裁決。PRISMA 流程如下：共檢索 {n_records} 篇，去除重複 {n_dups} 篇後進入篩選 {n_screened} 篇，全文審閱 {n_fulltext} 篇，最終納入 {n_included} 篇。

### Data extraction
由兩位研究者使用標準化表格獨立萃取資料，內容涵蓋研究特徵、族群基線、介入/比較組細節與 outcomes。若數據主要呈現在 figure/table，則以 OCR/人工核對方式取得；無法取得者以『』留空並註明原因。

### Risk of bias assessment
隨機對照試驗使用 RoB 2.0；觀察性研究可使用 ROBINS-I（若適用）。本工具可提供 AI 建議（可選），但最終評等必須由研究者確認。

### Statistical analysis
以反變異數法進行效應量合併。比值型指標（OR/RR/HR）以對數尺度合併後再轉回原尺度；連續型指標（MD/SMD）直接合併。異質性以 Q 與 I² 評估，必要時採隨機效應模型並探討異質性來源。統計顯著性以雙尾 p < 0.05 判定。

## Results
### Study selection
PRISMA 流程摘要：records={n_records}、duplicates removed={n_dups}、screened={n_screened}、full-text assessed={n_fulltext}、included={n_included}。

### Study characteristics
納入研究之設計、樣本數、追蹤期與介入內容整理於 Table 1（請填：『Table 1』）。

### Meta-analysis
共納入 {k} 篇研究進行合併分析。主要結果顯示（{model} 模型，{meas}）合併效應為 {eff}（95% CI {lcl}–{ucl}），異質性 I² = {I2}%。臨床意義與可能原因請見討論段落。

### Risk of bias
RoB 2.0 結果概要見 Figure『』。整體而言，主要偏倚風險來源可能包含『隨機化/盲測/缺失資料/選擇性報告』等。

## Discussion
本系統性回顧與統合分析整合現有證據，評估『{I}』與『{C}』於『{P}』之差異。主要發現顯示：『請根據合併效應與方向撰寫一句結論；若不確定用『』』。此外，研究間異質性（I² = {I2}%）提示 outcomes 定義、追蹤時間點、族群特徵或介入差異可能影響結果。臨床上，本結果可能意味著『』。

本研究限制包括：納入研究數量有限、研究設計與品質不一、outcome 報告不一致、以及 publication bias 評估之不確定性。未來仍需更高品質、標準化 outcome 的前瞻性研究，以確認『』。

## Conclusion
在『{P}』中，『{I}』相較於『{C}』在『{O}』上顯示『（優勢/無差異/仍不確定）』。本結論需結合偏倚風險與異質性審慎解讀。

## Appendix（搜尋式/補充材料）
- PubMed search string：『』
- 其他資料庫搜尋式：『』
- PRISMA checklist：『』
""".strip()
    return md

def draft_manuscript_with_llm(proto: Protocol, ma_summary: Optional[Dict[str,Any]], prisma: Dict[str,Any], style_guide: str) -> str:
    if not llm_available():
        return build_manuscript_skeleton(proto, ma_summary, prisma)

    sys = (
        "你是眼科/臨床研究領域的系統性回顧與統合分析寫作助手。"
        "請用繁體中文撰寫稿件初稿，並在任何缺資料之處用『』保留。"
        "不得捏造不存在的研究或數據；如果沒有研究特徵表/數值，就用『』。"
        "文章結構至少包含：Introduction/Methods/Results/Discussion/Conclusion/Appendix。"
        "語調請參考以下 style guide（只模仿風格，不可逐字複製）。"
    )
    user = {
        "protocol": proto.to_dict(),
        "prisma": prisma,
        "meta_analysis_summary": ma_summary or {},
        "style_guide": style_guide or "（未提供範本；請用一般醫學期刊寫作風格）"
    }
    try:
        content = call_openai_compatible(
            [{"role":"system","content":sys},{"role":"user","content":json.dumps(user, ensure_ascii=False)}],
            max_tokens=1700
        )
        return content
    except Exception:
        return build_manuscript_skeleton(proto, ma_summary, prisma)

def export_docx_from_markdown(md: str) -> Optional[bytes]:
    if not HAS_DOCX:
        return None
    # minimal markdown-to-docx: headings + paragraphs
    d = docx.Document()
    for line in (md or "").splitlines():
        line = line.rstrip()
        if not line:
            d.add_paragraph("")
            continue
        if line.startswith("# "):
            d.add_heading(line[2:].strip(), level=1)
        elif line.startswith("## "):
            d.add_heading(line[3:].strip(), level=2)
        elif line.startswith("### "):
            d.add_heading(line[4:].strip(), level=3)
        else:
            d.add_paragraph(line)
    bio = io.BytesIO()
    d.save(bio)
    return bio.getvalue()


# =========================================================
# MAIN: One-question input
# =========================================================
st.subheader("Research question（輸入一句話）")
st.session_state["question"] = st.text_input(
    "例如：『不同種類 EDOF IOL 之視覺品質與眩光/halo 比較』或『FLACS 是否優於傳統 phaco』",
    value=st.session_state.get("question",""),
)

run = st.button("Run / 執行（產出 PICO → 搜尋式 → 抓文獻 → PRISMA → MA/森林圖 → 稿件）")

if run:
    q = norm_text(st.session_state["question"])
    if not q:
        st.error("請先輸入一句研究問題。")
        st.stop()

    # 1) Protocol
    proto0 = question_to_protocol(q)
    proto = try_llm_fill_protocol(q, proto0)
    # If LLM not enabled, still enrich expansions
    proto.P_syn = proto.P_syn or expand_terms(proto.P)
    proto.I_syn = proto.I_syn or expand_terms(proto.I)
    proto.C_syn = proto.C_syn or expand_terms(proto.C)
    proto.O_syn = proto.O_syn or expand_terms(proto.O)
    proto.mesh_P = proto.mesh_P or propose_mesh_candidates(proto.P_syn)
    proto.mesh_I = proto.mesh_I or propose_mesh_candidates(proto.I_syn)
    proto.mesh_C = proto.mesh_C or propose_mesh_candidates(proto.C_syn)
    proto.mesh_O = proto.mesh_O or propose_mesh_candidates(proto.O_syn)

    st.session_state["protocol"] = proto

    # 2) PubMed query
    pub_q = build_pubmed_query(proto)
    st.session_state["pubmed_query"] = pub_q

    # 3) Retrieve PubMed records (with diagnostics)
    total, ids, es_url, es_diag = pubmed_esearch(pub_q, retmax=200, retstart=0)
    xml, ef_urls = pubmed_efetch_xml(ids[:200])
    df = parse_pubmed_xml_minimal(xml)

    st.session_state["pubmed_records"] = df
    st.session_state["diagnostics"] = {
        "pubmed_total_count": total,
        "esearch_url": es_url,
        "efetch_urls": ef_urls,
        "esearch_diag": es_diag,
        "warnings": [] if total > 0 else ["PubMed count=0：請檢查問題是否過短/縮寫未展開，或 PubMed 回應被阻擋。"],
    }

# =========================================================
# OUTPUTS
# =========================================================
if st.session_state.get("question"):
    st.markdown("---")
    st.header("Outputs")

    proto: Protocol = st.session_state.get("protocol") or Protocol(P_syn=[], I_syn=[], C_syn=[], O_syn=[], mesh_P=[], mesh_I=[], mesh_C=[], mesh_O=[])

    # Protocol display + editable advanced
    with st.expander("Protocol（PICO/criteria/plan）— 可展開修改", expanded=True):
        st.markdown("**目前 protocol（JSON）**")
        st.code(json.dumps(proto.to_dict(), ensure_ascii=False, indent=2), language="json")

        st.markdown("**可選：人工微調（建議只在必要時）**")
        c1, c2 = st.columns(2)
        with c1:
            proto.P = st.text_input("P（Population）", value=proto.P)
            proto.I = st.text_input("I（Intervention）", value=proto.I)
            proto.C = st.text_input("C（Comparison）", value=proto.C)
        with c2:
            proto.O = st.text_input("O（Outcome）", value=proto.O)
            proto.NOT = st.text_input("NOT（排除）", value=proto.NOT)
            proto.goal_mode = st.selectbox("Goal mode", options=["Fast / feasible (gap-fill)", "Rigorous / narrow scope"], index=0 if "Fast" in (proto.goal_mode or "") else 1)

        proto.inclusion = st.text_area("Inclusion criteria（PICO 層級；不確定用『』）", value=proto.inclusion, height=120)
        proto.exclusion = st.text_area("Exclusion criteria（PICO 層級；不確定用『』）", value=proto.exclusion, height=120)
        proto.outcomes_plan = st.text_area("Outcomes 規劃（primary/secondary；可含既有 SR/MA 與 RCT outcomes 考量）", value=proto.outcomes_plan, height=120)
        proto.extraction_plan = st.text_area("Extraction sheet 規劃（不要寫死欄位；描述欄位類別與 outcomes 清單）", value=proto.extraction_plan, height=120)

        st.session_state["protocol"] = proto

    # PubMed query + feasibility
    st.subheader("PubMed 搜尋式（自動產生，可複製）")
    st.code(st.session_state.get("pubmed_query",""), language="text")

    # Feasibility SR/MA existing
    feas_q = build_srma_feasibility_query(st.session_state.get("pubmed_query",""))
    with st.expander("可行性（先查既有 SR/MA 是否已有人做）", expanded=False):
        st.code(feas_q, language="text")
        try:
            cnt, _, url, _ = pubmed_esearch(feas_q, retmax=0)
            st.markdown(f"- 既有 SR/MA 相關筆數（PubMed count）：**{cnt}**")
            st.markdown(f"- esearch URL：`{url}`")
            st.caption("若 count 很高：可考慮縮小 PICO（例如族群/時間點/特定型號）；若 count 很低：可能可做 gap-fill。")
        except Exception as e:
            st.warning(f"可行性查詢失敗：{e}")

    # Diagnostics
    with st.expander("Diagnostics（PubMed 被擋/回傳非 JSON/XML 時很重要）", expanded=False):
        st.code(json.dumps(st.session_state.get("diagnostics",{}), ensure_ascii=False, indent=2), language="json")

    # Records table
    df = st.session_state.get("pubmed_records", pd.DataFrame())
    st.subheader("Records（PubMed 抓到的文獻）")
    if df is None or df.empty:
        st.info("目前沒有抓到 records。建議：把問題寫更清楚（族群 + 介入 + 比較），或在 PICO 展開處補同義詞/完整詞。")
    else:
        ensure_columns(df, ["record_id","pmid","year","title","abstract","doi"], default="")
        st.dataframe(df[["record_id","year","pmid","doi","title","abstract"]], use_container_width=True, height=360)

        # Title/Abstract screening (AI or rule-based)
        st.markdown("---")
        st.subheader("Step 3. Title/Abstract 粗篩（AI 優先；沒 LLM 就 rule-based）")

        if st.button("執行粗篩（自動填入 Include/Exclude/Unsure）"):
            decisions = ta_screen_with_llm(df, proto)
            for rid, v in decisions.items():
                st.session_state["ta_ai"][rid] = v.get("decision","Unsure")
                st.session_state["ta_reason"][rid] = v.get("reason","")

        # render screening UI
        rows = []
        for _, r in df.iterrows():
            rid = r["record_id"]
            cur = st.session_state["ta_ai"].get(rid, "")
            reason = st.session_state["ta_reason"].get(rid, "")
            rows.append({
                "record_id": rid,
                "year": r.get("year",""),
                "title": r.get("title",""),
                "decision": cur,
                "reason": reason,
            })
        sdf = pd.DataFrame(rows)
        sdf = ensure_columns(sdf, ["decision","reason"], "")

        edited = st.data_editor(
            sdf,
            use_container_width=True,
            hide_index=True,
            column_config={
                "record_id": st.column_config.TextColumn("record_id", disabled=True),
                "title": st.column_config.TextColumn("Title", disabled=True, width="large"),
                "decision": st.column_config.SelectboxColumn("TA decision", options=["", "Include", "Exclude", "Unsure"]),
                "reason": st.column_config.TextColumn("Reason", width="large"),
            }
        )

        # write back
        for _, r in edited.iterrows():
            rid = r["record_id"]
            st.session_state["ta_ai"][rid] = r.get("decision","")
            st.session_state["ta_reason"][rid] = r.get("reason","")

        # PRISMA counts (simple; can be refined)
        n_records = int(len(df))
        n_dups = 0  # PubMed IDs are unique in this prototype
        n_screened = n_records
        n_fulltext = int((edited["decision"] == "Include").sum())

        prisma = {
            "records": n_records,
            "duplicates_removed": n_dups,
            "screened": n_screened,
            "fulltext_assessed": n_fulltext,
            "included": int((edited["decision"] == "Include").sum()),  # until FT step implemented
        }

        c1, c2, c3, c4, c5 = st.columns(5)
        for col, lab, val in [
            (c1, "Records", prisma["records"]),
            (c2, "Duplicates removed", prisma["duplicates_removed"]),
            (c3, "Screened", prisma["screened"]),
            (c4, "Full-text assessed（暫以 Include 估）", prisma["fulltext_assessed"]),
            (c5, "Included（暫以 Include 估）", prisma["included"]),
        ]:
            with col:
                st.markdown(f"<div class='kpi'><div class='label'>{lab}</div><div class='value'>{val}</div></div>", unsafe_allow_html=True)

        st.markdown("---")
        st.subheader("Step 5. Data extraction（動態寬表）+ MA + 森林圖")

        # Candidate set: TA Include only
        cand_ids = edited[edited["decision"]=="Include"]["record_id"].tolist()
        cands = df[df["record_id"].isin(cand_ids)].copy()

        if cands.empty:
            st.info("尚無 Include 文獻，無法產生 extraction 寬表。")
        else:
            base = cands[["record_id","pmid","year","doi","title"]].copy()
            # core extraction cols (dynamic, user can add rows/cols)
            ensure_columns(base, [
                "Study_design","Country","N_total","Follow_up",
                "Population_key","Intervention_details","Comparator_details",
                "Primary_outcomes","Secondary_outcomes",
                "Outcome_label","Timepoint",
                "Effect_measure","Effect","Lower_CI","Upper_CI","Effect_unit",
                "Notes"
            ], default="")

            # merge previous edits
            prev = st.session_state.get("extract_df", pd.DataFrame())
            if isinstance(prev, pd.DataFrame) and (not prev.empty) and ("record_id" in prev.columns):
                # align cols
                for c in prev.columns:
                    if c not in base.columns:
                        base[c] = ""
                keep = [c for c in base.columns]
                prev2 = prev.reindex(columns=keep, fill_value="")
                base = base.merge(prev2.drop_duplicates(subset=["record_id"]), on="record_id", how="left", suffixes=("","_old"))
                for c in list(base.columns):
                    if c.endswith("_old"):
                        orig = c[:-4]
                        base[orig] = base.apply(lambda r: r[c] if str(r[c]).strip() not in ["", "nan", "None"] else r[orig], axis=1)
                        base = base.drop(columns=[c])

            st.caption("可先只填 Effect_measure + Effect + CI（與 Outcome_label/Timepoint）就能做 MA/森林圖；其餘欄位可後補。")
            ex = st.data_editor(
                base,
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                column_config={
                    "record_id": st.column_config.TextColumn("record_id", disabled=True),
                    "title": st.column_config.TextColumn("Title", disabled=True, width="large"),
                    "Effect_measure": st.column_config.SelectboxColumn("Effect measure", options=["", "OR","RR","HR","MD","SMD"]),
                    "Effect": st.column_config.NumberColumn("Effect", format="%.5f"),
                    "Lower_CI": st.column_config.NumberColumn("Lower CI", format="%.5f"),
                    "Upper_CI": st.column_config.NumberColumn("Upper CI", format="%.5f"),
                }
            )
            st.session_state["extract_df"] = ex

            st.download_button("下載 extraction 寬表（CSV）", data=to_csv_bytes(ex), file_name="extraction_wide.csv", mime="text/csv")

            # Optional AI extraction prompt display / apply
            with st.expander("（可選）AI extraction（含 OCR/figure/table 提示）", expanded=False):
                st.code(extraction_prompt(proto), language="text")
                if llm_available():
                    st.caption("若你在此頁面另外提供全文（PDF 或貼上），可再延伸做 FT/ROB2/自動回填。此版本先以 prompt 為主，保持穩定。")
                else:
                    st.info("未啟用 LLM：此區僅顯示 prompt。你可用寬表手動抽取。")

            # Meta-analysis builder
            st.markdown("---")
            st.subheader("MA（Fixed / Random effects）+ 森林圖")

            dfm = ex.copy()
            ensure_columns(dfm, ["Outcome_label","Effect_measure","Effect","Lower_CI","Upper_CI","Timepoint"], "")
            for c in ["Effect","Lower_CI","Upper_CI"]:
                dfm[c] = pd.to_numeric(dfm[c], errors="coerce")

            dfm = dfm.dropna(subset=["Effect","Lower_CI","Upper_CI"])
            dfm = dfm[dfm["Effect_measure"].astype(str).str.strip() != ""]

            if dfm.empty:
                st.info("請至少填入：Effect_measure + Effect + Lower_CI + Upper_CI（可加 Outcome_label/Timepoint）才能做 MA/森林圖。")
                ma_summary = None
            else:
                # choose outcome + measure
                outcomes = sorted([x for x in dfm["Outcome_label"].astype(str).unique().tolist() if x.strip()]) or ["(未命名 outcome)"]
                if outcomes == ["(未命名 outcome)"]:
                    dfm["Outcome_label"] = "(未命名 outcome)"
                chosen_outcome = st.selectbox("選擇 outcome", options=outcomes)
                sub = dfm[dfm["Outcome_label"]==chosen_outcome].copy()

                measures = sorted(sub["Effect_measure"].astype(str).unique().tolist())
                chosen_measure = st.selectbox("選擇 effect measure", options=measures)

                sub = sub[sub["Effect_measure"].astype(str)==chosen_measure].copy()
                if sub.empty:
                    st.warning("該 outcome 下沒有可用的 effect。")
                    ma_summary = None
                else:
                    # compute pooling
                    studies = []
                    effects_t = []
                    ses = []
                    for _, r in sub.iterrows():
                        title = r.get("title","")
                        yr = r.get("year","")
                        studies.append(f"{short(title, 60)} ({yr})")
                        eff = float(r["Effect"])
                        lcl = float(r["Lower_CI"])
                        ucl = float(r["Upper_CI"])
                        se = se_from_ci(eff, lcl, ucl, chosen_measure)
                        effects_t.append(transform_effect(eff, chosen_measure))
                        ses.append(se)

                    res = pool_fixed_random(effects_t, ses, chosen_measure)

                    model = st.radio("模型", options=["Fixed effect", "Random effects (DL)"], horizontal=True)
                    if model.startswith("Fixed"):
                        theta = res["fixed"]["theta"]
                        lth, uth = res["fixed"]["lcl"], res["fixed"]["ucl"]
                        model_label = "Fixed"
                    else:
                        theta = res["random"]["theta"]
                        lth, uth = res["random"]["lcl"], res["random"]["ucl"]
                        model_label = "Random"

                    pe = inverse_transform(theta, chosen_measure)
                    pl = inverse_transform(lth, chosen_measure)
                    pu = inverse_transform(uth, chosen_measure)

                    I2 = res["heterogeneity"]["I2"]
                    Q = res["heterogeneity"]["Q"]
                    tau2 = res["random"]["tau2"]

                    c1, c2, c3, c4 = st.columns(4)
                    with c1: st.metric("Studies (k)", res["k"])
                    with c2: st.metric("I² (%)", f"{I2:.1f}")
                    with c3: st.metric("Q", f"{Q:.2f}")
                    with c4: st.metric("tau²", f"{tau2:.4f}")

                    st.markdown(f"**Pooled effect ({model_label}, {chosen_measure})**：`{pe:.4f}`（95% CI `{pl:.4f}`–`{pu:.4f}`）")

                    # forest plot data in original scale
                    eff_os = sub["Effect"].astype(float).tolist()
                    lcl_os = sub["Lower_CI"].astype(float).tolist()
                    ucl_os = sub["Upper_CI"].astype(float).tolist()

                    if HAS_PLOTLY:
                        fig = forest_plot_plotly(studies, eff_os, lcl_os, ucl_os, (pe, pl, pu), chosen_measure, model_label)
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.warning("環境缺少 Plotly：改以表格顯示森林圖資料。")
                        show = pd.DataFrame({"study": studies, "effect": eff_os, "lcl": lcl_os, "ucl": ucl_os})
                        st.dataframe(show, use_container_width=True)

                    ma_summary = {
                        "k": res["k"],
                        "measure": chosen_measure,
                        "model": model_label,
                        "effect": f"{pe:.4f}",
                        "lcl": f"{pl:.4f}",
                        "ucl": f"{pu:.4f}",
                        "I2": f"{I2:.1f}",
                        "outcome": chosen_outcome,
                    }

            # ROB 2.0 manual + optional AI suggestion
            st.markdown("---")
            st.subheader("ROB 2.0（手動下拉；可選 AI 建議）")
            st.caption("ROB 2.0 需要標準化評分；本工具可提供建議但仍需人工核對。")

            rob_levels = ["", "Low risk", "Some concerns", "High risk"]
            domains = [
                ("D1", "Bias arising from the randomization process"),
                ("D2", "Bias due to deviations from intended interventions"),
                ("D3", "Bias due to missing outcome data"),
                ("D4", "Bias in measurement of the outcome"),
                ("D5", "Bias in selection of the reported result"),
                ("Overall", "Overall risk of bias"),
            ]

            for _, r in cands.iterrows():
                rid = r["record_id"]
                title = r.get("title","")
                with st.expander(f"ROB 2.0 - {short(title, 80)}", expanded=False):
                    cur = st.session_state["rob2"].get(rid, {}) or {}
                    updated = {}
                    for k, lab in domains:
                        updated[k] = st.selectbox(
                            lab,
                            options=rob_levels,
                            index=rob_levels.index(cur.get(k,"")) if cur.get(k,"") in rob_levels else 0,
                            key=f"rob_{rid}_{k}"
                        )
                    updated["Notes"] = st.text_area("Notes", value=cur.get("Notes",""), key=f"rob_{rid}_notes", height=80)
                    st.session_state["rob2"][rid] = updated

                    if llm_available():
                        with st.expander("（可選）AI 建議（需先貼上全文或自行上傳 PDF 後抽文字）", expanded=False):
                            st.caption("提示：此版本不強制全文流程；你可把全文貼在下面，產出建議後再自行填入上方下拉。")
                            ft = st.text_area("貼上全文（或節錄 Methods/Randomization/Masking/Outcomes）", value="", height=180, key=f"ftpaste_{rid}")
                            if st.button("產生 ROB2 建議", key=f"rob2_ai_{rid}"):
                                if not ft.strip():
                                    st.warning("請先貼上全文。")
                                else:
                                    try:
                                        content = call_openai_compatible(
                                            [{"role":"system","content":"請輸出 JSON。"},
                                             {"role":"user","content":rob2_ai_prompt(ft)}],
                                            max_tokens=900
                                        )
                                        st.code(content, language="json")
                                        st.info("請把建議轉寫到上方 ROB 2.0 下拉，並人工核對。")
                                    except Exception as e:
                                        st.error(f"AI 建議失敗：{e}")

            # export ROB2
            rob_rows = []
            for rid, d in st.session_state["rob2"].items():
                if isinstance(d, dict) and d:
                    row = {"record_id": rid}
                    row.update(d)
                    rob_rows.append(row)
            rob_df = pd.DataFrame(rob_rows) if rob_rows else pd.DataFrame()
            if not rob_df.empty:
                st.download_button("下載 ROB 2.0 表（CSV）", data=to_csv_bytes(rob_df), file_name="rob2.csv", mime="text/csv")

            # Manuscript drafting
            st.markdown("---")
            st.subheader("自動撰寫稿件（繁體中文）")
            st.caption("若資料不足會以『』保留，方便團隊後補。")

            if st.button("產生稿件初稿（Introduction/Methods/Results/Discussion）"):
                md = draft_manuscript_with_llm(proto, ma_summary, prisma, STYLE_GUIDE)
                st.session_state["manuscript_md"] = md

            md = st.session_state.get("manuscript_md","")
            if md:
                st.markdown(md)

                st.download_button(
                    "下載稿件（Markdown）",
                    data=(md or "").encode("utf-8"),
                    file_name="manuscript_draft_zhTW.md",
                    mime="text/markdown"
                )
                docx_bytes = export_docx_from_markdown(md)
                if docx_bytes:
                    st.download_button(
                        "下載稿件（DOCX）",
                        data=docx_bytes,
                        file_name="manuscript_draft_zhTW.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    st.info("環境未安裝 python-docx：已提供 Markdown 下載。")
