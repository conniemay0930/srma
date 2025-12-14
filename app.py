# app.py
# =========================================================
# 一句話帶你完成 MA（繁體中文）
# Author: Ya Hsin Yao
#
# 免責聲明：本工具僅供學術研究/教學用途，不構成醫療建議或法律意見；
# 使用者須自行驗證所有結果、引用與全文內容；請勿上傳可識別之病人資訊。
#
# 校內資源/授權提醒（重要）：
# - 若文章來自校內訂閱（付費期刊/EZproxy/館藏系統），請勿將受版權保護之全文
#   上傳至任何第三方服務或公開部署之網站（包含本 app 的雲端部署）。
# - 請遵守圖書館授權條款：避免大量下載/自動化批次擷取、避免共享全文給未授權者。
# =========================================================

from __future__ import annotations

import re
import math
import json
import html
from dataclasses import dataclass
from typing import Dict, List, Tuple, Optional, Any

import requests
import pandas as pd
import streamlit as st

# Robust XML parsing
import xml.etree.ElementTree as ET

# Optional: Plotly for forest plot (RevMan-like)
try:
    import plotly.graph_objects as go
    HAS_PLOTLY = True
except Exception:
    HAS_PLOTLY = False


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
.small { font-size: 0.9rem; color: var(--muted); }
.muted { color: var(--muted); }
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
.red { color: #b91c1c; font-weight: 650; }
.flow{ display:grid; grid-template-columns:1fr; gap:10px; }
.flow-row{ display:grid; grid-template-columns:1fr; gap:10px; }
.flow-box{ border:1px solid var(--line); border-radius:14px; padding:10px 12px; background:#fff; }
.flow-box .t{ font-weight:800; margin-bottom:2px; }
.flow-box .n{ color:var(--muted); font-size:0.92rem; }
.flow-arrow{ text-align:center; color:var(--muted); font-size:1.1rem; }
@media (min-width: 900px){ .flow-row{ grid-template-columns:1fr 1fr; gap:12px; } }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


# -------------------- Header --------------------
st.title("一句話帶你完成 MA")
st.caption("作者：Ya Hsin Yao　|　免責聲明：僅供學術用途；請自行驗證所有結果與引用。")

st.markdown(
    "<div class='notice'>"
    "<b>重要提醒（請務必閱讀）</b><br>"
    "1) 本工具輸出（含引用/數值/結論）可能不完整或不正確，<b>必須由研究者逐一核對原文</b>。<br>"
    "2) <b>請勿上傳可識別病人資訊</b>（姓名、病歷號、影像、日期等）。<br>"
    "3) <b>校內訂閱全文/館藏資源</b>可能受授權限制：避免將受版權保護的全文上傳到雲端；"
    "避免大量下載/自動化批次擷取；遵守圖書館授權條款。<br>"
    "4) 想提升檢索召回：研究問題請盡量包含『族群/情境 + 介入 + 比較 +（主要 outcome）』，縮寫請寫全名或具體型號。<br>"
    "</div>",
    unsafe_allow_html=True
)

with st.expander("目標要求", expanded=False):
    st.markdown(
        """
- 只輸入一句問題 → 自動產出：PICO/criteria、PubMed 搜尋式（MeSH+free text + 文章類型 filter）、抓文獻、
  Title/Abstract 粗篩（AI 可選 + 可人工修正）、可行性掃描（既有 SR/MA/NMA）、寬表萃取模板、MA + 森林圖、
  ROB 2.0（含理由）、稿件分段草稿。
- 預設不強迫輸入 PICO；必要時在展開區塊微調。
- 安全與授權：不鼓勵上傳受版權保護全文到雲端；避免濫用校內資源；輸出需人工核對。
        """.strip()
    )


# =========================================================
# Helpers
# =========================================================
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

def pubmed_link(pmid: str) -> str:
    pmid = str(pmid).strip()
    return f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/" if pmid else ""

def doi_link(doi: str) -> str:
    doi = (doi or "").strip()
    return f"https://doi.org/{doi}" if doi else ""

def format_abstract(text: str) -> str:
    t = (text or "").strip()
    if not t:
        return ""
    t = re.sub(r"\s*\n\s*", "\n", t)
    t = re.sub(
        r"(?<!\n)\b(PURPOSE|METHODS|RESULTS|CONCLUSIONS|CONCLUSION|BACKGROUND|DESIGN|SETTING|PATIENTS|INTERVENTION|MAIN OUTCOME MEASURES|IMPORTANCE|OBJECTIVE|DATA SOURCES|STUDY SELECTION|DATA EXTRACTION|LIMITATIONS)\s*:\s*",
        r"\n\n\1: ",
        t,
        flags=re.IGNORECASE,
    )
    if "\n" not in t and len(t) > 800:
        t = re.sub(r"(?<=\.)\s+(?=[A-Z])", "\n\n", t)
    return t.strip()


# =========================================================
# Protocol
# =========================================================
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
        }


# =========================================================
# Session state
# =========================================================
def init_state():
    ss = st.session_state

    # BYOK
    ss.setdefault("byok_enabled", False)
    ss.setdefault("byok_key", "")
    ss.setdefault("byok_base_url", "https://api.openai.com/v1")
    ss.setdefault("byok_model", "gpt-4o-mini")
    ss.setdefault("byok_temp", 0.2)
    ss.setdefault("byok_consent", False)

    # inputs
    ss.setdefault("question", "")
    ss.setdefault("article_type", "不限")
    ss.setdefault("custom_pubmed_filter", "")

    # artifacts
    ss.setdefault("protocol", Protocol(P_syn=[], I_syn=[], C_syn=[], O_syn=[], mesh_P=[], mesh_I=[], mesh_C=[], mesh_O=[]))
    ss.setdefault("pubmed_query", "")
    ss.setdefault("feas_query", "")
    ss.setdefault("pubmed_records", pd.DataFrame())
    ss.setdefault("srma_hits", pd.DataFrame())
    ss.setdefault("diagnostics", {})

    # TA screening
    ss.setdefault("ta_ai", {})
    ss.setdefault("ta_ai_reason", {})
    ss.setdefault("ta_ai_conf", {})
    ss.setdefault("ta_override", {})
    ss.setdefault("ta_override_reason", {})

    # extraction
    ss.setdefault("extract_df", pd.DataFrame())
    ss.setdefault("extract_saved", False)

    # MA
    ss.setdefault("ma_outcome_input", "")
    ss.setdefault("ma_measure_choice", "")
    ss.setdefault("ma_model_choice", "Fixed effect")
    ss.setdefault("ma_last_result", None)  # cache dict
    ss.setdefault("ma_skipped_rows", pd.DataFrame())

    # ROB2
    ss.setdefault("rob2", {})

    # Manuscript
    ss.setdefault("ms_sections", {})

init_state()


# =========================================================
# BYOK LLM
# =========================================================
def llm_available() -> bool:
    return bool(st.session_state.get("byok_enabled")) and bool(st.session_state.get("byok_key", "").strip()) and bool(st.session_state.get("byok_consent"))

def call_openai_compatible(messages: List[Dict[str, str]], max_tokens: int = 1400) -> str:
    base_url = (st.session_state.get("byok_base_url") or "").strip().rstrip("/")
    api_key = (st.session_state.get("byok_key") or "").strip()
    model = (st.session_state.get("byok_model") or "").strip()
    temperature = float(st.session_state.get("byok_temp") or 0.2)

    if not (base_url and api_key and model):
        raise RuntimeError("LLM 未設定完成（base_url / key / model）。")

    url = f"{base_url}/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {"model": model, "messages": messages, "temperature": temperature, "max_tokens": max_tokens}

    r = requests.post(url, headers=headers, json=payload, timeout=75)
    if r.status_code != 200:
        raise RuntimeError(f"LLM 呼叫失敗：HTTP {r.status_code} / {r.text[:300]}")
    data = r.json()
    return data["choices"][0]["message"]["content"]


# =========================================================
# Sidebar
# =========================================================
with st.sidebar:
    st.header("設定")
    st.subheader("LLM（使用者自備 key）")
    st.checkbox("啟用 LLM（BYOK）", key="byok_enabled")
    st.session_state["byok_consent"] = st.checkbox(
        "我理解並同意：僅供學術用途；輸出需人工核對；不輸入病人資訊；不違反校內授權。",
        value=bool(st.session_state.get("byok_consent", False)),
    )
    st.text_input("Base URL（OpenAI-compatible）", key="byok_base_url")
    st.text_input("Model", key="byok_model")
    st.text_input("API Key（只在本次 session）", type="password", key="byok_key")
    st.slider("Temperature", 0.0, 1.0, 0.2, 0.05, key="byok_temp")
    st.button("Clear key", on_click=lambda: st.session_state.update({"byok_key": ""}))

    st.markdown("---")
    st.subheader("顯示選項")
    st.checkbox("逐篇卡片顯示（推薦）", value=True, key="show_record_cards")


# =========================================================
# PICO parsing + expansions
# =========================================================
ABBR_MAP = {
    "EDOF": ["extended depth of focus", "extended depth-of-focus", "extended range of vision", "extended range-of-vision"],
    "IOL": ["intraocular lens", "intra-ocular lens"],
    "RCT": ["randomized controlled trial", "randomised controlled trial"],
    "NMA": ["network meta-analysis", "network meta analysis"],
    "FLACS": ["femtosecond laser-assisted cataract surgery", "femtosecond laser assisted cataract surgery"],
    "PHACO": ["phacoemulsification"],
}

ARTICLE_TYPE_FILTERS = {
    "不限": "",
    "RCT": "randomized controlled trial[pt] OR controlled clinical trial[pt] OR randomized[tiab] OR randomised[tiab]",
    "SR/MA": "systematic review[pt] OR meta-analysis[pt] OR \"systematic review\"[tiab] OR \"meta-analysis\"[tiab]",
    "NMA": "\"network meta-analysis\"[tiab] OR network meta analysis[tiab] OR NMA[tiab]",
    "Cohort": "cohort studies[MeSH Terms] OR cohort[tiab]",
    "Case-control": "case-control studies[MeSH Terms] OR case control[tiab]",
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
    text = norm_text(text)
    if not text:
        return []
    syn: List[str] = []
    parts = re.split(r"[;,/]+", text)
    for p in parts:
        p = p.strip()
        if not p:
            continue
        syn.append(p)
        key = p.upper()
        if key in ABBR_MAP:
            syn.extend(ABBR_MAP[key])
        toks = re.findall(r"[A-Za-z]{2,18}", p)
        for t in toks:
            tu = t.upper()
            if tu in ABBR_MAP:
                syn.extend(ABBR_MAP[tu])

    out, seen = [], set()
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
    mesh = []
    for t in terms or []:
        tl = t.lower()
        if "cataract" in tl:
            mesh += ["Cataract", "Cataract Extraction"]
        if "glaucoma" in tl:
            mesh += ["Glaucoma"]
        if "intraocular lens" in tl or "iol" in tl or "lens" in tl:
            mesh += ["Lenses, Intraocular", "Lens Implantation, Intraocular"]
    out, seen = [], set()
    for m in mesh:
        k = m.lower()
        if k not in seen:
            seen.add(k)
            out.append(m)
    return out

def question_to_protocol(question: str) -> Protocol:
    q = norm_text(question)
    left, right = split_vs(q)
    proto = Protocol(P="", I=left, C=right, O="")
    if proto.I and proto.C and proto.I.strip().lower() == proto.C.strip().lower():
        proto.C = "其他比較組（例如不同型號/設計）"
    proto.P_syn = expand_terms(proto.P)
    proto.I_syn = expand_terms(proto.I)
    proto.C_syn = expand_terms(proto.C)
    proto.O_syn = expand_terms(proto.O)
    proto.mesh_P = propose_mesh_candidates(proto.P_syn)
    proto.mesh_I = propose_mesh_candidates(proto.I_syn)
    proto.mesh_C = propose_mesh_candidates(proto.C_syn)
    proto.mesh_O = propose_mesh_candidates(proto.O_syn)
    return proto


# =========================================================
# PubMed query builder
# =========================================================
def quote_tiab(term: str) -> str:
    term = term.strip()
    if not term:
        return ""
    if "[" in term and "]" in term:
        return term
    return f"\"{term}\"[tiab]" if " " in term else f"{term}[tiab]"

def mesh_clause(mesh_terms: List[str]) -> str:
    items = []
    for m in mesh_terms or []:
        m = m.strip()
        if m:
            items.append(f"\"{m}\"[MeSH Terms]")
    return "(" + " OR ".join(items) + ")" if items else ""

def tiab_clause(syn: List[str]) -> str:
    items = []
    for s in syn or []:
        s = s.strip()
        if not s:
            continue
        items.append(quote_tiab(s))
    return "(" + " OR ".join(items) + ")" if items else ""

def build_pubmed_query(proto: Protocol, article_type: str, custom_filter: str) -> str:
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

    if P_block: blocks.append(P_block)
    if I_block: blocks.append(I_block)
    if C_block: blocks.append(C_block)
    if O_block: blocks.append(O_block)

    core = " AND ".join(blocks) if blocks else quote_tiab(proto.I or proto.P or "systematic review")
    not_block = (proto.NOT or "").strip()
    q = f"({core}) NOT ({not_block})" if not_block else core

    atf = (ARTICLE_TYPE_FILTERS.get(article_type, "") or "").strip()
    if atf:
        q = f"({q}) AND ({atf})"

    custom_filter = (custom_filter or "").strip()
    if custom_filter:
        q = f"({q}) AND ({custom_filter})"
    return q

def build_feasibility_query(pubmed_query: str) -> str:
    sr_filter = '(systematic review[pt] OR meta-analysis[pt] OR "systematic review"[tiab] OR "meta-analysis"[tiab] OR "network meta-analysis"[tiab] OR NMA[tiab])'
    return f"({pubmed_query}) AND {sr_filter}"


# =========================================================
# PubMed E-utilities
# =========================================================
EUTILS = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils"

def pubmed_esearch(term: str, retmax: int = 200, retstart: int = 0) -> Tuple[int, List[str], str, Dict[str, Any]]:
    params = {"db": "pubmed", "term": term, "retmode": "json", "retmax": retmax, "retstart": retstart}
    url = f"{EUTILS}/esearch.fcgi"
    r = requests.get(url, params=params, timeout=30)
    text = r.text or ""
    diag = {"status_code": r.status_code, "content_type": r.headers.get("content-type", ""), "body_head": text[:250]}
    if r.status_code != 200:
        return 0, [], r.url, diag
    try:
        data = r.json()
    except Exception:
        return 0, [], r.url, {**diag, "warning": "Non-JSON response; PubMed may be blocked or rate-limited."}
    es = data.get("esearchresult", {})
    count = int(es.get("count", 0) or 0)
    ids = es.get("idlist", []) or []
    return count, ids, r.url, diag

def pubmed_efetch_xml(pmids: List[str]) -> Tuple[List[str], List[str]]:
    """
    回傳「多段 XML 文件」list，避免把多個 XML doc 串起來導致 parse 失敗。
    """
    if not pmids:
        return [], []
    docs, urls = [], []
    for i in range(0, len(pmids), 200):
        sub = pmids[i:i+200]
        params = {"db": "pubmed", "id": ",".join(sub), "retmode": "xml"}
        url = f"{EUTILS}/efetch.fcgi"
        r = requests.get(url, params=params, timeout=60)
        urls.append(r.url)
        if r.status_code != 200:
            continue
        docs.append(r.text or "")
    return docs, urls


# -------------------- XML parsing (ElementTree) --------------------
def _get_text(node: Optional[ET.Element]) -> str:
    if node is None:
        return ""
    return norm_text("".join(node.itertext()))

def _find_first_author(article_el: ET.Element) -> str:
    # Try individual author
    auth_list = article_el.find(".//AuthorList")
    if auth_list is not None:
        for a in auth_list.findall(".//Author"):
            last = _get_text(a.find("LastName"))
            fore = _get_text(a.find("ForeName"))
            ini = _get_text(a.find("Initials"))
            coll = _get_text(a.find("CollectiveName"))
            if last:
                if ini:
                    return f"{last} {ini}".strip()
                if fore:
                    return f"{last} {fore}".strip()
                return last
            if coll:
                return coll
    return ""

def _find_year(pubmed_article: ET.Element) -> str:
    # Most common: ArticleDate or PubDate Year
    for path in [
        ".//PubDate/Year",
        ".//ArticleDate/Year",
        ".//PubDate/MedlineDate",
    ]:
        el = pubmed_article.find(path)
        if el is not None:
            t = _get_text(el)
            m = re.search(r"(\d{4})", t)
            return m.group(1) if m else ""
    return ""

def _find_doi(pubmed_article: ET.Element) -> str:
    for aid in pubmed_article.findall(".//ArticleId"):
        if aid.attrib.get("IdType") == "doi":
            return _get_text(aid)
    return ""

def _find_journal(pubmed_article: ET.Element) -> str:
    j = pubmed_article.find(".//Journal/Title")
    return _get_text(j)

def _find_title(pubmed_article: ET.Element) -> str:
    t = pubmed_article.find(".//ArticleTitle")
    return _get_text(t)

def _find_abstract(pubmed_article: ET.Element) -> str:
    parts = []
    for ab in pubmed_article.findall(".//Abstract/AbstractText"):
        label = ab.attrib.get("Label") or ab.attrib.get("NlmCategory") or ""
        txt = _get_text(ab)
        if not txt:
            continue
        if label:
            parts.append(f"{label}: {txt}")
        else:
            parts.append(txt)
    return "\n\n".join(parts).strip()

def _find_pmid(pubmed_article: ET.Element) -> str:
    pmid_el = pubmed_article.find(".//PMID")
    return _get_text(pmid_el)

def parse_pubmed_xml_minimal(xml_docs: List[str]) -> pd.DataFrame:
    """
    解析多段 PubMed XML 文件（ElementTree）
    欄位：pmid, year, title, abstract, doi, journal, first_author
    """
    rows = []
    for xml_text in xml_docs or []:
        if not xml_text or "<PubmedArticle" not in xml_text:
            continue
        # Some blocked responses are HTML
        if "<html" in xml_text.lower() and "<PubmedArticle" not in xml_text:
            continue

        try:
            root = ET.fromstring(xml_text)
        except Exception:
            # Try to salvage by trimming to PubmedArticleSet
            m = re.search(r"(<PubmedArticleSet[\s\S]*</PubmedArticleSet>)", xml_text)
            if not m:
                continue
            try:
                root = ET.fromstring(m.group(1))
            except Exception:
                continue

        for art in root.findall(".//PubmedArticle"):
            pmid = _find_pmid(art)
            if not pmid:
                continue
            title = _find_title(art)
            abstract = _find_abstract(art)
            year = _find_year(art)
            doi = _find_doi(art)
            journal = _find_journal(art)
            # first author needs Article element
            article_el = art.find(".//Article")
            first_author = _find_first_author(article_el) if article_el is not None else ""

            rows.append({
                "pmid": str(pmid),
                "year": year,
                "title": title,
                "abstract": abstract,
                "doi": doi,
                "journal": journal,
                "first_author": first_author,
                "record_id": f"PMID:{pmid}",
                "source": "PubMed"
            })

    df = pd.DataFrame(rows)
    return df


# =========================================================
# Title/Abstract screening (LLM optional)
# =========================================================
def heuristic_screen(title: str, abstract: str, proto: Protocol) -> Tuple[str, str, float]:
    title = title or ""
    abstract = abstract or ""
    blob = (title + " " + abstract).lower()

    if re.search(r"\b(mice|mouse|rat|rabbit|porcine|canine|in vitro)\b", blob):
        return "Exclude", "偵測到動物/體外相關字樣；通常不符合人體臨床 MA 納入（需人工確認）。", 0.86
    if re.search(r"\b(case report|case series)\b", blob):
        return "Exclude", "偵測到病例報告/病例系列字樣；多數 MA 會排除（需人工確認）。", 0.80

    i_terms = proto.I_syn or expand_terms(proto.I)
    c_terms = proto.C_syn or expand_terms(proto.C)

    hits_i = [t for t in i_terms[:60] if t and t.lower() in blob]
    hits_c = [t for t in c_terms[:60] if t and t.lower() in blob] if c_terms else []

    trial_like = bool(re.search(r"\b(randomized|randomised|randomly|trial|controlled|double-blind|single-blind)\b", blob))

    if hits_i and (trial_like or hits_c):
        reason = (
            "研究型態訊號："
            + ("疑似試驗/比較研究（randomized/trial/controlled/blind 等）" if trial_like else "可能為比較研究（未明確 randomized 訊號）")
            + "；"
            + f"介入/主題命中：{', '.join(hits_i[:6])}"
            + (f"；比較命中：{', '.join(hits_c[:5])}" if hits_c else "")
            + "。建議先保留進入 full-text。"
        )
        conf = 0.78 if trial_like else 0.65
        return "Include", reason, conf

    if hits_i:
        return "Unsure", f"命中介入/主題關鍵詞：{', '.join(hits_i[:6])}；但資訊不足以確認研究設計/比較組，建議人工檢視。", 0.55

    if len(blob.strip()) < 80:
        return "Unsure", "摘要資訊過少或僅短句/縮寫，無法可靠判讀；建議人工檢視。", 0.40

    return "Unsure", "未偵測到足夠的 PICO 關鍵詞或研究設計訊號；建議人工快速掃描以免漏掉。", 0.45

def screen_with_llm(records: List[Dict[str, Any]], proto: Protocol) -> Dict[str, Dict[str, Any]]:
    out: Dict[str, Dict[str, Any]] = {}

    if not llm_available():
        for r in records:
            rid = r["record_id"]
            d, rs, cf = heuristic_screen(r.get("title",""), r.get("abstract",""), proto)
            out[rid] = {"decision": d, "reason": rs, "confidence": cf}
        return out

    sys = (
        "你是資深系統性回顧研究助理，負責 Title/Abstract 粗篩（第一輪）。"
        "請以繁體中文輸出 JSON，且不得夾雜任何多餘文字。\n"
        "輸出格式：{decisions: [{record_id, decision, reason, confidence} ...]}\n"
        "decision 只能是 Include / Exclude / Unsure。\n"
        "reason 請清楚敘述：研究設計、介入/比較是否命中、族群是否合理、為何保留或排除。\n"
        "confidence 0~1。\n"
        "若資訊不足，請選 Unsure；不得捏造全文內容。"
    )
    user = {"protocol": proto.to_dict(), "records": records[:120]}

    try:
        content = call_openai_compatible(
            [{"role":"system","content":sys},{"role":"user","content":json.dumps(user, ensure_ascii=False)}],
            max_tokens=1900
        )
        js = json.loads(content)
        items = js.get("decisions", [])
        if isinstance(items, list):
            for it in items:
                rid = it.get("record_id")
                if not rid:
                    continue
                out[rid] = {
                    "decision": it.get("decision","Unsure"),
                    "reason": it.get("reason",""),
                    "confidence": float(it.get("confidence", 0.5) or 0.5),
                }
        for r in records:
            rid = r["record_id"]
            if rid not in out:
                d, rs, cf = heuristic_screen(r.get("title",""), r.get("abstract",""), proto)
                out[rid] = {"decision": d, "reason": rs, "confidence": cf}
        return out
    except Exception:
        for r in records:
            rid = r["record_id"]
            d, rs, cf = heuristic_screen(r.get("title",""), r.get("abstract",""), proto)
            out[rid] = {"decision": d, "reason": rs, "confidence": cf}
        return out


# =========================================================
# PRISMA (HTML)
# =========================================================
def compute_effective_decision(rid: str) -> str:
    od = (st.session_state["ta_override"].get(rid, "") or "").strip()
    ai = st.session_state["ta_ai"].get(rid, "Unsure")
    return od if od else ai

def compute_prisma(df: pd.DataFrame) -> Dict[str, Any]:
    if df is None or df.empty:
        return {"records_identified": 0, "records_excluded": 0, "fulltext_assessed": 0, "studies_included": 0, "included_meta": 0, "unsure_fulltext": 0}
    rids = df["record_id"].tolist()
    eff = [compute_effective_decision(r) for r in rids]
    total = len(rids)
    excluded = sum(1 for x in eff if x == "Exclude")
    included = sum(1 for x in eff if x == "Include")
    unsure = sum(1 for x in eff if x == "Unsure")
    fulltext_assessed = included + unsure
    return {
        "records_identified": total,
        "records_excluded": excluded,
        "fulltext_assessed": fulltext_assessed,
        "studies_included": included,
        "included_meta": included,
        "unsure_fulltext": unsure
    }

def render_prisma_html(pr: Dict[str, Any]):
    n_id = pr.get("records_identified", 0)
    n_exc = pr.get("records_excluded", 0)
    n_ft = pr.get("fulltext_assessed", 0)
    n_unsure = pr.get("unsure_fulltext", 0)
    n_inc = pr.get("studies_included", 0)
    n_meta = pr.get("included_meta", 0)

    st.markdown(
        f"""
<div class="flow">
  <div class="flow-box"><div class="t">Records identified</div><div class="n">n = {n_id}</div></div>
  <div class="flow-arrow">↓</div>
  <div class="flow-row">
    <div class="flow-box"><div class="t">Records screened (Title/Abstract)</div><div class="n">n = {n_id}</div></div>
    <div class="flow-box"><div class="t">Records excluded</div><div class="n">n = {n_exc}</div></div>
  </div>
  <div class="flow-arrow">↓</div>
  <div class="flow-row">
    <div class="flow-box"><div class="t">Full-text assessed for eligibility</div><div class="n">n = {n_ft}（Unsure：{n_unsure}）</div></div>
    <div class="flow-box"><div class="t">Studies included in meta-analysis</div><div class="n">n = {n_meta}</div></div>
  </div>
  <div class="flow-arrow">↓</div>
  <div class="flow-box"><div class="t">Studies included (qualitative synthesis)</div><div class="n">n = {n_inc}</div></div>
</div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# MA + forest plot
# =========================================================
RATIO_MEASURES = {"OR", "RR", "HR"}

def se_from_ci_safe(effect: float, lcl: float, ucl: float, measure: str) -> Tuple[Optional[float], Optional[str]]:
    m = (measure or "").upper().strip()
    if any(x is None or (isinstance(x, float) and math.isnan(x)) for x in [effect, lcl, ucl]):
        return None, "缺少 effect/CI"
    if ucl <= lcl:
        return None, "Upper CI 必須大於 Lower CI"
    if m in RATIO_MEASURES:
        if effect <= 0 or lcl <= 0 or ucl <= 0:
            return None, f"{m} 的 effect/CI 必須皆 > 0（否則無法取 log）"
        try:
            return (math.log(ucl) - math.log(lcl)) / 3.92, None
        except ValueError:
            return None, f"{m} 的 CI 無法取 log（請檢查是否 <=0）"
    return (ucl - lcl) / 3.92, None

def transform_effect(effect: float, measure: str) -> float:
    m = (measure or "").upper().strip()
    return math.log(effect) if m in RATIO_MEASURES else effect

def inverse_transform(theta: float, measure: str) -> float:
    m = (measure or "").upper().strip()
    return math.exp(theta) if m in RATIO_MEASURES else theta

def pool_fixed(effects: List[float], ses: List[float]) -> Tuple[float, float]:
    w = [1.0/(se*se) for se in ses]
    sumw = sum(w)
    theta = sum(w[i]*effects[i] for i in range(len(effects))) / sumw
    se = math.sqrt(1.0/sumw)
    return theta, se

def pool_random_DL(effects: List[float], ses: List[float]) -> Tuple[float, float, float, float, float]:
    k = len(effects)
    w = [1.0/(se*se) for se in ses]
    sumw = sum(w)
    theta_fixed = sum(w[i]*effects[i] for i in range(k)) / sumw
    Q = sum(w[i] * (effects[i]-theta_fixed)**2 for i in range(k))
    C = sumw - (sum(wi*wi for wi in w) / sumw)
    tau2 = max(0.0, (Q - (k-1)) / C) if (C > 0 and k > 1) else 0.0
    w_re = [1.0/(ses[i]**2 + tau2) for i in range(k)]
    sumw_re = sum(w_re)
    theta_re = sum(w_re[i]*effects[i] for i in range(k)) / sumw_re
    se_re = math.sqrt(1.0/sumw_re)
    I2 = max(0.0, (Q - (k-1)) / Q) * 100.0 if (Q > 0 and k > 1) else 0.0
    return theta_re, se_re, Q, I2, tau2

def forest_plot_revman(studies: List[str], eff: List[float], lcl: List[float], ucl: List[float],
                       weights: List[float], pooled: Tuple[float,float,float],
                       measure: str, model_label: str):
    if not HAS_PLOTLY:
        return None
    wmax = max(weights) if weights else 1.0
    sizes = [10 + 20*(wi/wmax) for wi in weights]
    y = list(range(len(studies)))[::-1]
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=eff, y=y, mode="markers",
        marker=dict(symbol="square", size=sizes, line=dict(width=1)),
        error_x=dict(type="data", symmetric=False,
                     array=[ucl[i]-eff[i] for i in range(len(eff))],
                     arrayminus=[eff[i]-lcl[i] for i in range(len(eff))]),
        showlegend=False
    ))
    pe, pl, pu = pooled
    pooled_y = -1
    fig.add_trace(go.Scatter(
        x=[pe], y=[pooled_y], mode="markers",
        marker=dict(symbol="diamond", size=14, line=dict(width=1)),
        error_x=dict(type="data", symmetric=False, array=[pu-pe], arrayminus=[pe-pl]),
        showlegend=False
    ))
    null = 1.0 if (measure or "").upper().strip() in RATIO_MEASURES else 0.0
    fig.add_vline(x=null, line_width=1, line_dash="dash")
    fig.update_layout(
        height=360 + 20*len(studies),
        xaxis_title=f"Effect ({measure})",
        yaxis=dict(tickmode="array", tickvals=y + [pooled_y], ticktext=studies[::-1] + [f"Pooled ({model_label})"]),
        margin=dict(l=10, r=10, t=35, b=10),
        showlegend=False,
    )
    return fig


# =========================================================
# UI: Inputs
# =========================================================
st.subheader("Research question（輸入一句話）")
colA, colB = st.columns([0.62, 0.38])
with colA:
    st.session_state["question"] = st.text_input(
        "例：『不同種類 EDOF IOL（A vs B）在白內障術後中距離視力與眩光比較』或『FLACS 是否優於傳統 phaco』",
        value=st.session_state.get("question",""),
    )
with colB:
    st.session_state["article_type"] = st.selectbox("文章類型（可選）", options=list(ARTICLE_TYPE_FILTERS.keys()),
                                                    index=list(ARTICLE_TYPE_FILTERS.keys()).index(st.session_state.get("article_type","不限"))
                                                    if st.session_state.get("article_type","不限") in ARTICLE_TYPE_FILTERS else 0)
    st.session_state["custom_pubmed_filter"] = st.text_input("自訂 PubMed filter（可選）",
                                                            value=st.session_state.get("custom_pubmed_filter",""),
                                                            help="例如：humans[MeSH Terms] OR human[tiab]；年份：2020:3000[pdat]")
run = st.button("Run / 執行（自動跑到 Outputs）", type="primary")


# =========================================================
# Pipeline run
# =========================================================
if run:
    q = norm_text(st.session_state["question"])
    if not q:
        st.error("請先輸入一句研究問題。")
        st.stop()

    with st.spinner("Step 0/4：生成 protocol…"):
        proto = question_to_protocol(q)
        st.session_state["protocol"] = proto

    with st.spinner("Step 1/4：產出 PubMed 搜尋式…"):
        pub_q = build_pubmed_query(proto, st.session_state["article_type"], st.session_state["custom_pubmed_filter"])
        st.session_state["pubmed_query"] = pub_q

    with st.spinner("Step 2/4：可行性掃描（既有 SR/MA/NMA）…"):
        feas_q = build_feasibility_query(st.session_state["pubmed_query"])
        st.session_state["feas_query"] = feas_q
        cnt_feas, ids_feas, feas_url, feas_diag = pubmed_esearch(feas_q, retmax=20, retstart=0)
        feas_docs, _ = pubmed_efetch_xml(ids_feas[:20])
        df_feas = parse_pubmed_xml_minimal(feas_docs)
        st.session_state["srma_hits"] = df_feas
        st.session_state["diagnostics"] = {"feasibility": {"count": cnt_feas, "esearch_url": feas_url, "diag": feas_diag}}

    with st.spinner("Step 3/4：抓取 PubMed 文獻…"):
        total, ids, es_url, es_diag = pubmed_esearch(st.session_state["pubmed_query"], retmax=200, retstart=0)
        docs, ef_urls = pubmed_efetch_xml(ids[:200])
        df = parse_pubmed_xml_minimal(docs)
        st.session_state["pubmed_records"] = df

        d = st.session_state.get("diagnostics", {}) or {}
        d.update({
            "pubmed_total_count": total,
            "esearch_url": es_url,
            "efetch_urls": ef_urls,
            "esearch_diag": es_diag,
            "warnings": [] if total > 0 else ["PubMed count=0：請把問題寫更具體（縮寫寫全名/加型號/族群/outcome），或看 Diagnostics 是否被阻擋。"],
        })
        st.session_state["diagnostics"] = d

    with st.spinner("Step 4/4：Title/Abstract 粗篩（AI 可選）…"):
        df = st.session_state.get("pubmed_records", pd.DataFrame())
        if df is not None and not df.empty:
            recs = []
            for _, r in df.iterrows():
                recs.append({
                    "record_id": r["record_id"],
                    "title": r.get("title",""),
                    "abstract": r.get("abstract",""),
                    "year": r.get("year",""),
                    "doi": r.get("doi",""),
                    "journal": r.get("journal",""),
                    "first_author": r.get("first_author",""),
                })
            results = screen_with_llm(recs, st.session_state["protocol"])
            for rid, v in results.items():
                st.session_state["ta_ai"][rid] = v.get("decision","Unsure")
                st.session_state["ta_ai_reason"][rid] = v.get("reason","")
                st.session_state["ta_ai_conf"][rid] = float(v.get("confidence", 0.5) or 0.5)

    st.success("Done。請往下查看 Outputs。")


# =========================================================
# Outputs
# =========================================================
if st.session_state.get("question"):
    proto: Protocol = st.session_state.get("protocol")
    df = st.session_state.get("pubmed_records", pd.DataFrame())
    df_feas = st.session_state.get("srma_hits", pd.DataFrame())
    diag = st.session_state.get("diagnostics", {}) or {}

    tabs = st.tabs([
        "總覽（PRISMA）",
        "Step 1 搜尋式",
        "Step 2 可行性（SR/MA/NMA）",
        "Step 3+4 Records + 粗篩（合併）",
        "Step 5 萃取（寬表）",
        "Step 6 MA + 森林圖（按鈕更新）",
        "Diagnostics"
    ])

    with tabs[0]:
        pr = compute_prisma(df) if df is not None else {}
        total = int(diag.get("pubmed_total_count", 0) or 0)
        feas_cnt = int((diag.get("feasibility", {}) or {}).get("count", 0) or 0)
        includes = 0
        unsure = 0
        excluded = 0
        if df is not None and not df.empty:
            for rid in df["record_id"].tolist():
                eff = compute_effective_decision(rid)
                includes += (eff == "Include")
                unsure += (eff == "Unsure")
                excluded += (eff == "Exclude")

        c1, c2, c3, c4 = st.columns(4)
        with c1: st.markdown(f"<div class='kpi'><div class='label'>PubMed count</div><div class='value'>{total}</div></div>", unsafe_allow_html=True)
        with c2: st.markdown(f"<div class='kpi'><div class='label'>既有 SR/MA/NMA</div><div class='value'>{feas_cnt}</div></div>", unsafe_allow_html=True)
        with c3: st.markdown(f"<div class='kpi'><div class='label'>TA Include / Unsure</div><div class='value'>{includes} / {unsure}</div></div>", unsafe_allow_html=True)
        with c4: st.markdown(f"<div class='kpi'><div class='label'>LLM</div><div class='value'>{'ON' if llm_available() else 'OFF'}</div></div>", unsafe_allow_html=True)

        if df is None or df.empty:
            st.info("尚無 records。")
        else:
            render_prisma_html(pr)

    with tabs[1]:
        st.markdown("### Step 1：PubMed 搜尋式（可直接複製）")
        st.code(st.session_state.get("pubmed_query",""), language="text")
        st.markdown(f"- 文章類型 filter：**{st.session_state.get('article_type','不限')}**")
        if st.session_state.get("custom_pubmed_filter","").strip():
            st.markdown(f"- 自訂 filter：`{st.session_state.get('custom_pubmed_filter','')}`")

    with tabs[2]:
        st.markdown("### Step 2：可行性掃描（既有 SR/MA/NMA）")
        st.code(st.session_state.get("feas_query",""), language="text")
        feas = (diag.get("feasibility", {}) or {})
        st.markdown(f"- SR/MA/NMA count：**{feas.get('count','')}**")
        if df_feas is not None and not df_feas.empty:
            st.dataframe(df_feas[["record_id","year","first_author","journal","title","doi"]], use_container_width=True, height=320)
        else:
            st.info("未抓到 SR/MA/NMA 命中（可能題目很窄、或 PubMed 回應受阻）。")

    # -------------------- Combined Records + Screening --------------------
    with tabs[3]:
        st.markdown("### Step 3+4：Records + 粗篩（AI 保留 + 人工修正）")
        if df is None or df.empty:
            st.warning("沒有抓到 records。建議把研究問題寫更具體（縮寫寫全名、加型號、族群、outcome）。")
        else:
            ensure_columns(df, ["record_id","pmid","year","doi","journal","first_author","title","abstract","source"], "")

            st.caption("建議閱讀方式：先看 AI 建議與理由 → 再看 abstract → 需要就 override。")
            for _, r in df.iterrows():
                rid = r["record_id"]
                pmid = r.get("pmid","")
                doi = r.get("doi","")
                year = r.get("year","")
                fa = r.get("first_author","") or "—"
                journal = r.get("journal","") or "—"
                title = r.get("title","")
                abstract = r.get("abstract","")
                source = r.get("source","PubMed")

                ai_d = st.session_state["ta_ai"].get(rid, "Unsure")
                ai_r = st.session_state["ta_ai_reason"].get(rid, "")
                ai_c = float(st.session_state["ta_ai_conf"].get(rid, 0.5) or 0.5)
                eff = compute_effective_decision(rid)

                badge = "badge-warn"
                if eff == "Include": badge = "badge-ok"
                elif eff == "Exclude": badge = "badge-bad"

                with st.expander(f"{rid}｜{short(title, 110)}", expanded=False):
                    st.markdown(
                        f"""
<div class="card">
  <span class="badge {badge}">Effective：{eff}</span>
  <span class="badge">AI：{ai_d}</span>
  <span class="badge">信心度：{ai_c:.2f}</span>
  <br><br>
  <b>PMID:</b> {pmid}　　<b>DOI:</b> {doi or "—"}　　<b>Year:</b> {year or "—"}<br>
  <b>First author:</b> {fa}　　<b>Journal:</b> {journal}　　<b>Source:</b> {source}<br><br>
  <b>Links:</b>
  {f'<a href="{pubmed_link(pmid)}" target="_blank">PubMed/Link</a>' if pmid else '—'}
  &nbsp;|&nbsp;
  {f'<a href="{doi_link(doi)}" target="_blank">DOI</a>' if doi else '—'}
  <br><br>
  <b>AI Title/Abstract 建議</b><br>
  理由：{ai_r or "（無）"}
</div>
                        """,
                        unsafe_allow_html=True
                    )
                    st.markdown("**Abstract（分段顯示）**")
                    st.write(format_abstract(abstract) if abstract else "（無 abstract）")

                    st.markdown("<hr class='hr'/>", unsafe_allow_html=True)
                    st.subheader("人工修正（Override）")
                    new_dec = st.selectbox(
                        "人工 Override（留空＝採用 AI）",
                        options=["", "Include", "Exclude", "Unsure"],
                        index=["", "Include", "Exclude", "Unsure"].index(st.session_state["ta_override"].get(rid, "") if st.session_state["ta_override"].get(rid, "") in ["", "Include", "Exclude", "Unsure"] else ""),
                        key=f"ov_dec_{rid}"
                    )
                    new_reason = st.text_area(
                        "人工理由（建議寫：PICO/研究設計/族群/介入比較/outcome 不符等）",
                        value=st.session_state["ta_override_reason"].get(rid, ""),
                        height=90,
                        key=f"ov_rs_{rid}"
                    )
                    if st.button("儲存這篇修正", key=f"save_{rid}"):
                        st.session_state["ta_override"][rid] = new_dec
                        st.session_state["ta_override_reason"][rid] = new_reason
                        st.success("已儲存。")

    # -------------------- Extraction (form) --------------------
    with tabs[4]:
        st.markdown("### Step 5：Data extraction（寬表）")
        if df is None or df.empty:
            st.info("沒有 records。")
        else:
            include_ids = [rid for rid in df["record_id"].tolist() if compute_effective_decision(rid) == "Include"]
            cands = df[df["record_id"].isin(include_ids)].copy()

            if cands.empty:
                st.warning("目前沒有 Effective=Include 的研究；請先在 Step 3+4 進行 override。")
            else:
                st.caption("為避免每輸入一格就 rerun/跳動：本區改成『按儲存才寫入』；MA 也改為按鈕更新。")

                base = cands[["record_id","pmid","year","doi","journal","first_author","title"]].copy()
                ensure_columns(base, [
                    "Outcome_label","Timepoint",
                    "Effect_measure","Effect","Lower_CI","Upper_CI","Effect_unit",
                    "Notes"
                ], default="")

                prev = st.session_state.get("extract_df", pd.DataFrame())
                if isinstance(prev, pd.DataFrame) and (not prev.empty) and ("record_id" in prev.columns):
                    # merge previous content by record_id + Outcome_label + Timepoint (loose)
                    for c in prev.columns:
                        if c not in base.columns:
                            base[c] = ""
                    # append prev rows instead of 1:1 merge (more flexible)
                    base = pd.concat([base, prev], ignore_index=True)
                    base = base.drop_duplicates(subset=["record_id","Outcome_label","Timepoint","Effect_measure","Effect","Lower_CI","Upper_CI"], keep="first")

                # Quick add row
                with st.expander("快速新增一列（推薦，最不會跳）", expanded=False):
                    with st.form("quick_add"):
                        rid_sel = st.selectbox("record_id", options=include_ids)
                        oc = st.text_input("Outcome_label（必填）")
                        tp = st.text_input("Timepoint（可空）")
                        meas = st.selectbox("Effect_measure（必填）", options=["OR","RR","HR","MD","SMD"])
                        eff = st.text_input("Effect（必填，數字）")
                        lcl = st.text_input("Lower_CI（必填，數字）")
                        ucl = st.text_input("Upper_CI（必填，數字）")
                        unit = st.text_input("Effect_unit（可空）")
                        notes = st.text_area("Notes（可空）", height=80)
                        submitted = st.form_submit_button("新增到寬表")
                        if submitted:
                            new_row = {
                                "record_id": rid_sel,
                                "pmid": "", "year": "", "doi": "", "journal": "", "first_author": "", "title": "",
                                "Outcome_label": oc.strip(),
                                "Timepoint": tp.strip(),
                                "Effect_measure": meas.strip(),
                                "Effect": eff.strip(),
                                "Lower_CI": lcl.strip(),
                                "Upper_CI": ucl.strip(),
                                "Effect_unit": unit.strip(),
                                "Notes": notes.strip(),
                            }
                            st.session_state["extract_df"] = pd.concat([st.session_state.get("extract_df", pd.DataFrame()), pd.DataFrame([new_row])], ignore_index=True)
                            st.session_state["extract_saved"] = True
                            st.success("已新增。請在下方寬表檢視/修正。")

                # Main editable table (form commit)
                st.markdown("#### 寬表（按『儲存寬表』才會寫入）")
                with st.form("extract_form"):
                    # Make numeric inputs TEXT to avoid Streamlit numeric flicker; parse later
                    col_cfg = {
                        "record_id": st.column_config.TextColumn("record_id", disabled=True),
                        "title": st.column_config.TextColumn("Title", disabled=True, width="large"),
                        "Effect_measure": st.column_config.SelectboxColumn("Effect measure", options=["", "OR","RR","HR","MD","SMD"]),
                        "Effect": st.column_config.TextColumn("Effect（文字輸入，避免跳）"),
                        "Lower_CI": st.column_config.TextColumn("Lower CI（文字輸入，避免跳）"),
                        "Upper_CI": st.column_config.TextColumn("Upper CI（文字輸入，避免跳）"),
                    }
                    ex = st.data_editor(
                        base,
                        use_container_width=True,
                        hide_index=True,
                        num_rows="dynamic",
                        column_config=col_cfg
                    )
                    save_ok = st.form_submit_button("儲存寬表（commit）")
                    if save_ok:
                        st.session_state["extract_df"] = ex.copy()
                        st.session_state["extract_saved"] = True
                        st.success("已儲存寬表。")

                # Validation (non-blocking)
                ex2 = st.session_state.get("extract_df", ex)
                v = ex2.copy()
                for c in ["Effect","Lower_CI","Upper_CI"]:
                    v[c] = pd.to_numeric(v[c], errors="coerce")

                issues = []
                for _, row in v.iterrows():
                    rid = row.get("record_id","")
                    outcome = str(row.get("Outcome_label","") or "").strip()
                    meas = str(row.get("Effect_measure","") or "").strip().upper()

                    missing = []
                    if not outcome: missing.append("Outcome_label")
                    if not meas: missing.append("Effect_measure")
                    if pd.isna(row.get("Effect")): missing.append("Effect")
                    if pd.isna(row.get("Lower_CI")): missing.append("Lower_CI")
                    if pd.isna(row.get("Upper_CI")): missing.append("Upper_CI")

                    invalid = []
                    if meas in ["OR","RR","HR"]:
                        for k in ["Effect","Lower_CI","Upper_CI"]:
                            val = row.get(k)
                            if pd.notna(val) and float(val) <= 0:
                                invalid.append(f"{k}<=0（{meas} 需 >0 才能 log）")
                    if pd.notna(row.get("Lower_CI")) and pd.notna(row.get("Upper_CI")) and float(row.get("Upper_CI")) <= float(row.get("Lower_CI")):
                        invalid.append("Upper_CI <= Lower_CI")

                    if missing or invalid:
                        issues.append({"record_id": rid, "missing": ", ".join(missing), "invalid": "; ".join(invalid)})

                if issues:
                    st.markdown("<div class='card'><b>資料檢核（不會卡住，但建議修正）</b><br>"
                                "<span class='red'>紅色提示代表缺資料或數值不合法；不修正仍可下一步，但該筆可能不會納入 MA。</span></div>",
                                unsafe_allow_html=True)
                    st.dataframe(pd.DataFrame(issues), use_container_width=True, height=220)
                else:
                    st.success("目前寬表看起來沒有缺資料/明顯不合法。")

                st.download_button("下載 extraction 寬表（CSV）", data=to_csv_bytes(ex2), file_name="extraction_wide.csv", mime="text/csv")

    # -------------------- MA + Forest (button) --------------------
    with tabs[5]:
        st.markdown("### Step 6：MA + 森林圖（按鈕更新）")
        ex = st.session_state.get("extract_df", pd.DataFrame())
        if ex is None or ex.empty:
            st.info("尚未建立 extraction 寬表。")
        else:
            dfm = ex.copy()
            ensure_columns(dfm, ["Outcome_label","Effect_measure","Effect","Lower_CI","Upper_CI","Timepoint","title","year"], "")
            for c in ["Effect","Lower_CI","Upper_CI"]:
                dfm[c] = pd.to_numeric(dfm[c], errors="coerce")
            dfm["Outcome_label"] = dfm["Outcome_label"].astype(str).str.strip()
            dfm["Effect_measure"] = dfm["Effect_measure"].astype(str).str.strip()

            outcomes = sorted([x for x in dfm["Outcome_label"].unique().tolist() if x])
            if not outcomes:
                st.warning("你尚未填入 Outcome_label。")
                st.stop()

            st.caption("這裡改成：選完 outcome/measure → 按『更新 MA/森林圖』才會計算，避免輸入時一直跳。")
            st.session_state["ma_outcome_input"] = st.selectbox("Outcome_label", options=outcomes, index=0)
            sub0 = dfm[dfm["Outcome_label"] == st.session_state["ma_outcome_input"]].copy()
            measures = sorted([m for m in sub0["Effect_measure"].unique().tolist() if m])
            if not measures:
                st.warning("該 outcome 尚未填 Effect_measure。")
                st.stop()

            st.session_state["ma_measure_choice"] = st.selectbox("Effect measure", options=measures, index=0)
            st.session_state["ma_model_choice"] = st.radio("模型", options=["Fixed effect", "Random effects (DL)"], horizontal=True)

            if st.button("更新 MA / 森林圖"):
                sub = sub0[sub0["Effect_measure"] == st.session_state["ma_measure_choice"]].copy()

                studies, eff_os, lcl_os, ucl_os = [], [], [], []
                effects_t, ses, weights = [], [], []
                skipped = []

                for _, r in sub.iterrows():
                    label = f"{short(str(r.get('title','') or ''), 60)} ({str(r.get('year','') or '')})"
                    try:
                        eff = float(r["Effect"]); lcl = float(r["Lower_CI"]); ucl = float(r["Upper_CI"])
                    except Exception:
                        skipped.append({"study": label, "reason": "Effect/CI 非數字或缺失"})
                        continue
                    se, err = se_from_ci_safe(eff, lcl, ucl, st.session_state["ma_measure_choice"])
                    if err:
                        skipped.append({"study": label, "effect": eff, "lcl": lcl, "ucl": ucl, "reason": err})
                        continue

                    studies.append(label)
                    eff_os.append(eff); lcl_os.append(lcl); ucl_os.append(ucl)
                    effects_t.append(transform_effect(eff, st.session_state["ma_measure_choice"]))
                    ses.append(se)

                if skipped:
                    st.session_state["ma_skipped_rows"] = pd.DataFrame(skipped)
                else:
                    st.session_state["ma_skipped_rows"] = pd.DataFrame()

                if len(studies) < 2:
                    st.error("可用研究數 < 2（扣除不合法列後）。請修正 CI/measure 或補齊更多研究。")
                    st.session_state["ma_last_result"] = None
                else:
                    if st.session_state["ma_model_choice"].startswith("Fixed"):
                        theta, se = pool_fixed(effects_t, ses)
                        lth, uth = theta - 1.96*se, theta + 1.96*se
                        weights = [1.0/(s*s) for s in ses]
                        model_label = "Fixed"
                        I2 = 0.0; Q = 0.0; tau2 = 0.0
                    else:
                        theta, se, Q, I2, tau2 = pool_random_DL(effects_t, ses)
                        lth, uth = theta - 1.96*se, theta + 1.96*se
                        # RE weights for display
                        weights = [1.0/(ses[i]*ses[i] + tau2) for i in range(len(ses))]
                        model_label = "Random (DL)"

                    pe = inverse_transform(theta, st.session_state["ma_measure_choice"])
                    pl = inverse_transform(lth, st.session_state["ma_measure_choice"])
                    pu = inverse_transform(uth, st.session_state["ma_measure_choice"])

                    st.session_state["ma_last_result"] = {
                        "k": len(studies),
                        "model": model_label,
                        "measure": st.session_state["ma_measure_choice"],
                        "pooled": (pe, pl, pu),
                        "I2": I2, "Q": Q, "tau2": tau2,
                        "studies": studies,
                        "eff": eff_os, "lcl": lcl_os, "ucl": ucl_os,
                        "weights": weights
                    }

            # Render cached result
            res = st.session_state.get("ma_last_result")
            if isinstance(res, dict):
                c1, c2, c3, c4 = st.columns(4)
                with c1: st.metric("Studies (k)", res["k"])
                with c2: st.metric("I² (%)", f"{res['I2']:.1f}")
                with c3: st.metric("Q", f"{res['Q']:.2f}")
                with c4: st.metric("tau²", f"{res['tau2']:.4f}")

                pe, pl, pu = res["pooled"]
                st.markdown(f"**Pooled effect ({res['model']}, {res['measure']})**：`{pe:.4f}`（95% CI `{pl:.4f}`–`{pu:.4f}`）")

                skipped_df = st.session_state.get("ma_skipped_rows", pd.DataFrame())
                if skipped_df is not None and not skipped_df.empty:
                    st.markdown("<div class='card'><span class='badge badge-warn'>已跳過不合法列</span> "
                                "<span class='red'>避免整段消失：不合法列不納入 MA，請回 Step 5 修正。</span></div>", unsafe_allow_html=True)
                    st.dataframe(skipped_df, use_container_width=True, height=220)

                if HAS_PLOTLY:
                    fig = forest_plot_revman(res["studies"], res["eff"], res["lcl"], res["ucl"], res["weights"], res["pooled"], res["measure"], res["model"])
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("環境缺少 Plotly：請在 requirements.txt 加入 plotly（見下方指引）。目前先用表格呈現。")
                    st.dataframe(pd.DataFrame({
                        "study": res["studies"],
                        "effect": res["eff"],
                        "lcl": res["lcl"],
                        "ucl": res["ucl"],
                        "weight": res["weights"]
                    }), use_container_width=True)

                # ---- Auto MA writing (always visible) ----
                st.markdown("#### 自動輸出 MA Results 段落（一定顯示；LLM 可選）")
                results_template = (
                    f"【Results（統合分析）】\n"
                    f"本次統合分析納入 {res['k']} 篇研究，採用 {res['model']} 模型，效應量指標為 {res['measure']}。\n"
                    f"合併效應為 {pe:.4f}（95% CI {pl:.4f}–{pu:.4f}）。\n"
                    f"異質性：I² = {res['I2']:.1f}%（Q = {res['Q']:.2f}）。\n"
                    f"『請補：各研究的關鍵差異、臨床解讀、敏感度分析/亞組分析（若有）』\n"
                )
                st.text_area("Results 段落（可直接複製到稿件）", value=results_template, height=160)

                if llm_available():
                    if st.button("用 LLM 生成更完整的 Results/Discussion（缺失用『』）"):
                        with st.spinner("LLM 生成中…"):
                            sys = (
                                "你是資深系統性回顧寫作助理。請用繁體中文撰寫 Results 與 Discussion（僅針對統合分析結果），"
                                "不得捏造數據或引用。需要補充的地方用『』占位。"
                            )
                            payload = {"ma_result": res, "note": "請保守，避免過度推論。"}
                            try:
                                content = call_openai_compatible(
                                    [{"role":"system","content":sys},{"role":"user","content":json.dumps(payload, ensure_ascii=False)}],
                                    max_tokens=1200
                                )
                                st.text_area("LLM 輸出（請人工核對）", value=content, height=320)
                            except Exception as e:
                                st.error(f"LLM 呼叫失敗：{e}")
                else:
                    st.caption("未啟用 BYOK：仍可用上方模板；需要更完整段落再開啟 BYOK。")
            else:
                st.info("請先按『更新 MA / 森林圖』。")

    with tabs[6]:
        st.markdown("### Diagnostics")
        st.code(json.dumps(diag, ensure_ascii=False, indent=2), language="json")
 # -------------------- Step 6 MA + Forest --------------------
    with tabs[6]:
        st.markdown("### Step 6：MA + 森林圖（RevMan-like）")
        ex = st.session_state.get("extract_df", pd.DataFrame())
        if ex is None or ex.empty:
            st.info("尚未建立 extraction 寬表。")
        else:
            dfm = ex.copy()
            ensure_columns(dfm, ["Outcome_label","Effect_measure","Effect","Lower_CI","Upper_CI","Timepoint"], "")
            for c in ["Effect","Lower_CI","Upper_CI"]:
                dfm[c] = pd.to_numeric(dfm[c], errors="coerce")

            dfm["Outcome_label"] = dfm["Outcome_label"].astype(str).str.strip()
            dfm["Effect_measure"] = dfm["Effect_measure"].astype(str).str.strip()

            # outcome input
            available_outcomes = sorted([x for x in dfm["Outcome_label"].unique().tolist() if x])
            if not available_outcomes:
                st.warning("你尚未在寬表填入 Outcome_label。仍可先繼續，但 MA 需要至少一個 outcome 命名一致。")
                available_outcomes = ["(未命名 outcome)"]
                dfm.loc[dfm["Outcome_label"] == "", "Outcome_label"] = "(未命名 outcome)"

            st.caption("Outcome_label 請在寬表統一命名；下方手動輸入用於選取。")
            default_outcome = st.session_state.get("ma_outcome_input") or available_outcomes[0]
            chosen_outcome = st.text_input("Outcome_label（手動輸入/可貼上）", value=default_outcome, key="ma_outcome_input").strip()
            if not chosen_outcome:
                chosen_outcome = available_outcomes[0]

            sub = dfm[dfm["Outcome_label"] == chosen_outcome].copy()
            if sub.empty:
                st.warning("找不到你輸入的 Outcome_label 對應列。請確認拼字（含空白/大小寫）。")
                st.stop()

            measures = sorted([m for m in sub["Effect_measure"].unique().tolist() if m])
            if not measures:
                st.warning("該 outcome 尚未填 Effect_measure。")
                st.stop()

            prev_meas = st.session_state.get("ma_measure_choice") or measures[0]
            if prev_meas not in measures:
                prev_meas = measures[0]
            chosen_measure = st.selectbox("選擇 effect measure", options=measures, index=measures.index(prev_meas), key="ma_measure_choice")
            sub = sub[sub["Effect_measure"] == chosen_measure].copy()

            # validate and build list
            studies, eff_os, lcl_os, ucl_os, effects_t, ses = [], [], [], [], [], []
            skipped = []

            for _, r in sub.iterrows():
                title = str(r.get("title","") or "")
                year = str(r.get("year","") or "")
                label = f"{short(title, 60)} ({year})"

                try:
                    eff = float(r["Effect"])
                    lcl = float(r["Lower_CI"])
                    ucl = float(r["Upper_CI"])
                except Exception:
                    skipped.append({"study": label, "reason": "Effect/CI 非數字或缺失"})
                    continue

                se, err = se_from_ci_safe(eff, lcl, ucl, chosen_measure)
                if err:
                    skipped.append({"study": label, "effect": eff, "lcl": lcl, "ucl": ucl, "reason": err})
                    continue

                studies.append(label)
                eff_os.append(eff); lcl_os.append(lcl); ucl_os.append(ucl)
                effects_t.append(transform_effect(eff, chosen_measure))
                ses.append(se)

            if skipped:
                st.markdown("<div class='card'><span class='badge badge-warn'>已跳過不合法列</span> "
                            "<span class='red'>這些列不會納入 MA（避免整段消失）。</span></div>", unsafe_allow_html=True)
                st.dataframe(pd.DataFrame(skipped), use_container_width=True, height=220)

            if len(studies) < 2:
                st.error("可用研究數 < 2（扣除不合法列後）。請修正 CI/measure 或補齊更多研究。")
                st.stop()

            res = pool_fixed_random(effects_t, ses)
            model = st.radio("模型", options=["Fixed effect", "Random effects (DL)"], horizontal=True, key="ma_model_choice")

            if model.startswith("Fixed"):
                theta = res["fixed"]["theta"]
                lth, uth = res["fixed"]["lcl"], res["fixed"]["ucl"]
                w = res["w_fixed"]
                model_label = "Fixed"
            else:
                theta = res["random"]["theta"]
                lth, uth = res["random"]["lcl"], res["random"]["ucl"]
                # approximate RE weights for display
                tau2 = res["random"]["tau2"]
                w = [1.0/(se*se + tau2) for se in ses]
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

            # RevMan-like forest plot
            if HAS_PLOTLY:
                fig = forest_revman_plotly(studies, eff_os, lcl_os, ucl_os, w, (pe, pl, pu), chosen_measure, model_label)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("環境缺少 Plotly：改以表格顯示森林圖資料。")
                st.dataframe(pd.DataFrame({"study": studies, "effect": eff_os, "lcl": lcl_os, "ucl": ucl_os, "weight": w}), use_container_width=True)

    # -------------------- Step 6b ROB2 --------------------
    with tabs[7]:
        st.markdown("### Step 6b：ROB 2.0（需理由；可人工修正）")
        if df is None or df.empty:
            st.info("沒有 records。")
        else:
            include_ids = [rid for rid in df["record_id"].tolist() if compute_effective_decision(rid) == "Include"]
            cands = df[df["record_id"].isin(include_ids)].copy()

            if cands.empty:
                st.warning("目前沒有 Effective=Include 的研究；請先在 Step 4 進行 override。")
            else:
                st.caption("ROB 2.0 建議在納入後做。此處要求：每個 domain + overall 都要填等級與理由。")

                for _, r in cands.iterrows():
                    rid = r["record_id"]
                    title = r.get("title","")
                    rob = st.session_state["rob2"].get(rid) or rob2_default()
                    st.session_state["rob2"][rid] = rob  # ensure exists

                    with st.expander(f"{rid}｜{short(title, 110)}", expanded=False):
                        st.markdown("<div class='card'><b>ROB 2.0 評估</b><br><span class='small'>每個 domain 與 overall 都需填理由（可引用文中資訊；若不確定請寫明『不確定』並列需查核點）。</span></div>", unsafe_allow_html=True)

                        for key, label in ROB_DOMAINS:
                            col1, col2 = st.columns([0.3, 0.7])
                            with col1:
                                rob[key] = st.selectbox(label, options=ROB_LEVELS, index=ROB_LEVELS.index(rob.get(key,"NA") if rob.get(key,"NA") in ROB_LEVELS else "NA"), key=f"rob_{rid}_{key}")
                            with col2:
                                rob[f"{key}_reason"] = st.text_area("理由", value=rob.get(f"{key}_reason",""), height=85, key=f"rob_{rid}_{key}_rs")

                        st.markdown("<hr class='hr'/>", unsafe_allow_html=True)
                        col1, col2 = st.columns([0.3, 0.7])
                        with col1:
                            rob["overall"] = st.selectbox("Overall ROB", options=ROB_LEVELS, index=ROB_LEVELS.index(rob.get("overall","NA") if rob.get("overall","NA") in ROB_LEVELS else "NA"), key=f"rob_{rid}_overall")
                        with col2:
                            rob["overall_reason"] = st.text_area("Overall 理由", value=rob.get("overall_reason",""), height=90, key=f"rob_{rid}_overall_rs")

                        ok, missing = rob2_is_complete(rob)
                        if ok:
                            st.success("ROB 2.0 已填完整（含理由）。")
                        else:
                            st.markdown(f"<span class='red'>尚未完整：</span>{'；'.join(missing)}", unsafe_allow_html=True)

                # Export ROB2 table
                export_rows = []
                for rid in include_ids:
                    rob = st.session_state["rob2"].get(rid) or rob2_default()
                    export_rows.append({"record_id": rid, **rob})
                rob_df = pd.DataFrame(export_rows)
                st.download_button("下載 ROB 2.0（CSV）", data=to_csv_bytes(rob_df), file_name="rob2.csv", mime="text/csv")

    # -------------------- Step 7 Manuscript draft --------------------
    with tabs[8]:
        st.markdown("### Step 7：稿件草稿（分段呈現；可 BYOK 生成）")
        st.caption("未能自動推論或需要你補的地方會用『』占位；請務必人工核對與改寫。")

        # Minimal draft template (always available)
        pr = compute_prisma(df) if df is not None and not df.empty else {}
        n_id = pr.get("records_identified", "—")
        n_inc = pr.get("studies_included", "—")
        n_meta = pr.get("included_meta", "—")

        default_intro = (
            "【Introduction】\n"
            "『背景：此領域臨床上重要的未解問題與現有證據不足之處』\n"
            f"本研究旨在比較『{st.session_state.get('question','')}』相關介入之臨床結果，並以系統性回顧與統合分析整合現有證據。\n"
        )

        default_methods = (
            "【Methods】\n"
            "本研究遵循 PRISMA 流程。\n"
            f"檢索來源：PubMed（搜尋式見附錄/Step 1）。\n"
            "納入標準：『族群/研究設計/介入與比較/outcome/追蹤時間』\n"
            "排除標準：動物/體外/病例報告等（並依題目調整）。\n"
            "篩選流程：Title/Abstract 粗篩後進行全文評讀；分歧以討論解決。\n"
            "資料萃取：以寬表蒐集 effect 與 95% CI（必要時由原文推算）。\n"
            "偏倚風險：採 ROB 2.0，並記錄各 domain 與 overall 的理由。\n"
            "統計方法：以固定效應或隨機效應模型進行統合分析；異質性以 I² 評估。\n"
        )

        default_results = (
            "【Results】\n"
            f"共檢索到 {n_id} 筆紀錄，最終納入 {n_inc} 篇研究；其中 {n_meta} 篇具備可用數據納入統合分析（詳 PRISMA）。\n"
            "主要結局：『Outcome_label、效應量、95% CI、I²』\n"
            "次要結局：『……』\n"
        )

        default_disc = (
            "【Discussion】\n"
            "本研究整合現有證據，顯示『主要發現與臨床意義』。\n"
            "可能機轉：『……』\n"
            "限制：研究數量、異質性、測量差異、偏倚風險、出版偏倚等。\n"
            "未來研究：建議更多高品質 RCT/一致 outcome 報告。\n"
        )

        default_other = (
            "【結論】\n"
            "『一句話總結主要結論與臨床含意；避免過度推論。』\n\n"
            "【關鍵字】\n"
            "『3–6 個 keywords』\n\n"
            "【附錄：PubMed 搜尋式】\n"
            f"{st.session_state.get('pubmed_query','')}\n"
        )

        # Show + optional AI generation
        st.markdown("#### 手動模板（一定有）")
        intro = st.text_area("Introduction", value=st.session_state["ms_sections"].get("intro", default_intro), height=170)
        methods = st.text_area("Methods", value=st.session_state["ms_sections"].get("methods", default_methods), height=220)
        results = st.text_area("Results", value=st.session_state["ms_sections"].get("results", default_results), height=170)
        discussion = st.text_area("Discussion", value=st.session_state["ms_sections"].get("discussion", default_disc), height=190)
        other = st.text_area("Conclusion/Keywords/Appendix", value=st.session_state["ms_sections"].get("other", default_other), height=210)

        if st.button("儲存草稿"):
            st.session_state["ms_sections"]["intro"] = intro
            st.session_state["ms_sections"]["methods"] = methods
            st.session_state["ms_sections"]["results"] = results
            st.session_state["ms_sections"]["discussion"] = discussion
            st.session_state["ms_sections"]["other"] = other
            st.success("已儲存草稿。")

        st.markdown("<hr class='hr'/>", unsafe_allow_html=True)
        st.markdown("#### AI 生成（可選；需 BYOK）")
        if not llm_available():
            st.info("未啟用 BYOK：此區塊自動降級（不會卡住）。")
        else:
            if st.button("用 LLM 生成/改寫草稿（保持學術口吻，缺失用『』）"):
                with st.spinner("LLM 生成中…"):
                    sys = (
                        "你是資深眼科/臨床研究寫作助理。請用繁體中文撰寫一篇系統性回顧與統合分析的稿件草稿，"
                        "分段輸出 Introduction/Methods/Results/Discussion/Conclusion/Keywords/Appendix。"
                        "若資料不足或需要作者補充，請用全形括號『』留下待填欄位，不得捏造數據或引用。"
                    )
                    payload = {
                        "question": st.session_state.get("question",""),
                        "pubmed_query": st.session_state.get("pubmed_query",""),
                        "prisma": compute_prisma(df) if df is not None else {},
                        "notes": "請保守，避免過度推論；引用請用『待補引用』標示。"
                    }
                    try:
                        content = call_openai_compatible(
                            [{"role":"system","content":sys},{"role":"user","content":json.dumps(payload, ensure_ascii=False)}],
                            max_tokens=1900
                        )
                        st.text_area("LLM 輸出（請務必人工核對）", value=content, height=420)
                    except Exception as e:
                        st.error(f"LLM 呼叫失敗：{e}")

    # -------------------- Diagnostics --------------------
    with tabs[9]:
        st.markdown("### Diagnostics")
        st.code(json.dumps(diag, ensure_ascii=False, indent=2), language="json")
