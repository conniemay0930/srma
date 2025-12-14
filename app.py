# app.py
# =========================================================
# 一句話帶你完成 MA（BYOK）
# Author: Ya Hsin Yao
#
# 免責聲明：本工具僅供學術研究/教學用途，不構成醫療建議或法律意見；
# 使用者須自行驗證所有結果、引用與全文內容；請勿上傳可識別之病人資訊。
#
# 校內資源/授權提醒（重要）：
# - 若文章來自校內訂閱（付費期刊/EZproxy/館藏系統），請勿將受版權保護之全文
#   上傳至任何第三方服務或公開部署之網站（包含本 app 的雲端部署）。
# - 請遵守圖書館授權條款：避免大量下載/自動化批次擷取、避免共享全文給未授權者。
# - 若不確定是否可上傳：建議改用「本機版」或僅上傳你有權分享的開放取用全文（OA）。
#
# Privacy notice (BYOK):
# - Key only used for this session; do not use on untrusted deployments;
# - do not upload identifiable patient info.
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

# Optional: Plotly for forest plot
try:
    import plotly.graph_objects as go
    HAS_PLOTLY = True
except Exception:
    HAS_PLOTLY = False

# Optional: graphviz for PRISMA flowchart
try:
    from graphviz import Digraph
    HAS_GRAPHVIZ = True
except Exception:
    HAS_GRAPHVIZ = False

# -------------------- Page config --------------------
st.set_page_config(page_title="一句話帶你完成 MA（繁體中文）", layout="wide")

# -------------------- Styles --------------------
CSS = """
<style>
.small { font-size: 0.88rem; color: #555; }
.muted { color: #6b7280; }
.kpi { border: 1px solid #e5e7eb; border-radius: 14px; padding: 0.75rem 0.9rem; background: #fafafa; }
.kpi .label { font-size: 0.82rem; color: #6b7280; }
.kpi .value { font-size: 1.25rem; font-weight: 800; color: #111827; }
.notice { border-left: 4px solid #f59e0b; background: #fff7ed; padding: 0.85rem 1rem; border-radius: 12px; }
.card { border: 1px solid #dde2eb; border-radius: 14px; padding: 0.9rem 1rem; background:#fff; margin-bottom: 0.9rem; }
.badge { display:inline-block; padding:0.15rem 0.55rem; border-radius: 999px; font-size:0.78rem; margin-right:0.35rem; border:1px solid rgba(0,0,0,0.06); }
.badge-ok { background:#d1fae5; color:#065f46; }
.badge-warn { background:#fef3c7; color:#92400e; }
.badge-bad { background:#fee2e2; color:#991b1b; }
hr.soft { border: none; border-top: 1px solid #eef2f7; margin: 0.8rem 0; }
pre code { font-size: 0.86rem !important; }
a { text-decoration: none; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

# -------------------- Header --------------------
st.title("一句話帶你完成 MA")
st.caption("作者：Ya Hsin Yao　|　Language：繁體中文　|　免責聲明：僅供學術用途；請自行驗證所有結果與引用。")

st.markdown(
    "<div class='notice'>"
    "<b>重要提醒（請務必閱讀）</b><br>"
    "1) 本工具輸出（含引用/數值/結論）可能不完整或不正確，<b>必須由研究者逐一核對原文</b>。<br>"
    "2) <b>請勿上傳可識別病人資訊</b>（姓名、病歷號、影像、日期等）。<br>"
    "3) <b>校內訂閱全文/館藏資源</b>可能受授權限制：避免將受版權保護的全文上傳到雲端服務或公開部署環境；"
    "避免大量下載/自動化批次擷取；遵守圖書館授權條款。<br>"
    "4) 想提升檢索召回：研究問題請盡量包含『族群/情境 + 介入 + 比較 +（主要 outcome）』，縮寫請寫全名或具體型號。<br>"
    "</div>",
    unsafe_allow_html=True
)

with st.expander("目標要求（給學長看的版本）", expanded=False):
    st.markdown(
        """
- 只輸入一句問題 → 自動產出：PICO/criteria、PubMed 搜尋式（MeSH+free text）、抓文獻、AI Title/Abstract 粗篩（可人工修正）、可行性掃描（SR/MA/NMA）、寬表萃取模板、MA + 森林圖、稿件分段草稿。
- 降低人工輸入：預設只需要一句研究問題；必要時在展開區塊微調 PICO/criteria。
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

    # pipeline artifacts
    ss.setdefault("protocol", Protocol(P_syn=[], I_syn=[], C_syn=[], O_syn=[], mesh_P=[], mesh_I=[], mesh_C=[], mesh_O=[]))
    ss.setdefault("pubmed_query", "")
    ss.setdefault("feas_query", "")
    ss.setdefault("pubmed_records", pd.DataFrame())
    ss.setdefault("srma_hits", pd.DataFrame())
    ss.setdefault("diagnostics", {})

    # TA screening
    ss.setdefault("ta_ai", {})               # record_id -> Include/Exclude/Unsure
    ss.setdefault("ta_ai_reason", {})        # record_id -> reason
    ss.setdefault("ta_ai_conf", {})          # record_id -> float 0-1
    ss.setdefault("ta_override", {})         # record_id -> Include/Exclude/Unsure/""
    ss.setdefault("ta_override_reason", {})  # record_id -> text

    # extraction / MA
    ss.setdefault("extract_df", pd.DataFrame())
    ss.setdefault("ma_outcome_input", "")
    ss.setdefault("ma_measure_choice", "")
    ss.setdefault("ma_model_choice", "Fixed effect")

    # manuscript
    ss.setdefault("ms_sections", {})
    ss.setdefault("ms_full_md", "")

    # prisma
    ss.setdefault("last_prisma", {})

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
    st.checkbox("啟用 LLM（BYOK）", key="byok_enabled", help="預設關閉；關閉時流程自動降級，不會卡在 AI 萃取/ROB2。")

    st.markdown(
        "<div class='small muted'>"
        "Key only used for this session（不寫入 secrets、不落盤）。<br>"
        "do not upload identifiable patient info.<br>"
        "校內訂閱全文請避免上傳到雲端部署環境。"
        "</div>",
        unsafe_allow_html=True
    )

    st.session_state["byok_consent"] = st.checkbox(
        "我理解並同意：僅供學術用途；輸出需人工核對；不輸入病人資訊；不違反校內授權。",
        value=bool(st.session_state.get("byok_consent", False))
    )

    st.text_input("Base URL（OpenAI-compatible）", key="byok_base_url")
    st.text_input("Model", key="byok_model")
    st.text_input("API Key（只在本次 session）", type="password", key="byok_key")
    st.slider("Temperature", 0.0, 1.0, 0.2, 0.05, key="byok_temp")
    st.button("Clear key", on_click=lambda: st.session_state.update({"byok_key": ""}))

    st.markdown("---")
    st.subheader("顯示選項")
    st.checkbox("Records：表格顯示 abstract（較占空間）", value=True, key="show_abs_in_table")
    st.checkbox("Records：顯示逐篇卡片（推薦）", value=True, key="show_record_cards")

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
        toks = re.findall(r"[A-Za-z]{2,15}", p)
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

    # 最小假設：I=left, C=right（若無 vs，C 留空）
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

    if P_block:
        blocks.append(P_block)
    if I_block:
        blocks.append(I_block)
    if C_block:
        blocks.append(C_block)
    if O_block:
        blocks.append(O_block)

    core = " AND ".join(blocks) if blocks else quote_tiab(proto.I or proto.P or "systematic review")

    not_block = (proto.NOT or "").strip()
    q = f"({core}) NOT ({not_block})" if not_block else core

    atf = ARTICLE_TYPE_FILTERS.get(article_type, "") or ""
    atf = atf.strip()
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

def pubmed_efetch_xml(pmids: List[str]) -> Tuple[str, List[str]]:
    if not pmids:
        return "", []
    chunks, urls = [], []
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
    解析 PubMed XML：
    - pmid, year, title, abstract, doi, journal, first_author
    """
    xml_text = xml_text or ""
    if "<PubmedArticle" not in xml_text:
        return pd.DataFrame()

    articles = re.split(r"<PubmedArticle\b", xml_text)[1:]
    rows = []
    for a in articles:
        chunk = "<PubmedArticle" + a

        pmid_m = re.search(r"<PMID[^>]*>(\d+)</PMID>", chunk)
        pmid = pmid_m.group(1) if pmid_m else ""

        title_m = re.search(r"<ArticleTitle>(.*?)</ArticleTitle>", chunk, flags=re.DOTALL)
        title = norm_text(re.sub(r"<.*?>", "", title_m.group(1))) if title_m else ""

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

        journal = ""
        j = re.search(r"<Journal>.*?<Title>(.*?)</Title>.*?</Journal>", chunk, flags=re.DOTALL)
        if j:
            journal = norm_text(re.sub(r"<.*?>", "", j.group(1)))

        first_author = ""
        # 取第一個 Author 的 LastName + Initials
        am = re.search(r"<AuthorList>.*?<Author\b.*?>.*?<LastName>(.*?)</LastName>.*?(?:<Initials>(.*?)</Initials>)?.*?</Author>", chunk, flags=re.DOTALL)
        if am:
            ln = norm_text(re.sub(r"<.*?>", "", am.group(1)))
            ini = norm_text(re.sub(r"<.*?>", "", am.group(2))) if am.group(2) else ""
            first_author = (ln + (" " + ini if ini else "")).strip()

        rows.append({
            "pmid": pmid,
            "year": year,
            "title": title,
            "abstract": abstract,
            "doi": doi,
            "journal": journal,
            "first_author": first_author,
        })

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["pmid"] = df["pmid"].astype(str)
    df["record_id"] = df["pmid"].apply(lambda x: f"PMID:{x}")
    df["source"] = "PubMed"
    return df

# =========================================================
# Title/Abstract screening (LLM optional)
# =========================================================
def heuristic_screen(title: str, abstract: str, proto: Protocol) -> Tuple[str, str, float]:
    """
    沒有 LLM 時也要給：
    - decision: Include/Exclude/Unsure
    - reason: 更清楚的敘述
    - confidence: 0-1
    """
    title = title or ""
    abstract = abstract or ""
    blob = (title + " " + abstract).lower()

    # 強排除：動物、體外、病例
    if re.search(r"\b(mice|mouse|rat|rabbit|porcine|canine|in vitro)\b", blob):
        return "Exclude", "偵測到動物/體外相關字樣，較可能非納入範圍（需人工確認）。", 0.85
    if re.search(r"\b(case report|case series)\b", blob):
        return "Exclude", "偵測到病例報告/病例系列字樣，通常不符合 MA 納入（需人工確認）。", 0.80

    # 介入/比較關鍵詞命中
    i_terms = proto.I_syn or expand_terms(proto.I)
    c_terms = proto.C_syn or expand_terms(proto.C)
    hits_i = [t for t in i_terms[:40] if t and t.lower() in blob]
    hits_c = [t for t in c_terms[:40] if t and t.lower() in blob] if c_terms else []

    # trial-like
    trial_like = bool(re.search(r"\b(randomized|randomised|randomly|trial|controlled)\b", blob))

    if hits_i and (trial_like or hits_c):
        reason = (
            f"疑似臨床試驗/比較研究（偵測到 randomized/trial/controlled 等字樣），"
            f"且命中介入關鍵詞：{', '.join(hits_i[:4])}"
            + (f"；比較關鍵詞：{', '.join(hits_c[:3])}" if hits_c else "")
            + "。建議先保留進入 full-text 評讀。"
        )
        conf = 0.75 if trial_like else 0.65
        return "Include", reason, conf

    if hits_i:
        reason = f"命中介入關鍵詞：{', '.join(hits_i[:4])}；但尚不足以確認研究設計/比較組，建議標記 Unsure 進入人工檢視。"
        return "Unsure", reason, 0.55

    # 太短/縮寫導致召回不足時，先給 Unsure
    if len(blob.strip()) < 80:
        return "Unsure", "摘要資訊過少或僅有短句/縮寫，無法可靠判讀；建議人工檢視。", 0.40

    return "Unsure", "未偵測到足夠的 PICO 關鍵詞或研究設計訊號；建議人工快速掃描標題摘要以免漏掉。", 0.45

def screen_with_llm(records: List[Dict[str, Any]], proto: Protocol) -> Dict[str, Dict[str, Any]]:
    """
    回傳：
    record_id -> {decision, reason, confidence}
    """
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
        "reason 請寫清楚：為何保留/排除（例如：研究設計、介入比較、族群、是否臨床試驗等），避免空泛字眼。\n"
        "confidence 0~1，代表你對 decision 的把握。\n"
        "若資訊不足，請選 Unsure，不得捏造全文內容。"
    )
    user = {"protocol": proto.to_dict(), "records": records[:120]}

    try:
        content = call_openai_compatible(
            [{"role":"system","content":sys},{"role":"user","content":json.dumps(user, ensure_ascii=False)}],
            max_tokens=1800
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
        # fallback for missing
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
# PRISMA flowchart
# =========================================================
def compute_prisma(df: pd.DataFrame) -> Dict[str, Any]:
    """
    以目前狀態計算一個「可用於 demo」的 PRISMA 計數：
    - duplicates: 先設 0（你之後若加去重，可替換）
    - screened: records
    - excluded_ta: effective Exclude
    - fulltext_assessed: effective Include + Unsure（保守留給 full-text）
    - included: effective Include
    """
    if df is None or df.empty:
        return {
            "records_identified": 0,
            "duplicates_removed": 0,
            "records_screened": 0,
            "records_excluded": 0,
            "fulltext_assessed": 0,
            "fulltext_excluded": 0,
            "studies_included": 0,
            "included_meta": 0,
        }

    rids = df["record_id"].tolist()
    eff_decisions = []
    for rid in rids:
        od = (st.session_state["ta_override"].get(rid, "") or "").strip()
        ai = st.session_state["ta_ai"].get(rid, "Unsure")
        eff = od if od else ai
        eff_decisions.append(eff)

    total = len(rids)
    excluded = sum(1 for x in eff_decisions if x == "Exclude")
    included = sum(1 for x in eff_decisions if x == "Include")
    unsure = sum(1 for x in eff_decisions if x == "Unsure")

    fulltext_assessed = included + unsure  # demo 保守：Unsure 也進 fulltext
    return {
        "records_identified": total,
        "duplicates_removed": 0,
        "records_screened": total,
        "records_excluded": excluded,
        "fulltext_assessed": fulltext_assessed,
        "fulltext_excluded": 0,
        "studies_included": included,
        "included_meta": included,  # demo：先同 included（你之後可改成只選進 MA 的）
        "unsure_fulltext": unsure,
    }

def prisma_flowchart(pr: Dict[str, Any]):
    if not HAS_GRAPHVIZ:
        st.info("此環境未安裝 graphviz：改用文字版 PRISMA。")
        st.json(pr)
        return

    dot = Digraph(format="png")
    dot.attr(rankdir="TB", fontsize="10")

    def box(name: str, label: str):
        dot.node(name, label=label, shape="box", style="rounded")

    n_id = pr.get("records_identified", 0)
    n_dup = pr.get("duplicates_removed", 0)
    n_scr = pr.get("records_screened", 0)
    n_exc = pr.get("records_excluded", 0)
    n_ft = pr.get("fulltext_assessed", 0)
    n_ft_exc = pr.get("fulltext_excluded", 0)
    n_inc = pr.get("studies_included", 0)
    n_meta = pr.get("included_meta", 0)
    n_unsure = pr.get("unsure_fulltext", 0)

    box("id", f"Records identified\n(n = {n_id})")
    box("dup", f"Duplicates removed\n(n = {n_dup})")
    box("scr", f"Records screened (Title/Abstract)\n(n = {n_scr})")
    box("exc", f"Records excluded\n(n = {n_exc})")
    box("ft", f"Full-text assessed for eligibility\n(n = {n_ft})\n(Unsure needing full-text: {n_unsure})")
    box("ft_exc", f"Full-text excluded\n(n = {n_ft_exc})")
    box("inc", f"Studies included in qualitative synthesis\n(n = {n_inc})")
    box("meta", f"Studies included in meta-analysis\n(n = {n_meta})")

    dot.edge("id", "dup")
    dot.edge("dup", "scr")
    dot.edge("scr", "exc")
    dot.edge("scr", "ft")
    dot.edge("ft", "ft_exc")
    dot.edge("ft", "inc")
    dot.edge("inc", "meta")

    st.graphviz_chart(dot)

# =========================================================
# Meta-analysis + forest plot (robust)
# =========================================================
RATIO_MEASURES = {"OR", "RR", "HR"}

def se_from_ci_safe(effect: float, lcl: float, ucl: float, measure: str) -> Tuple[Optional[float], Optional[str]]:
    """
    永不 raise；回傳 (se, err)
    """
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

def pool_fixed_random(effects: List[float], ses: List[float], measure: str) -> Dict[str, Any]:
    k = len(effects)
    w = [1.0/(se*se) for se in ses]
    sumw = sum(w)
    theta_fixed = sum(w[i]*effects[i] for i in range(k)) / sumw
    se_fixed = math.sqrt(1.0/sumw)

    Q = sum(w[i] * (effects[i]-theta_fixed)**2 for i in range(k))
    C = sumw - (sum(wi*wi for wi in w) / sumw)
    tau2 = max(0.0, (Q - (k-1)) / C) if (C > 0 and k > 1) else 0.0

    w_re = [1.0/(ses[i]**2 + tau2) for i in range(k)]
    sumw_re = sum(w_re)
    theta_re = sum(w_re[i]*effects[i] for i in range(k)) / sumw_re
    se_re = math.sqrt(1.0/sumw_re)

    I2 = max(0.0, (Q - (k-1)) / Q) * 100.0 if (Q > 0 and k > 1) else 0.0

    def ci(theta, se):
        return theta - 1.96*se, theta + 1.96*se

    lf, uf = ci(theta_fixed, se_fixed)
    lr, ur = ci(theta_re, se_re)

    return {
        "k": k,
        "fixed": {"theta": theta_fixed, "se": se_fixed, "lcl": lf, "ucl": uf},
        "random": {"theta": theta_re, "se": se_re, "lcl": lr, "ucl": ur, "tau2": tau2},
        "heterogeneity": {"Q": Q, "df": k-1, "I2": I2},
        "measure": measure
    }

def forest_plot_plotly(studies: List[str], eff: List[float], lcl: List[float], ucl: List[float],
                       pooled: Tuple[float,float,float], measure: str, model_label: str):
    if not HAS_PLOTLY:
        return None
    y = list(range(len(studies)))[::-1]
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=eff, y=y, mode="markers",
        error_x=dict(
            type="data", symmetric=False,
            array=[ucl[i]-eff[i] for i in range(len(eff))],
            arrayminus=[eff[i]-lcl[i] for i in range(len(eff))],
        ),
        hovertext=studies,
        showlegend=False
    ))

    pe, pl, pu = pooled
    pooled_y = -1
    fig.add_trace(go.Scatter(
        x=[pe], y=[pooled_y], mode="markers",
        marker=dict(symbol="diamond", size=12),
        error_x=dict(type="data", symmetric=False, array=[pu-pe], arrayminus=[pe-pl]),
        hovertext=[f"Pooled {model_label}"],
        showlegend=False
    ))

    fig.update_layout(
        height=360 + 18*len(studies),
        xaxis_title=f"Effect ({measure})",
        yaxis=dict(
            tickmode="array",
            tickvals=y + [pooled_y],
            ticktext=studies[::-1] + [f"Pooled ({model_label})"],
        ),
        margin=dict(l=10, r=10, t=35, b=10),
        showlegend=False,
    )
    fig.add_vline(x=1.0 if (measure or "").upper().strip() in RATIO_MEASURES else 0.0, line_width=1, line_dash="dash")
    return fig

# =========================================================
# UI: inputs
# =========================================================
st.subheader("Research question（輸入一句話）")
colA, colB = st.columns([0.62, 0.38])

with colA:
    st.session_state["question"] = st.text_input(
        "例：『不同種類 EDOF IOL 於白內障術後視覺品質（對比敏感度/眩光）比較』或『FLACS 是否優於傳統 phaco』",
        value=st.session_state.get("question",""),
    )

with colB:
    st.session_state["article_type"] = st.selectbox(
        "文章類型（可選，會影響 PubMed filter）",
        options=list(ARTICLE_TYPE_FILTERS.keys()),
        index=list(ARTICLE_TYPE_FILTERS.keys()).index(st.session_state.get("article_type","不限")) if st.session_state.get("article_type","不限") in ARTICLE_TYPE_FILTERS else 0
    )
    st.session_state["custom_pubmed_filter"] = st.text_input(
        "自訂 PubMed filter（可選）",
        value=st.session_state.get("custom_pubmed_filter",""),
        help="例如：humans[MeSH Terms] OR human[tiab]；或限制年份：2020:3000[pdat]"
    )

run = st.button("Run / 執行（自動跑到 Outputs）", type="primary")

# =========================================================
# Pipeline run
# =========================================================
if run:
    q = norm_text(st.session_state["question"])
    if not q:
        st.error("請先輸入一句研究問題。")
        st.stop()

    with st.spinner("Step 0/4：生成 protocol（最小自動）…"):
        proto = question_to_protocol(q)
        st.session_state["protocol"] = proto

    with st.spinner("Step 1/4：產出 PubMed 搜尋式（MeSH + free text + 類型 filter）…"):
        pub_q = build_pubmed_query(proto, st.session_state["article_type"], st.session_state["custom_pubmed_filter"])
        st.session_state["pubmed_query"] = pub_q

    with st.spinner("Step 2/4：可行性掃描（既有 SR/MA/NMA）…"):
        feas_q = build_feasibility_query(st.session_state["pubmed_query"])
        st.session_state["feas_query"] = feas_q
        cnt_feas, ids_feas, feas_url, feas_diag = pubmed_esearch(feas_q, retmax=20, retstart=0)
        xml_feas, _ = pubmed_efetch_xml(ids_feas[:20])
        df_feas = parse_pubmed_xml_minimal(xml_feas)
        st.session_state["srma_hits"] = df_feas

        st.session_state["diagnostics"] = {
            "feasibility": {"count": cnt_feas, "esearch_url": feas_url, "diag": feas_diag},
        }

    with st.spinner("Step 3/4：抓取 PubMed 文獻…"):
        total, ids, es_url, es_diag = pubmed_esearch(st.session_state["pubmed_query"], retmax=200, retstart=0)
        xml, ef_urls = pubmed_efetch_xml(ids[:200])
        df = parse_pubmed_xml_minimal(xml)
        st.session_state["pubmed_records"] = df

        d = st.session_state.get("diagnostics", {}) or {}
        d.update({
            "pubmed_total_count": total,
            "esearch_url": es_url,
            "efetch_urls": ef_urls,
            "esearch_diag": es_diag,
            "warnings": [] if total > 0 else ["PubMed count=0：請把問題寫更具體（縮寫寫全名/加上型號/族群/outcome），或看 Diagnostics 是否被阻擋。"],
        })
        st.session_state["diagnostics"] = d

    with st.spinner("Step 4/4：Title/Abstract 粗篩（AI 可選 + 可人工修正）…"):
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
        "Step 3 Records（表格/卡片）",
        "Step 4 粗篩（AI+人工修正）",
        "Step 5-6 萃取/MA/森林圖",
        "Diagnostics"
    ])

    # -------------------- Overview / PRISMA --------------------
    with tabs[0]:
        st.markdown("### 流程總覽（PRISMA）")

        total = int(diag.get("pubmed_total_count", 0) or 0)
        feas_cnt = int((diag.get("feasibility", {}) or {}).get("count", 0) or 0)

        includes = 0
        if df is not None and not df.empty:
            for rid in df["record_id"].tolist():
                od = (st.session_state["ta_override"].get(rid, "") or "").strip()
                ai = st.session_state["ta_ai"].get(rid, "Unsure")
                eff = od if od else ai
                if eff == "Include":
                    includes += 1

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f"<div class='kpi'><div class='label'>PubMed count</div><div class='value'>{total}</div></div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"<div class='kpi'><div class='label'>既有 SR/MA/NMA</div><div class='value'>{feas_cnt}</div></div>", unsafe_allow_html=True)
        with c3:
            st.markdown(f"<div class='kpi'><div class='label'>TA Include（有效決策）</div><div class='value'>{includes}</div></div>", unsafe_allow_html=True)
        with c4:
            st.markdown(f"<div class='kpi'><div class='label'>LLM</div><div class='value'>{'ON' if llm_available() else 'OFF'}</div></div>", unsafe_allow_html=True)

        if df is None or df.empty:
            st.info("尚無 records。")
        else:
            pr = compute_prisma(df)
            st.session_state["last_prisma"] = pr
            prisma_flowchart(pr)

    # -------------------- Step 1 Query --------------------
    with tabs[1]:
        st.markdown("### Step 1：PubMed 搜尋式（可直接複製）")
        st.code(st.session_state.get("pubmed_query",""), language="text")
        st.markdown(f"- 文章類型 filter：**{st.session_state.get('article_type','不限')}**")
        if st.session_state.get("custom_pubmed_filter","").strip():
            st.markdown(f"- 自訂 filter：`{st.session_state.get('custom_pubmed_filter','')}`")

    # -------------------- Step 2 Feasibility --------------------
    with tabs[2]:
        st.markdown("### Step 2：可行性掃描（既有 SR/MA/NMA）")
        st.code(st.session_state.get("feas_query",""), language="text")
        feas = (diag.get("feasibility", {}) or {})
        st.markdown(f"- SR/MA/NMA count：**{feas.get('count','')}**")

        if df_feas is not None and not df_feas.empty:
            show = df_feas[["record_id","year","first_author","journal","title","doi"]].copy()
            st.dataframe(show, use_container_width=True, height=320)
        else:
            st.info("目前未抓到 SR/MA/NMA 命中（可能是題目太窄、或 PubMed 回應受阻）。")

    # -------------------- Step 3 Records --------------------
    with tabs[3]:
        st.markdown("### Step 3：Records")
        if df is None or df.empty:
            st.warning("沒有抓到 records。建議把研究問題寫更具體（縮寫寫全名、加型號、族群、outcome）。")
        else:
            ensure_columns(df, ["record_id","pmid","year","doi","journal","first_author","title","abstract","source"], "")

            cols = ["record_id","year","first_author","journal","pmid","doi","title"]
            if st.session_state.get("show_abs_in_table", True):
                cols += ["abstract"]
            st.dataframe(df[cols], use_container_width=True, height=380)

            if st.session_state.get("show_record_cards", True):
                st.markdown("#### 逐篇展開（推薦）")
                for _, r in df.iterrows():
                    rid = r["record_id"]
                    pmid = r.get("pmid","")
                    doi = r.get("doi","")
                    year = r.get("year","")
                    fa = r.get("first_author","")
                    journal = r.get("journal","")
                    title = r.get("title","")
                    abstract = r.get("abstract","")
                    source = r.get("source","PubMed")

                    ai_d = st.session_state["ta_ai"].get(rid, "Unsure")
                    ai_r = st.session_state["ta_ai_reason"].get(rid, "")
                    ai_c = st.session_state["ta_ai_conf"].get(rid, 0.5)

                    badge = "badge-warn"
                    if ai_d == "Include":
                        badge = "badge-ok"
                    elif ai_d == "Exclude":
                        badge = "badge-bad"

                    header = f"{short(title, 110)}"
                    with st.expander(header, expanded=False):
                        st.markdown(
                            f"""
<div class="card">
<span class="badge {badge}">AI 建議：{ai_d}</span>
<span class="badge">信心度：{ai_c:.2f}</span><br><br>

<b>ID:</b> {rid}　　<b>PMID:</b> {pmid}　　<b>DOI:</b> {doi or "—"}<br>
<b>Year:</b> {year or "—"}　　<b>First author:</b> {fa or "—"}　　<b>Journal:</b> {journal or "—"}　　<b>Source:</b> {source}<br><br>

<b>Links:</b>
{f'<a href="{pubmed_link(pmid)}" target="_blank">PubMed</a>' if pmid else '—'}
&nbsp;|&nbsp;
{f'<a href="{doi_link(doi)}" target="_blank">DOI</a>' if doi else '—'}<br><br>

<b>AI 理由（Title/Abstract）</b>：{ai_r or "（無）"}<br>
</div>
                            """,
                            unsafe_allow_html=True
                        )

                        st.markdown("**Abstract（可人工快速掃描）**")
                        st.write(abstract if abstract else "（無 abstract）")

    # -------------------- Step 4 Screening + Override --------------------
    with tabs[4]:
        st.markdown("### Step 4：Title/Abstract 粗篩（AI 保留 + 人工修正）")
        if df is None or df.empty:
            st.info("沒有 records。")
        else:
            st.caption("你可以在這裡把 AI 建議改成 Include/Exclude/Unsure，並寫下人工理由。Effective decision 會用於 PRISMA 與後續 extraction/MA。")

            rows = []
            for _, r in df.iterrows():
                rid = r["record_id"]
                ai_d = st.session_state["ta_ai"].get(rid, "Unsure")
                ai_r = st.session_state["ta_ai_reason"].get(rid, "")
                ai_c = st.session_state["ta_ai_conf"].get(rid, 0.5)

                od = (st.session_state["ta_override"].get(rid, "") or "").strip()
                orr = st.session_state["ta_override_reason"].get(rid, "")
                eff = od if od else ai_d

                rows.append({
                    "record_id": rid,
                    "year": r.get("year",""),
                    "first_author": r.get("first_author",""),
                    "journal": r.get("journal",""),
                    "title": r.get("title",""),
                    "AI_decision": ai_d,
                    "AI_confidence": float(ai_c),
                    "AI_reason": ai_r,
                    "Override_decision": od,
                    "Override_reason": orr,
                    "Effective_decision": eff,
                })

            sdf = pd.DataFrame(rows)
            st.dataframe(sdf[["record_id","year","first_author","journal","title","AI_decision","AI_confidence","Effective_decision"]], use_container_width=True, height=280)

            st.markdown("#### 展開逐篇人工修正（推薦）")
            for _, r in sdf.iterrows():
                rid = r["record_id"]
                title = r["title"]
                with st.expander(f"{rid}｜{short(title, 110)}", expanded=False):
                    st.markdown(f"**AI 建議**：{r['AI_decision']}　|　**信心度**：{r['AI_confidence']:.2f}")
                    st.markdown(f"**AI 理由**：{r['AI_reason'] or '（無）'}")

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

            st.download_button("下載粗篩結果（含 AI + override）", data=to_csv_bytes(sdf), file_name="screening_ta_ai_override.csv", mime="text/csv")

    # -------------------- Step 5-6 Extraction + MA --------------------
    with tabs[5]:
        st.markdown("### Step 5：Extraction 寬表 → Step 6：MA + 森林圖")
        if df is None or df.empty:
            st.info("沒有 records。")
        else:
            # include set
            include_ids = []
            for rid in df["record_id"].tolist():
                od = (st.session_state["ta_override"].get(rid, "") or "").strip()
                ai_d = st.session_state["ta_ai"].get(rid, "Unsure")
                eff = od if od else ai_d
                if eff == "Include":
                    include_ids.append(rid)

            cands = df[df["record_id"].isin(include_ids)].copy()
            if cands.empty:
                st.warning("目前沒有 Effective=Include 的研究；請先在 Step 4 進行 override。")
            else:
                base = cands[["record_id","pmid","year","doi","journal","first_author","title"]].copy()
                ensure_columns(base, [
                    "Outcome_label","Timepoint",
                    "Effect_measure","Effect","Lower_CI","Upper_CI","Effect_unit",
                    "Notes"
                ], default="")

                # merge previous edits (stable UX)
                prev = st.session_state.get("extract_df", pd.DataFrame())
                if isinstance(prev, pd.DataFrame) and (not prev.empty) and ("record_id" in prev.columns):
                    for c in prev.columns:
                        if c not in base.columns:
                            base[c] = ""
                    keep = [c for c in base.columns]
                    prev2 = prev.reindex(columns=keep, fill_value="")
                    base = base.merge(prev2.drop_duplicates(subset=["record_id"]), on="record_id", how="left", suffixes=("","_old"))
                    for c in list(base.columns):
                        if c.endswith("_old"):
                            orig = c[:-4]
                            base[orig] = base.apply(lambda rr: rr[c] if str(rr[c]).strip() not in ["", "nan", "None"] else rr[orig], axis=1)
                            base = base.drop(columns=[c])

                st.caption("提示：OR/RR/HR 的 effect/CI 必須都 > 0；若有 0 或負值會自動跳過並列原因（避免整段消失）。")
                ex = st.data_editor(
                    base,
                    use_container_width=True,
                    hide_index=True,
                    num_rows="dynamic",
                    column_config={
                        "record_id": st.column_config.TextColumn("record_id", disabled=True),
                        "title": st.column_config.TextColumn("Title", disabled=True, width="large"),
                        "Effect_measure": st.column_config.SelectboxColumn("Effect measure", options=["", "OR","RR","HR","MD","SMD"]),
                        "Effect": st.column_config.NumberColumn("Effect", format="%.6f"),
                        "Lower_CI": st.column_config.NumberColumn("Lower CI", format="%.6f"),
                        "Upper_CI": st.column_config.NumberColumn("Upper CI", format="%.6f"),
                    }
                )
                st.session_state["extract_df"] = ex
                st.download_button("下載 extraction 寬表（CSV）", data=to_csv_bytes(ex), file_name="extraction_wide.csv", mime="text/csv")

                st.markdown("<hr class='soft'/>", unsafe_allow_html=True)
                st.markdown("### Step 6：MA + 森林圖")

                dfm = ex.copy()
                ensure_columns(dfm, ["Outcome_label","Effect_measure","Effect","Lower_CI","Upper_CI","Timepoint"], "")
                for c in ["Effect","Lower_CI","Upper_CI"]:
                    dfm[c] = pd.to_numeric(dfm[c], errors="coerce")
                dfm = dfm.dropna(subset=["Effect","Lower_CI","Upper_CI"])
                dfm = dfm[dfm["Effect_measure"].astype(str).str.strip() != ""]

                if dfm.empty:
                    st.info("請至少填：Effect_measure + Effect + Lower_CI + Upper_CI（可加 Outcome_label/Timepoint）才能做 MA。")
                else:
                    available_outcomes = sorted([x for x in dfm["Outcome_label"].astype(str).unique().tolist() if x.strip()])
                    if not available_outcomes:
                        dfm["Outcome_label"] = "(未命名 outcome)"
                        available_outcomes = ["(未命名 outcome)"]

                    st.caption("可用 outcomes（供參考）： " + " | ".join(available_outcomes[:20]) + (" …" if len(available_outcomes) > 20 else ""))

                    default_outcome = st.session_state.get("ma_outcome_input") or available_outcomes[0]
                    chosen_outcome = st.text_input("Outcome_label（手動輸入/可貼上）", value=default_outcome, key="ma_outcome_input").strip()
                    if not chosen_outcome:
                        chosen_outcome = available_outcomes[0]

                    sub = dfm[dfm["Outcome_label"].astype(str).str.strip() == chosen_outcome].copy()
                    if sub.empty:
                        st.warning("你輸入的 Outcome_label 在寬表中找不到對應列。請確認拼字（含大小寫/空白），或先在寬表統一命名。")
                        st.stop()

                    measures = sorted(sub["Effect_measure"].astype(str).unique().tolist())
                    prev_meas = st.session_state.get("ma_measure_choice") or (measures[0] if measures else "")
                    if prev_meas not in measures and measures:
                        prev_meas = measures[0]
                    chosen_measure = st.selectbox("選擇 effect measure", options=measures, index=measures.index(prev_meas) if prev_meas in measures else 0, key="ma_measure_choice")
                    sub = sub[sub["Effect_measure"].astype(str)==chosen_measure].copy()

                    studies, eff_os, lcl_os, ucl_os = [], [], [], []
                    effects_t, ses = [], []
                    skipped = []

                    for _, r in sub.iterrows():
                        label = f"{short(r.get('title',''), 60)} ({r.get('year','')})"
                        eff = float(r["Effect"])
                        lcl = float(r["Lower_CI"])
                        ucl = float(r["Upper_CI"])

                        se, err = se_from_ci_safe(eff, lcl, ucl, chosen_measure)
                        if err:
                            skipped.append({"study": label, "measure": chosen_measure, "effect": eff, "lcl": lcl, "ucl": ucl, "reason": err})
                            continue

                        studies.append(label)
                        eff_os.append(eff); lcl_os.append(lcl); ucl_os.append(ucl)
                        effects_t.append(transform_effect(eff, chosen_measure))
                        ses.append(se)

                    if skipped:
                        st.warning(f"有 {len(skipped)} 列因數值不合法被自動跳過（避免 app 跳掉）。請修正後可納入。")
                        with st.expander("查看被跳過的列（原因）", expanded=False):
                            st.dataframe(pd.DataFrame(skipped), use_container_width=True, height=220)

                    if len(studies) < 2:
                        st.error("可用研究數 < 2（扣除不合法列後）。請修正 CI/measure 或補齊更多研究。")
                        st.stop()

                    res = pool_fixed_random(effects_t, ses, chosen_measure)

                    model = st.radio("模型", options=["Fixed effect", "Random effects (DL)"], horizontal=True, key="ma_model_choice")
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

                    if HAS_PLOTLY:
                        fig = forest_plot_plotly(studies, eff_os, lcl_os, ucl_os, (pe, pl, pu), chosen_measure, model_label)
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info("環境缺少 Plotly：改以表格顯示森林圖資料。")
                        st.dataframe(pd.DataFrame({"study": studies, "effect": eff_os, "lcl": lcl_os, "ucl": ucl_os}), use_container_width=True)

    # -------------------- Diagnostics --------------------
    with tabs[6]:
        st.markdown("### Diagnostics")
        st.code(json.dumps(diag, ensure_ascii=False, indent=2), language="json")
