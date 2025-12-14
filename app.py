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
#   直接上傳至任何第三方服務或公開部署之網站（包含本 app 的雲端部署）。
# - 請遵守圖書館授權條款：避免大量下載/自動化批次擷取、避免共享全文給未授權者。
# - 若不確定是否可上傳：建議改用「本機版」或僅上傳你有權分享的開放取用全文（OA）。
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
import hashlib
from dataclasses import dataclass
from typing import Dict, List, Tuple, Optional, Any

import requests
import pandas as pd
import streamlit as st

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


# -------------------- Styles --------------------
CSS = """
<style>
.small { font-size: 0.88rem; color: #555; }
.muted { color: #6b7280; }
.kpi { border: 1px solid #e5e7eb; border-radius: 14px; padding: 0.75rem 0.9rem; background: #fafafa; }
.kpi .label { font-size: 0.82rem; color: #6b7280; }
.kpi .value { font-size: 1.25rem; font-weight: 800; color: #111827; }
.card { border: 1px solid #dde2eb; border-radius: 14px; padding: 0.9rem 1rem; background:#fff; margin-bottom: 0.9rem; }
.notice { border-left: 4px solid #f59e0b; background: #fff7ed; padding: 0.85rem 1rem; border-radius: 12px; }
.badge { display:inline-block; padding:0.15rem 0.55rem; border-radius: 999px; font-size:0.78rem; margin-right:0.35rem; border:1px solid rgba(0,0,0,0.06); }
.badge-ok { background:#d1fae5; color:#065f46; }
.badge-warn { background:#fef3c7; color:#92400e; }
.badge-bad { background:#fee2e2; color:#991b1b; }
hr.soft { border: none; border-top: 1px solid #eef2f7; margin: 0.8rem 0; }
pre code { font-size: 0.86rem !important; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

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


# -------------------- Helpers --------------------
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


# -------------------- Protocol model --------------------
@dataclass
class Protocol:
    P: str = ""
    I: str = ""
    C: str = ""
    O: str = ""
    NOT: str = "animal OR mice OR rat OR in vitro OR case report"
    goal_mode: str = "Fast / feasible (gap-fill)"  # or Rigorous / narrow scope

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
    feasibility_note: str = ""  # new: feasibility report / PICO adjustment guidance

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


# -------------------- Session state --------------------
def init_state():
    ss = st.session_state

    ss.setdefault("byok_enabled", False)
    ss.setdefault("byok_key", "")
    ss.setdefault("byok_base_url", "https://api.openai.com/v1")
    ss.setdefault("byok_model", "gpt-4o-mini")
    ss.setdefault("byok_temp", 0.2)
    ss.setdefault("byok_consent", False)

    ss.setdefault("question", "")
    ss.setdefault("protocol", Protocol(P_syn=[], I_syn=[], C_syn=[], O_syn=[], mesh_P=[], mesh_I=[], mesh_C=[], mesh_O=[]))

    ss.setdefault("pubmed_query", "")
    ss.setdefault("feas_query", "")
    ss.setdefault("pubmed_records", pd.DataFrame())
    ss.setdefault("srma_hits", pd.DataFrame())  # feasibility SR/MA/NMA list
    ss.setdefault("diagnostics", {})

    # Screening: keep AI + override (manual)
    ss.setdefault("ta_ai", {})            # record_id -> decision
    ss.setdefault("ta_ai_reason", {})     # record_id -> reason
    ss.setdefault("ta_override", {})      # record_id -> decision override
    ss.setdefault("ta_override_reason", {})  # record_id -> reason override

    # Extraction
    ss.setdefault("extract_df", pd.DataFrame())

    # RoB2
    ss.setdefault("rob2", {})

    # Manuscript sections
    ss.setdefault("ms_sections", {})  # dict of section->text
    ss.setdefault("ms_full_md", "")

init_state()


# -------------------- BYOK LLM --------------------
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


# -------------------- Sidebar --------------------
with st.sidebar:
    st.header("設定")

    st.subheader("LLM（使用者自備 key）")
    st.checkbox("啟用 LLM（BYOK）", key="byok_enabled", help="預設關閉；關閉時流程自動降級，不會卡在 AI extraction/ROB2。")

    st.markdown(
        "<div class='small muted'>"
        "Key only used for this session（不寫入 secrets、不落盤）。<br>"
        "請勿上傳可識別病人資訊。<br>"
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
    st.checkbox("Records：顯示逐篇卡片（可展開看全文 abstract）", value=True, key="show_record_cards")

    st.markdown("---")
    st.subheader("故障排除")
    st.markdown(
        "- PubMed count=0：通常是問題太短/縮寫未展開（例如只寫 EDOF）。\n"
        "- Diagnostics 回傳 HTML：可能被擋或限流，稍後重試。\n"
        "- 森林圖：優先 Plotly；若缺少 Plotly 會改用表格。\n"
    )


# -------------------- Template style guide (optional) --------------------
def read_docx_text(file_bytes: bytes, max_chars: int = 20000) -> str:
    if not HAS_DOCX:
        return ""
    d = docx.Document(io.BytesIO(file_bytes))
    paras = [p.text.strip() for p in d.paragraphs if p.text and p.text.strip()]
    return ("\n".join(paras))[:max_chars]


# -------------------- PICO parsing + expansions --------------------
ABBR_MAP = {
    "EDOF": ["extended depth of focus", "extended depth-of-focus", "extended range of vision", "extended range-of-vision"],
    "IOL": ["intraocular lens", "intra-ocular lens"],
    "RCT": ["randomized controlled trial", "randomised controlled trial"],
    "NMA": ["network meta-analysis", "network meta analysis"],
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
    syn = []
    parts = re.split(r"[;,/]+", text)
    for p in parts:
        p = p.strip()
        if not p:
            continue
        syn.append(p)
        key = p.upper()
        if key in ABBR_MAP:
            syn.extend(ABBR_MAP[key])
        toks = re.findall(r"[A-Za-z]{2,10}", p)
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

    P = ""
    I = left
    C = right
    O = ""

    if I and C and I.strip().lower() == C.strip().lower():
        C = "其他比較組（例如不同型號/設計）"

    proto = Protocol(P=P, I=I, C=C, O=O)
    proto.P_syn = expand_terms(proto.P)
    proto.I_syn = expand_terms(proto.I)
    proto.C_syn = expand_terms(proto.C)
    proto.O_syn = expand_terms(proto.O)

    proto.mesh_P = propose_mesh_candidates(proto.P_syn)
    proto.mesh_I = propose_mesh_candidates(proto.I_syn)
    proto.mesh_C = propose_mesh_candidates(proto.C_syn)
    proto.mesh_O = propose_mesh_candidates(proto.O_syn)

    return proto


# -------------------- PubMed query builder --------------------
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

def build_pubmed_query(proto: Protocol) -> str:
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
    return f"({core}) NOT ({not_block})" if not_block else core

def build_feasibility_query(pubmed_query: str) -> str:
    sr_filter = '(systematic review[pt] OR meta-analysis[pt] OR "systematic review"[tiab] OR "meta-analysis"[tiab] OR "network meta-analysis"[tiab] OR NMA[tiab])'
    return f"({pubmed_query}) AND {sr_filter}"


# -------------------- PubMed E-utilities --------------------
EUTILS = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils"

def pubmed_esearch(term: str, retmax: int = 200, retstart: int = 0) -> Tuple[int, List[str], str, Dict[str, Any]]:
    params = {"db": "pubmed", "term": term, "retmode": "json", "retmax": retmax, "retstart": retstart}
    url = f"{EUTILS}/esearch.fcgi"
    r = requests.get(url, params=params, timeout=30)
    text = r.text or ""
    diag = {"status_code": r.status_code, "content_type": r.headers.get("content-type", ""), "body_head": text[:200]}
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

        rows.append({"pmid": pmid, "year": year, "title": title, "abstract": abstract, "doi": doi})

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["pmid"] = df["pmid"].astype(str)
    df["record_id"] = df["pmid"].apply(lambda x: f"PMID:{x}")
    return df


# -------------------- AI prompts --------------------
def build_protocol_llm_prompt(question: str, proto: Protocol) -> List[Dict[str, str]]:
    sys = (
        "你是資深系統性回顧/統合分析（SR/MA）研究助理。"
        "請用繁體中文輸出 JSON（不可夾雜多餘文字）。"
        "你必須：\n"
        "1) 從一句問題推導 PICO（不確定用『』保留）\n"
        "2) 在 PICO 層級擬定 inclusion/exclusion（並指出：Fast/gap-fill vs rigorous 的取捨需要人工決策）\n"
        "3) outcomes 規劃需同時考量：既有 SR/MA/NMA 常用 outcomes + 過去 RCT primary/secondary outcomes\n"
        "4) extraction sheet 不可寫死欄位：提出『欄位類別』與可擴充 outcomes 欄（含 effect/CI）\n"
        "5) 給出 feasibility_note：如何調整 PICO 以提高可行性（例如縮小族群/特定型號/特定 timepoint/outcome）\n"
        "注意：不得捏造不存在的研究結果；不得要求帳密；不得輸出病人可識別資訊。"
    )
    user = {"question": question, "current_guess_protocol": proto.to_dict()}
    return [{"role": "system", "content": sys}, {"role": "user", "content": json.dumps(user, ensure_ascii=False)}]

def try_llm_fill_protocol(question: str, proto: Protocol) -> Protocol:
    if not llm_available():
        return proto
    try:
        content = call_openai_compatible(build_protocol_llm_prompt(question, proto), max_tokens=1600)
        js = json.loads(content)

        pico = js.get("pico", {}) or {}
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

        feas = js.get("feasibility", {}) or {}
        proto.feasibility_note = norm_text(feas.get("note", proto.feasibility_note))

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

        # normalize lists
        for k in ["P_syn","I_syn","C_syn","O_syn","mesh_P","mesh_I","mesh_C","mesh_O"]:
            v = getattr(proto, k, None) or []
            setattr(proto, k, [norm_text(x) for x in v if norm_text(x)])

        return proto
    except Exception:
        return proto

def rule_based_ta_screen(title: str, abstract: str, proto: Protocol) -> Tuple[str, str]:
    blob = ((title or "") + " " + (abstract or "")).lower()
    if re.search(r"\b(mice|mouse|rat|porcine|rabbit|canine)\b", blob):
        return "Exclude", "疑似動物/非人體研究（rule-based）"

    i_terms = proto.I_syn or expand_terms(proto.I)
    hit = any((t.lower() in blob) for t in i_terms[:25] if t)
    if not hit and proto.I:
        return "Unsure", "未偵測到明顯介入關鍵詞（rule-based；可能縮寫/同義詞不足）"
    return "Include", "符合基本關鍵詞訊號（rule-based）"

def ta_screen_with_llm(df: pd.DataFrame, proto: Protocol) -> Dict[str, Dict[str, str]]:
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
        "你是系統性回顧的 title/abstract 粗篩評讀者（等同兩位評讀者合併建議）。"
        "請用繁體中文輸出 JSON（不可夾雜多餘文字）。\n"
        "規則：decision 只能是 Include / Exclude / Unsure；reason 要可核對且簡短。\n"
        "不得捏造全文內容；若資訊不足，請選 Unsure。"
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
    user = {"protocol": proto.to_dict(), "records": records[:120]}
    try:
        content = call_openai_compatible(
            [{"role":"system","content":sys},{"role":"user","content":json.dumps(user, ensure_ascii=False)}],
            max_tokens=1800
        )
        js = json.loads(content)
        items = js.get("decisions", js)
        if isinstance(items, dict):
            for rid, v in items.items():
                if isinstance(v, dict):
                    out[rid] = {"decision": v.get("decision","Unsure"), "reason": v.get("reason","")}
        elif isinstance(items, list):
            for v in items:
                rid = v.get("record_id")
                if rid:
                    out[rid] = {"decision": v.get("decision","Unsure"), "reason": v.get("reason","")}
        return out
    except Exception:
        for _, r in df.iterrows():
            rid = r["record_id"]
            d, rs = rule_based_ta_screen(r.get("title",""), r.get("abstract",""), proto)
            out[rid] = {"decision": d, "reason": rs}
        return out


# -------------------- Extraction prompt (keeps your senior requirements) --------------------
def extraction_prompt(proto: Protocol) -> str:
    return f"""
你是系統性回顧與統合分析的資料萃取助理。請依照下列 Protocol 抽取資料，輸出 JSON。

【Protocol（PICO + criteria）】
P: {proto.P or "『』"}
I: {proto.I or "『』"}
C: {proto.C or "『』"}
O: {proto.O or "『』"}

Inclusion（PICO 層級；不確定用『』）：
{proto.inclusion or "『請根據現有證據/臨床情境擬定；並標示 Fast/gap-fill vs rigorous 的取捨需要人工決策』"}

Exclusion（PICO 層級；不確定用『』）：
{proto.exclusion or "『』"}

【重要：OCR / Figure / Table】
- 若全文 outcome 數據只在 Figure/Table：請主動提示「需要 OCR」，並嘗試從表格/圖說擷取數值。
- 若 OCR 仍無法讀取，請回傳空白並在 notes 說明缺漏位置（例如：Table 2 無法辨識）。

【Extraction sheet（不要寫死欄位；PICO 層級規劃）】
請先規劃「本題目應有的 extraction sheet 欄位類別」，至少包含：
1) Study characteristics（作者、年份、設計、國家、樣本數、追蹤期）
2) Population baseline（年齡、性別、納入條件關鍵、重要共病）
3) Intervention/Comparator details（器材/術式/型號/參數/時間點）
4) Outcomes 規劃：同時考量
   - 既有 SR/MA/NMA 常用 outcomes（若有）
   - 過去 RCT 的 primary/secondary outcomes（若有）
5) Effect size for MA（若可）：effect_measure（OR/RR/HR/MD/SMD）、effect、lower_CI、upper_CI、timepoint、unit
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
  "extracted_fields": {{ "...": "..." }},
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


# -------------------- Meta-analysis + forest plot --------------------
def se_from_ci(effect: float, lcl: float, ucl: float, measure: str) -> float:
    measure = (measure or "").upper().strip()
    if measure in {"OR","RR","HR"}:
        return (math.log(ucl) - math.log(lcl)) / 3.92
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

    pooled_y = -1
    pe, pl, pu = pooled
    fig.add_trace(go.Scatter(
        x=[pe], y=[pooled_y], mode="markers",
        marker=dict(symbol="diamond", size=12),
        error_x=dict(type="data", symmetric=False, array=[pu-pe], arrayminus=[pe-pl]),
        hovertext=[f"Pooled {model_label}"],
        showlegend=False
    ))

    ytickvals = y + [pooled_y]
    yticktext = studies[::-1] + [f"Pooled ({model_label})"]

    fig.update_layout(
        height=360 + 18*len(studies),
        xaxis_title=f"Effect ({measure})",
        yaxis=dict(tickmode="array", tickvals=ytickvals, ticktext=yticktext),
        margin=dict(l=10, r=10, t=35, b=10),
        showlegend=False,
    )

    if (measure or "").upper().strip() in {"OR","RR","HR"}:
        fig.add_vline(x=1.0, line_width=1, line_dash="dash")
    else:
        fig.add_vline(x=0.0, line_width=1, line_dash="dash")
    return fig


# -------------------- Manuscript: split sections --------------------
def manuscript_skeleton_sections(proto: Protocol, prisma: Dict[str,Any], ma_summary: Optional[Dict[str,Any]]) -> Dict[str, str]:
    P = proto.P or "『』"
    I = proto.I or "『』"
    C = proto.C or "『』"
    O = proto.O or "『』"

    n_records = prisma.get("records", "『』")
    n_dups = prisma.get("duplicates_removed", "『』")
    n_screened = prisma.get("screened", "『』")
    n_fulltext = prisma.get("fulltext_assessed", "『』")
    n_included = prisma.get("included", "『』")

    if ma_summary:
        meas = ma_summary.get("measure","『』")
        model = ma_summary.get("model","『』")
        eff = ma_summary.get("effect","『』")
        lcl = ma_summary.get("lcl","『』")
        ucl = ma_summary.get("ucl","『』")
        I2 = ma_summary.get("I2","『』")
        k = ma_summary.get("k","『』")
        outcome = ma_summary.get("outcome","『』")
    else:
        meas = model = eff = lcl = ucl = I2 = k = outcome = "『』"

    intro = f"""在『{P}』情境下，『{I}』與『{C}』常用以改善『{O}』。然而，現有研究族群、介入細節與評估時間點可能不一致，導致結論分歧。本研究旨在系統性整合現有證據，以比較『{I}』與『{C}』在『{P}』上的 outcomes 表現，並評估研究間異質性與限制。"""

    methods = f"""本研究依循 PRISMA 指南進行系統性回顧與統合分析。研究問題以 PICO 架構界定：Population『{P}』、Intervention『{I}』、Comparator『{C}』、Outcomes『{O}』。納入/排除標準於 PICO 層級擬定（必要處以『』保留）：
- Inclusion：{proto.inclusion or "『』"}
- Exclusion：{proto.exclusion or "『』"}

資料庫以 PubMed/Medline 為主（可擴充：『EMBASE/CENTRAL/Scopus/Web of Science』），並先進行既有 SR/MA/NMA 可行性掃描以調整 PICO 範圍（Fast/gap-fill vs rigorous 需人工決策）。資料萃取採雙人獨立，若數據位於 figure/table，需 OCR/人工核對。偏倚風險評估採 RoB 2.0（最終需人工確認）。統計以反變異數法合併效應，異質性以 Q 與 I² 評估，必要時使用隨機效應模型。"""

    results = f"""PRISMA 流程摘要：records={n_records}、duplicates removed={n_dups}、screened={n_screened}、full-text assessed={n_fulltext}、included={n_included}。
在 outcome『{outcome}』之統合分析中，共納入 {k} 篇研究。結果顯示（{model} 模型，{meas}）合併效應為 {eff}（95% CI {lcl}–{ucl}），異質性 I² = {I2}%。"""

    discussion = f"""本研究整合現有證據以比較『{I}』與『{C}』於『{P}』之 outcomes 表現。主要發現提示『請依效應方向補一句結論；缺資料用『』』。異質性（I² = {I2}%）可能來自族群差異、介入細節、追蹤時間點與 outcome 定義不一。限制包含：研究數量與品質、報告不一致、publication bias 評估之不確定性。未來仍需標準化 outcome 的高品質研究以確認『』。"""

    conclusion = f"""在『{P}』中，『{I}』相較於『{C}』在『{O}』上顯示『（優勢/無差異/仍不確定）』。本結論需結合偏倚風險與異質性審慎解讀。"""

    appendix = """- PubMed search string：『』
- 其他資料庫搜尋式：『』
- PRISMA checklist：『』"""

    return {
        "標題": f"『{I}』相較於『{C}』於『{P}』之系統性回顧與統合分析",
        "Introduction": intro,
        "Methods": methods,
        "Results": results,
        "Discussion": discussion,
        "Conclusion": conclusion,
        "Appendix": appendix,
        "授權/校內資源提醒": "若全文來自校內訂閱/付費期刊，請遵守圖書館授權條款，避免將受版權保護之全文上傳到雲端服務或公開部署環境；避免大量自動化下載或共享給未授權者。"
    }

def draft_manuscript_sections_with_llm(proto: Protocol, prisma: Dict[str,Any], ma_summary: Optional[Dict[str,Any]]) -> Dict[str, str]:
    if not llm_available():
        return manuscript_skeleton_sections(proto, prisma, ma_summary)

    sys = (
        "你是系統性回顧與統合分析的寫作助手。請用繁體中文產出『分段』稿件，並以 JSON 回傳。\n"
        "要求：\n"
        "- sections 必須包含：標題、Introduction、Methods、Results、Discussion、Conclusion、Appendix、授權/校內資源提醒\n"
        "- 不得捏造不存在的研究或數據；缺資料處請用『』保留\n"
        "- 需保留偏倚風險與校內資源授權提醒\n"
        "輸出格式：{ \"sections\": { ... } }"
    )
    user = {"protocol": proto.to_dict(), "prisma": prisma, "ma_summary": ma_summary or {}}
    try:
        content = call_openai_compatible(
            [{"role":"system","content":sys},{"role":"user","content":json.dumps(user, ensure_ascii=False)}],
            max_tokens=1800
        )
        js = json.loads(content)
        secs = js.get("sections", {})
        if isinstance(secs, dict) and secs:
            return {k: str(v) for k, v in secs.items()}
        return manuscript_skeleton_sections(proto, prisma, ma_summary)
    except Exception:
        return manuscript_skeleton_sections(proto, prisma, ma_summary)

def sections_to_markdown(secs: Dict[str,str]) -> str:
    title = secs.get("標題", "『』")
    md = [f"# {title}", ""]
    md.append("## 免責聲明（學術用途）")
    md.append("本稿件由工具自動產生初稿，僅供學術研究與教學用途；不構成醫療建議或法律意見。所有資料、引用、數值與結論需由作者團隊逐一核對後方可使用。請勿輸入或上傳任何可識別之病人資訊。")
    md.append("")
    md.append("## 授權/校內資源提醒")
    md.append(secs.get("授權/校內資源提醒","『』"))
    md.append("")
    for k in ["Introduction","Methods","Results","Discussion","Conclusion","Appendix"]:
        md.append(f"## {k}")
        md.append(secs.get(k,"『』"))
        md.append("")
    return "\n".join(md).strip()


# -------------------- UI: question + run --------------------
st.subheader("Research question（輸入一句話）")
st.session_state["question"] = st.text_input(
    "例：『不同種類 EDOF IOL 於白內障術後視覺品質（對比敏感度/眩光）比較』或『FLACS 是否優於傳統 phaco』",
    value=st.session_state.get("question",""),
)

with st.expander("開始前檢查清單（建議）", expanded=False):
    st.markdown(
        "- 問題是否包含：族群/情境 + 介入 + 比較 +（主要 outcome）？\n"
        "- 是否只有縮寫（例如 EDOF）？建議加上全名或具體型號/術式。\n"
        "- 若要用 AI：請在側邊欄啟用 LLM 並勾選同意。\n"
        "- 校內訂閱全文：避免上傳雲端；若需全文分析請改本機跑。\n"
    )

run = st.button("Run / 執行（Step 0～7 自動跑到 Outputs）", type="primary")


# -------------------- Pipeline run --------------------
if run:
    q = norm_text(st.session_state["question"])
    if not q:
        st.error("請先輸入一句研究問題。")
        st.stop()

    # Step 0: Protocol
    with st.spinner("Step 0/7：生成 protocol（PICO/criteria/outcomes/extraction/feasibility）…"):
        proto0 = question_to_protocol(q)
        proto = try_llm_fill_protocol(q, proto0)

        # fill expansions if empty
        proto.P_syn = proto.P_syn or expand_terms(proto.P)
        proto.I_syn = proto.I_syn or expand_terms(proto.I)
        proto.C_syn = proto.C_syn or expand_terms(proto.C)
        proto.O_syn = proto.O_syn or expand_terms(proto.O)
        proto.mesh_P = proto.mesh_P or propose_mesh_candidates(proto.P_syn)
        proto.mesh_I = proto.mesh_I or propose_mesh_candidates(proto.I_syn)
        proto.mesh_C = proto.mesh_C or propose_mesh_candidates(proto.C_syn)
        proto.mesh_O = proto.mesh_O or propose_mesh_candidates(proto.O_syn)
        st.session_state["protocol"] = proto

    # Step 1: build query
    with st.spinner("Step 1/7：產出 PubMed 搜尋式（MeSH + free text）…"):
        pub_q = build_pubmed_query(proto)
        st.session_state["pubmed_query"] = pub_q

    # Step 2: feasibility scan (SR/MA/NMA)
    with st.spinner("Step 2/7：可行性掃描（既有 SR/MA/NMA）…"):
        feas_q = build_feasibility_query(pub_q)
        st.session_state["feas_query"] = feas_q

        cnt_feas, ids_feas, feas_url, feas_diag = pubmed_esearch(feas_q, retmax=20, retstart=0)
        xml_feas, _ = pubmed_efetch_xml(ids_feas[:20])
        df_feas = parse_pubmed_xml_minimal(xml_feas)
        st.session_state["srma_hits"] = df_feas

    # Step 3: retrieve pubmed
    with st.spinner("Step 3/7：抓取 PubMed 文獻…"):
        total, ids, es_url, es_diag = pubmed_esearch(pub_q, retmax=200, retstart=0)
        xml, ef_urls = pubmed_efetch_xml(ids[:200])
        df = parse_pubmed_xml_minimal(xml)

        st.session_state["pubmed_records"] = df
        st.session_state["diagnostics"] = {
            "pubmed_total_count": total,
            "esearch_url": es_url,
            "efetch_urls": ef_urls,
            "esearch_diag": es_diag,
            "feasibility": {"count": cnt_feas, "esearch_url": feas_url, "diag": feas_diag},
            "warnings": [] if total > 0 else ["PubMed count=0：請檢查問題是否過短/縮寫未展開，或 PubMed 回應被阻擋。"],
        }

    # Step 4: TA screening AI/rule-based
    with st.spinner("Step 4/7：Title/Abstract 粗篩（保留 AI 判讀 + 允許人工 override）…"):
        df = st.session_state.get("pubmed_records", pd.DataFrame())
        if df is not None and not df.empty:
            decisions = ta_screen_with_llm(df, st.session_state["protocol"])
            for rid, v in decisions.items():
                st.session_state["ta_ai"][rid] = v.get("decision","Unsure")
                st.session_state["ta_ai_reason"][rid] = v.get("reason","")
        st.success("Done。請往下查看 Outputs。")


# =========================================================
# Outputs (Beautiful + preserves advantages)
# =========================================================
if st.session_state.get("question"):
    proto: Protocol = st.session_state.get("protocol")
    df = st.session_state.get("pubmed_records", pd.DataFrame())
    df_feas = st.session_state.get("srma_hits", pd.DataFrame())
    diag = st.session_state.get("diagnostics", {}) or {}

    tabs = st.tabs(["總覽", "Step 0 Protocol", "Step 1-2 檢索/可行性", "Step 3 Records", "Step 4 粗篩（AI+人工）", "Step 5-6 萃取/MA/森林圖", "Step 7 稿件（分段顯示）", "Diagnostics"])

    # -------------------- Overview --------------------
    with tabs[0]:
        st.markdown("### 流程總覽")
        c1, c2, c3, c4 = st.columns(4)

        total = int(diag.get("pubmed_total_count", 0) or 0)
        feas_cnt = int((diag.get("feasibility", {}) or {}).get("count", 0) or 0)

        # effective include = override if set else AI
        includes = 0
        if df is not None and not df.empty:
            for rid in df["record_id"].tolist():
                od = st.session_state["ta_override"].get(rid, "").strip()
                ai = st.session_state["ta_ai"].get(rid, "Unsure")
                effective = od if od else ai
                if effective == "Include":
                    includes += 1

        with c1:
            st.markdown(f"<div class='kpi'><div class='label'>PubMed count</div><div class='value'>{total}</div></div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"<div class='kpi'><div class='label'>既有 SR/MA/NMA（可行性）</div><div class='value'>{feas_cnt}</div></div>", unsafe_allow_html=True)
        with c3:
            st.markdown(f"<div class='kpi'><div class='label'>TA 粗篩 Include（有效決策）</div><div class='value'>{includes}</div></div>", unsafe_allow_html=True)
        with c4:
            st.markdown(f"<div class='kpi'><div class='label'>LLM 狀態</div><div class='value'>{'ON' if llm_available() else 'OFF'}</div></div>", unsafe_allow_html=True)

        st.markdown("<hr class='soft'/>", unsafe_allow_html=True)
        st.markdown("**你現在應該做什麼？**")
        if total == 0:
            st.warning("PubMed 沒抓到文獻：請把研究問題寫更具體（縮寫寫全名/加上型號/術式/族群/outcome），或看 Diagnostics 是否被擋。")
        elif includes == 0:
            st.info("已有 records，但尚未 Include：請到 Step 4 檢視 AI 判讀與 abstract，必要時手動 override。")
        else:
            st.success("已有 Include：請到 Step 5-6 填寫/確認 effect + CI，產出 MA/森林圖；再到 Step 7 產生分段稿件。")

    # -------------------- Step 0 Protocol --------------------
    with tabs[1]:
        st.markdown("### Step 0：Protocol（PICO / criteria / plans / feasibility note）")
        left, right = st.columns([1.05, 0.95])

        with left:
            st.markdown("#### 目前 protocol（JSON）")
            st.code(json.dumps(proto.to_dict(), ensure_ascii=False, indent=2), language="json")

        with right:
            st.markdown("#### （可選）人工微調（展開才改）")
            with st.expander("展開修改 PICO/criteria（不改也可繼續）", expanded=False):
                proto.P = st.text_input("P", value=proto.P)
                proto.I = st.text_input("I", value=proto.I)
                proto.C = st.text_input("C", value=proto.C)
                proto.O = st.text_input("O", value=proto.O)
                proto.NOT = st.text_input("NOT", value=proto.NOT)
                proto.goal_mode = st.selectbox("Goal mode", options=["Fast / feasible (gap-fill)", "Rigorous / narrow scope"],
                                              index=0 if "Fast" in (proto.goal_mode or "") else 1)

                proto.inclusion = st.text_area("Inclusion criteria（PICO 層級）", value=proto.inclusion, height=110)
                proto.exclusion = st.text_area("Exclusion criteria（PICO 層級）", value=proto.exclusion, height=110)

                proto.outcomes_plan = st.text_area("Outcomes 規劃（SR/MA + RCT primary/secondary）", value=proto.outcomes_plan, height=110)
                proto.extraction_plan = st.text_area("Extraction sheet 規劃（不要寫死）", value=proto.extraction_plan, height=110)
                proto.feasibility_note = st.text_area("Feasibility note（如何調整 PICO 提高可行性）", value=proto.feasibility_note, height=90)

                st.session_state["protocol"] = proto

        st.markdown("<hr class='soft'/>", unsafe_allow_html=True)
        st.markdown("#### 學長要求（對照）")
        st.markdown(
            "- 在開始所有步驟前：PICO 完成後先做既有 SR/MA/NMA 可行性掃描（見 Step 1-2）。\n"
            "- inclusion criteria 決策寫在 PICO 層級，並保留 Fast/gap-fill vs rigorous 的人工取捨。\n"
            "- extraction table 不寫死：以 PICO 層級規劃欄位類別與 outcomes（見 extraction_plan + Step 5）。\n"
        )

    # -------------------- Step 1-2 Search + feasibility --------------------
    with tabs[2]:
        st.markdown("### Step 1：PubMed 搜尋式（可複製）")
        st.code(st.session_state.get("pubmed_query",""), language="text")

        st.markdown("### Step 2：可行性掃描（既有 SR/MA/NMA）")
        st.code(st.session_state.get("feas_query",""), language="text")

        feas = (diag.get("feasibility", {}) or {})
        st.markdown(f"- 既有 SR/MA/NMA count：**{feas.get('count','')}**")
        if proto.feasibility_note:
            st.markdown("**Feasibility note（自動/半自動）**")
            st.info(proto.feasibility_note)

        if df_feas is not None and not df_feas.empty:
            st.markdown("**既有 SR/MA/NMA（前幾篇）**")
            show = df_feas[["record_id","year","title","doi"]].copy()
            st.dataframe(show, use_container_width=True, height=280)
            with st.expander("展開查看摘要（如有）", expanded=False):
                for _, r in df_feas.iterrows():
                    st.markdown(f"**{r.get('title','')}**  ({r.get('year','')})  — {r.get('record_id','')}")
                    abst = r.get("abstract","")
                    st.markdown(abst if abst else "_（無 abstract）_")
                    st.markdown("<hr class='soft'/>", unsafe_allow_html=True)
        else:
            st.caption("（未抓到 SR/MA/NMA 列表，可能是 count=0 或 PubMed 回應被擋）")

    # -------------------- Step 3 Records --------------------
    with tabs[3]:
        st.markdown("### Step 3：Records（含 abstract）")
        if df is None or df.empty:
            st.warning("沒有抓到 records。建議：把問題寫更清楚（族群+介入+比較+outcome；縮寫寫全名或型號），或看 Diagnostics 是否被擋。")
        else:
            ensure_columns(df, ["record_id","pmid","year","doi","title","abstract"], "")
            cols = ["record_id","year","pmid","doi","title"] + (["abstract"] if st.session_state.get("show_abs_in_table", True) else [])
            st.dataframe(df[cols], use_container_width=True, height=420)

            if st.session_state.get("show_record_cards", True):
                st.markdown("#### 逐篇卡片（可展開看 abstract 全文）")
                # filter UI
                qf = st.text_input("快速搜尋（標題/摘要關鍵字）", value="", key="record_filter_q")
                fdf = df.copy()
                if qf.strip():
                    qq = qf.strip().lower()
                    fdf = fdf[(fdf["title"].str.lower().str.contains(qq, na=False)) | (fdf["abstract"].str.lower().str.contains(qq, na=False))]

                for _, r in fdf.head(80).iterrows():
                    rid = r["record_id"]
                    title = r.get("title","")
                    year = r.get("year","")
                    doi = r.get("doi","")
                    pmid = r.get("pmid","")
                    ai_d = st.session_state["ta_ai"].get(rid, "Unsure")
                    ai_r = st.session_state["ta_ai_reason"].get(rid, "")
                    od = st.session_state["ta_override"].get(rid, "").strip()
                    effective = od if od else ai_d

                    badge = "badge-warn"
                    if effective == "Include":
                        badge = "badge-ok"
                    elif effective == "Exclude":
                        badge = "badge-bad"

                    st.markdown(
                        f"<div class='card'>"
                        f"<span class='badge {badge}'>Effective: {effective}</span>"
                        f"<span class='badge badge-warn'>AI: {ai_d}</span>"
                        f"<div style='font-weight:800; font-size:1.02rem; margin-top:0.35rem;'>{html.escape(title)}</div>"
                        f"<div class='small muted'>Year: {year} | PMID: {pmid} | DOI: {doi}</div>"
                        f"</div>",
                        unsafe_allow_html=True
                    )

                    with st.expander("展開：Abstract / AI 解釋 / 人工 override", expanded=False):
                        st.markdown("**Abstract**")
                        st.write(r.get("abstract","") or "（無 abstract）")

                        st.markdown("**AI 判讀（保留）**")
                        st.write(f"- Decision: **{ai_d}**")
                        st.write(f"- Reason: {ai_r or '（無）'}")

                        st.markdown("**人工 override（可選；不會覆蓋 AI 原判讀）**")
                        col1, col2 = st.columns([0.35, 0.65])
                        with col1:
                            newd = st.selectbox("Override decision", options=["", "Include","Exclude","Unsure"],
                                                index=["", "Include","Exclude","Unsure"].index(od) if od in ["","Include","Exclude","Unsure"] else 0,
                                                key=f"od_{rid}")
                        with col2:
                            newr = st.text_input("Override reason（可空）", value=st.session_state["ta_override_reason"].get(rid,""),
                                                 key=f"or_{rid}")

                        st.session_state["ta_override"][rid] = newd
                        st.session_state["ta_override_reason"][rid] = newr

                st.caption("卡片預設最多顯示 80 篇（避免頁面太重）。可用上方搜尋縮小。")

    # -------------------- Step 4 Screening table (AI + manual) --------------------
    with tabs[4]:
        st.markdown("### Step 4：Title/Abstract 粗篩（AI 判讀保留 + 人工修正）")
        if df is None or df.empty:
            st.info("沒有 records。")
        else:
            st.markdown("**批次操作**")
            colA, colB, colC = st.columns([0.22, 0.22, 0.56])
            with colA:
                if st.button("重新執行粗篩（用目前 protocol）"):
                    decisions = ta_screen_with_llm(df, st.session_state["protocol"])
                    for rid, v in decisions.items():
                        st.session_state["ta_ai"][rid] = v.get("decision","Unsure")
                        st.session_state["ta_ai_reason"][rid] = v.get("reason","")
                    st.success("已更新 AI 粗篩結果（AI 判讀保留；override 不變）。")
            with colB:
                if st.button("清除所有 override"):
                    st.session_state["ta_override"] = {}
                    st.session_state["ta_override_reason"] = {}
                    st.success("已清除 override。")
            with colC:
                st.caption("Effective decision = Override（若有）否則 AI。AI decision/reason 永遠保留。")

            rows = []
            for _, r in df.iterrows():
                rid = r["record_id"]
                ai_d = st.session_state["ta_ai"].get(rid, "Unsure")
                ai_r = st.session_state["ta_ai_reason"].get(rid, "")
                od = st.session_state["ta_override"].get(rid, "").strip()
                orr = st.session_state["ta_override_reason"].get(rid, "")
                eff = od if od else ai_d

                rows.append({
                    "record_id": rid,
                    "year": r.get("year",""),
                    "title": r.get("title",""),
                    "AI_decision": ai_d,
                    "AI_reason": ai_r,
                    "Override_decision": od,
                    "Override_reason": orr,
                    "Effective_decision": eff,
                })

            sdf = pd.DataFrame(rows)
            st.dataframe(sdf, use_container_width=True, height=420)

            st.download_button("下載粗篩結果（含 AI+override）", data=to_csv_bytes(sdf), file_name="screening_ta_ai_override.csv", mime="text/csv")

            # PRISMA prototype counts based on effective decision
            eff_includes = int((sdf["Effective_decision"] == "Include").sum())
            prisma = {
                "records": int(len(df)),
                "duplicates_removed": 0,
                "screened": int(len(df)),
                "fulltext_assessed": eff_includes,
                "included": eff_includes,
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

    # -------------------- Step 5-6 Extraction + MA + forest --------------------
    with tabs[5]:
        st.markdown("### Step 5：Extraction 寬表（不寫死欄位）")
        if df is None or df.empty:
            st.info("沒有 records。")
        else:
            # Build include set based on effective decision
            rows_eff = []
            for rid in df["record_id"].tolist():
                od = st.session_state["ta_override"].get(rid, "").strip()
                ai_d = st.session_state["ta_ai"].get(rid, "Unsure")
                eff = od if od else ai_d
                if eff == "Include":
                    rows_eff.append(rid)

            cands = df[df["record_id"].isin(rows_eff)].copy()
            if cands.empty:
                st.warning("目前沒有 Effective=Include 的研究。請先到 Step 4 檢視 abstract 與 override。")
            else:
                base = cands[["record_id","pmid","year","doi","title"]].copy()
                ensure_columns(base, [
                    "Study_design","Country","N_total","Follow_up",
                    "Population_key","Intervention_details","Comparator_details",
                    "Outcome_label","Timepoint",
                    "Effect_measure","Effect","Lower_CI","Upper_CI","Effect_unit",
                    "Notes"
                ], default="")

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

                st.caption("先填 Effect_measure + Effect + CI（與 Outcome_label/Timepoint）即可做 MA/森林圖；其他欄位可後補。")
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

                with st.expander("（可選）AI extraction prompt（含 OCR/figure/table 提示；符合學長要求）", expanded=False):
                    st.code(extraction_prompt(proto), language="text")
                    if not llm_available():
                        st.info("目前未啟用 LLM：此處提供 prompt 作為手動抽取/外部 LLM 使用。")

                st.markdown("<hr class='soft'/>", unsafe_allow_html=True)
                st.markdown("### Step 6：MA + 森林圖（Fixed/Random）")

                dfm = ex.copy()
                ensure_columns(dfm, ["Outcome_label","Effect_measure","Effect","Lower_CI","Upper_CI","Timepoint"], "")
                for c in ["Effect","Lower_CI","Upper_CI"]:
                    dfm[c] = pd.to_numeric(dfm[c], errors="coerce")
                dfm = dfm.dropna(subset=["Effect","Lower_CI","Upper_CI"])
                dfm = dfm[dfm["Effect_measure"].astype(str).str.strip() != ""]

                ma_summary = None
                if dfm.empty:
                    st.info("請至少填入：Effect_measure + Effect + Lower_CI + Upper_CI（可加 Outcome_label/Timepoint）才能做 MA/森林圖。")
                else:
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
                    else:
                        studies, effects_t, ses = [], [], []
                        for _, r in sub.iterrows():
                            studies.append(f"{short(r.get('title',''), 60)} ({r.get('year','')})")
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

                        eff_os = sub["Effect"].astype(float).tolist()
                        lcl_os = sub["Lower_CI"].astype(float).tolist()
                        ucl_os = sub["Upper_CI"].astype(float).tolist()

                        if HAS_PLOTLY:
                            fig = forest_plot_plotly(studies, eff_os, lcl_os, ucl_os, (pe, pl, pu), chosen_measure, model_label)
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.warning("環境缺少 Plotly：改以表格顯示森林圖資料。")
                            st.dataframe(pd.DataFrame({"study": studies, "effect": eff_os, "lcl": lcl_os, "ucl": ucl_os}),
                                         use_container_width=True)

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

                # Store MA summary for manuscript use
                st.session_state["last_prisma"] = {
                    "records": int(len(df)),
                    "duplicates_removed": 0,
                    "screened": int(len(df)),
                    "fulltext_assessed": int(len(cands)),
                    "included": int(len(cands)),
                }
                st.session_state["last_ma_summary"] = ma_summary or {}

    # -------------------- Step 7 Manuscript --------------------
    with tabs[6]:
        st.markdown("### Step 7：自動撰寫稿件（分段呈現，可編輯）")
        prisma = st.session_state.get("last_prisma", {"records":"『』","duplicates_removed":"『』","screened":"『』","fulltext_assessed":"『』","included":"『』"})
        ma_summary = st.session_state.get("last_ma_summary", {}) or None

        col1, col2 = st.columns([0.62, 0.38])
        with col1:
            st.markdown("**生成方式**")
            st.caption("LLM ON：以 AI 產出分段稿件；LLM OFF：用模板產出（缺資料用『』）。")
        with col2:
            if st.button("產生/更新稿件（分段）", type="primary"):
                secs = draft_manuscript_sections_with_llm(proto, prisma, ma_summary)
                st.session_state["ms_sections"] = secs
                st.session_state["ms_full_md"] = sections_to_markdown(secs)
                st.success("已產生分段稿件。")

        secs = st.session_state.get("ms_sections", {}) or {}
        if not secs:
            st.info("尚未產生稿件。點上方「產生/更新稿件（分段）」即可。")
        else:
            st.markdown("#### 分段顯示（可直接在頁面修改）")
            sec_tabs = st.tabs(["標題", "Introduction", "Methods", "Results", "Discussion", "Conclusion", "Appendix", "授權提醒", "整篇 Markdown"])
            order = ["標題","Introduction","Methods","Results","Discussion","Conclusion","Appendix","授權/校內資源提醒"]

            # Title
            with sec_tabs[0]:
                secs["標題"] = st.text_area("標題", value=secs.get("標題",""), height=90)

            # Main sections
            mapping = {
                1:"Introduction", 2:"Methods", 3:"Results", 4:"Discussion", 5:"Conclusion", 6:"Appendix"
            }
            for idx, name in mapping.items():
                with sec_tabs[idx]:
                    secs[name] = st.text_area(name, value=secs.get(name,""), height=260)

            with sec_tabs[7]:
                secs["授權/校內資源提醒"] = st.text_area("授權/校內資源提醒", value=secs.get("授權/校內資源提醒",""), height=140)

            # Full markdown
            with sec_tabs[8]:
                st.session_state["ms_full_md"] = sections_to_markdown(secs)
                st.code(st.session_state["ms_full_md"], language="markdown")
                st.download_button("下載稿件（Markdown）", data=st.session_state["ms_full_md"].encode("utf-8"),
                                   file_name="manuscript_draft_zhTW.md", mime="text/markdown")

            # Save back
            st.session_state["ms_sections"] = secs

    # -------------------- Diagnostics --------------------
    with tabs[7]:
        st.markdown("### Diagnostics（PubMed 被擋/限流時必看）")
        st.code(json.dumps(diag, ensure_ascii=False, indent=2), language="json")
        st.markdown(
            "- 若 `content_type` 顯示 `text/html` 或 warning 指向 non-JSON：可能被擋或限流。\n"
            "- 建議稍後重試或換網路環境。\n"
            "- PubMed count=0：多半是縮寫/自由詞擴充不足（例如只寫 EDOF），請改成更完整問題。\n"
        )
