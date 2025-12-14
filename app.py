# -*- coding: utf-8 -*-
from __future__ import annotations

import html
import io
import json
import math
import re
import time
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import pandas as pd
import requests
import streamlit as st

# ---------------------------
# Optional deps (degrade gracefully)
# ---------------------------
try:
    from PyPDF2 import PdfReader
    HAS_PYPDF2 = True
except Exception:
    HAS_PYPDF2 = False

try:
    import matplotlib.pyplot as plt
    HAS_MPL = True
except Exception:
    HAS_MPL = False

try:
    from docx import Document
    from docx.shared import Pt
    HAS_DOCX = True
except Exception:
    HAS_DOCX = False


# =========================
# Constants
# =========================
NCBI_ESEARCH = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
NCBI_EFETCH  = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"

EXPECTED_RECORD_COLS = [
    "record_id","pmid","pmcid","doi",
    "title","abstract","year","journal","first_author",
    "url","doi_url","pmc_url","source"
]

ROB2_DOMAINS = ["D1","D2","D3","D4","D5","Overall"]


# =========================
# i18n (zh-TW default)
# =========================
STR = {
    "zh-TW": {
        "app_title": "SR/MA 一句話原型（BYOK）",
        "app_caption": "輸入一句問題 → 自動產出：PICO/criteria/schema、PubMed 搜尋式、抓文獻、（可選）全文抽取/ROB2/MA、Word 草稿（若安裝）。",
        "lang": "語言",
        "lang_help": "可切換繁體中文 / English",
        "main_question": "研究問題（可一句話／可中文）",
        "q_placeholder": "例如：不同種類 EDOF IOL 在白內障手術後的視覺表現比較（Symfony vs Vivity vs AT LARA...）",
        "run": "開始執行（一步到位）",
        "advanced": "進階設定（可選）",
        "scope_pref": "範圍偏好",
        "scope_fast": "快速可行（找 gap、盡快成稿）",
        "scope_rigorous": "嚴謹完整（範圍較廣、較保守）",
        "require_CO": "搜尋時要求 C/O（可能降低召回）",
        "apply_rct": "套用 RCT filter（可能漏掉研究）",
        "max_records": "PubMed 最多抓幾筆",
        "polite_delay": "禮貌延遲（秒）",
        "feasibility_scan": "可行性：先掃描既有 SR/MA/NMA",
        "auto_pmc": "自動抓 PMC 開放取用全文（OA）",
        "auto_pmc_limit": "自動抓全文最多幾篇（僅 PMC OA）",
        "inst_access": "校內資源（可選）",
        "openurl": "OpenURL resolver base（可選）",
        "ezproxy": "EZproxy 前綴（可選）",
        "llm_block": "LLM（可選，使用者自備 API key）",
        "llm_toggle": "啟用 LLM 功能（用你自己的 API key）",
        "llm_notice": "Key only used for this session; do not use on untrusted deployments; do not upload identifiable patient info.",
        "base_url": "Base URL",
        "model": "模型名稱",
        "api_key": "API Key（不儲存、僅此 session）",
        "clear_key": "清除 API Key",
        "cleared": "已清除。",
        "server_side": "伺服器端呼叫（你輸入的 key 會送到本 app 伺服器執行 API 呼叫）",
        "no_key_degrade": "未啟用 LLM 或未提供 key：將自動降級（只做 PubMed/可行性/TA 初篩/匯出表格）。",
        "outputs": "輸出",
        "protocol": "Protocol（自動生成）",
        "pubmed_query": "PubMed 搜尋式（自動生成）",
        "diagnostics": "診斷資訊",
        "show_diag": "展開診斷",
        "srma_title": "可行性掃描：既有 SR/MA/NMA",
        "records": "文獻清單（含 TA 初篩）",
        "no_records": "沒有抓到文獻。通常是搜尋式太字面或太窄；請改成更像文獻會寫的句子（含場景/介入/比較）。",
        "download_csv": "下載 CSV",
        "extraction_wide": "抽取寬表（可編輯）",
        "rob2": "ROB 2.0（草稿／可編輯）",
        "rob2_note": "ROB 2.0 通常在全文納入後做。未啟用 LLM 或沒有全文時，這裡會空白。",
        "ma": "Meta-analysis（fixed effect；需 Effect/CI）",
        "ma_not_ready": "尚無法做 MA：需要 FT=Include for meta-analysis 且寬表中有 Effect + CI（可手填或由 LLM 抽取）。",
        "prisma": "PRISMA（草稿計數）",
        "export_word": "匯出 Word 草稿",
        "word_missing": "Word 匯出需要安裝 python-docx；目前環境未偵測到。",
        "done": "完成。請往下查看輸出。",
        "step0": "0/7 生成 protocol（PICO/criteria/schema/可行性建議）…",
        "step1": "1/7 生成 PubMed 搜尋式…",
        "step2": "2/7 可行性掃描（SR/MA/NMA）…",
        "step3": "3/7 抓 PubMed 文獻…",
        "step4": "4/7 Title/Abstract 初篩…",
        "step5": "5/7 自動抓 PMC OA 全文（可選）…",
        "step6": "6/7 全文抽取 + ROB2（需要 LLM + 全文）…",
        "step7": "7/7 嘗試 MA（fixed effect）…",
        "warn_llm_failed": "LLM 呼叫失敗，已自動降級（不會中斷流程）。",
        "privacy_tip": "提示：請勿上傳/貼上可識別病人資訊；校內訂閱全文請依規定使用（建議自行下載後再決定是否提供給 LLM）。",
        "upload_pdf": "上傳全文 PDF（可選；用於 extraction/ROB2）",
        "pdf_parse_fail": "PDF 解析失敗或無文字層，可能需要 OCR（請在 notes 標記 OCR REQUIRED）。",
        "manual_fulltext": "或貼上全文文字（可選）",
        "use_fulltext_from": "全文來源",
        "fulltext_from_none": "不提供全文（只做 TA 初篩）",
        "fulltext_from_pmc": "使用 PMC OA（若有抓到）",
        "fulltext_from_upload": "使用上傳 PDF / 貼上的全文",
        "ft_limit_hint": "全文越多、越貴、越慢；建議先小量測試。",
        "ft_include": "Include for meta-analysis",
        "ft_exclude": "Exclude after full-text",
        "ft_not": "Not reviewed yet",
        "ta_include": "Include",
        "ta_exclude": "Exclude",
        "ta_unsure": "Unsure",
    },
    "en": {
        "app_title": "SR/MA One-question prototype (BYOK)",
        "app_caption": "One question → PICO/criteria/schema, PubMed query, retrieval, optional full-text extraction/ROB2/MA, Word draft (if available).",
        "lang": "Language",
        "lang_help": "Switch zh-TW / English",
        "main_question": "Research question (one sentence; can be non-English)",
        "q_placeholder": "e.g., Compare different EDOF IOL types after cataract surgery (Symfony vs Vivity vs AT LARA...)",
        "run": "Run (one-click)",
        "advanced": "Advanced settings (optional)",
        "scope_pref": "Scope preference",
        "scope_fast": "Fast / feasible (gap-fill)",
        "scope_rigorous": "Rigorous / comprehensive",
        "require_CO": "Require C/O in search (may reduce recall)",
        "apply_rct": "Apply RCT filter (may miss studies)",
        "max_records": "Max PubMed records",
        "polite_delay": "Polite delay (sec)",
        "feasibility_scan": "Feasibility: scan existing SR/MA/NMA",
        "auto_pmc": "Auto-fetch PMC OA full text",
        "auto_pmc_limit": "Auto fulltext limit (PMC OA only)",
        "inst_access": "Institution access (optional)",
        "openurl": "OpenURL resolver base (optional)",
        "ezproxy": "EZproxy prefix (optional)",
        "llm_block": "LLM (optional, BYOK)",
        "llm_toggle": "Enable LLM features (use your own API key)",
        "llm_notice": "Key only used for this session; do not use on untrusted deployments; do not upload identifiable patient info.",
        "base_url": "Base URL",
        "model": "Model",
        "api_key": "API Key (session only; not stored)",
        "clear_key": "Clear API key",
        "cleared": "Cleared.",
        "server_side": "Server-side API call (your key is sent to this app server to call the API)",
        "no_key_degrade": "LLM disabled or missing key: auto-degrade (PubMed/feasibility/TA screening/exports only).",
        "outputs": "Outputs",
        "protocol": "Protocol (auto)",
        "pubmed_query": "PubMed query (auto)",
        "diagnostics": "Diagnostics",
        "show_diag": "Show diagnostics",
        "srma_title": "Feasibility: existing SR/MA/NMA scan",
        "records": "Records (with TA screening)",
        "no_records": "No records retrieved. Usually the query is too literal/narrow; rewrite the question with context/intervention/comparator.",
        "download_csv": "Download CSV",
        "extraction_wide": "Extraction wide table (editable)",
        "rob2": "ROB 2.0 (draft/editable)",
        "rob2_note": "ROB2 is usually after FT inclusion. Empty if LLM/full text not available.",
        "ma": "Meta-analysis (fixed effect; needs Effect/CI)",
        "ma_not_ready": "MA not ready: needs FT=Include for meta-analysis and Effect+CI in the wide table (manual or LLM).",
        "prisma": "PRISMA (draft counts)",
        "export_word": "Export Word draft",
        "word_missing": "Word export requires python-docx; not detected.",
        "done": "Done. Scroll down for outputs.",
        "step0": "0/7 Building protocol…",
        "step1": "1/7 Building PubMed query…",
        "step2": "2/7 Feasibility scan…",
        "step3": "3/7 Fetching PubMed…",
        "step4": "4/7 Title/Abstract screening…",
        "step5": "5/7 Fetching PMC OA full text…",
        "step6": "6/7 Full-text extraction + ROB2 (needs LLM + FT)…",
        "step7": "7/7 Meta-analysis attempt…",
        "warn_llm_failed": "LLM call failed; degraded without stopping the pipeline.",
        "privacy_tip": "Do not upload identifiable patient info. Handle subscription full text per your institution policies.",
        "upload_pdf": "Upload full-text PDF (optional; for extraction/ROB2)",
        "pdf_parse_fail": "PDF parse failed/no text layer; OCR may be required (mark OCR REQUIRED in notes).",
        "manual_fulltext": "Or paste full text (optional)",
        "use_fulltext_from": "Full-text source",
        "fulltext_from_none": "No full text (TA only)",
        "fulltext_from_pmc": "Use PMC OA (if fetched)",
        "fulltext_from_upload": "Use uploaded PDF / pasted text",
        "ft_limit_hint": "More full text = higher cost/time; start small.",
        "ft_include": "Include for meta-analysis",
        "ft_exclude": "Exclude after full-text",
        "ft_not": "Not reviewed yet",
        "ta_include": "Include",
        "ta_exclude": "Exclude",
        "ta_unsure": "Unsure",
    }
}


def t(key: str) -> str:
    lang = st.session_state.get("LANG", "zh-TW")
    return STR.get(lang, STR["zh-TW"]).get(key, key)


# =========================
# Utilities
# =========================
def norm_text(x: str) -> str:
    if x is None:
        return ""
    x = html.unescape(str(x))
    x = re.sub(r"\s+", " ", x).strip()
    return x

def ensure_cols(df: pd.DataFrame, cols: List[str], default="") -> pd.DataFrame:
    if df is None or not isinstance(df, pd.DataFrame):
        df = pd.DataFrame()
    for c in cols:
        if c not in df.columns:
            df[c] = default
    return df

def safe_empty_records_df() -> pd.DataFrame:
    df = pd.DataFrame(columns=EXPECTED_RECORD_COLS)
    return ensure_cols(df, EXPECTED_RECORD_COLS, "")

def to_csv_bytes(df: pd.DataFrame) -> bytes:
    df = df if isinstance(df, pd.DataFrame) else pd.DataFrame()
    return df.to_csv(index=False).encode("utf-8-sig")

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

def quote_if_needed(x: str) -> str:
    x = (x or "").strip()
    if not x:
        return ""
    if '"' in x:
        return x
    if re.search(r"\s|-|/|:", x):
        return f'"{x}"'
    return x

def has_advanced_syntax(term: str) -> bool:
    low = (term or "").lower()
    return any(tok in low for tok in [" or ", " and ", " not ", "[tiab]", "[mesh", "[mh]", "(", ")", '"', ":"])

def pubmed_url(pmid: str) -> str:
    pmid = (pmid or "").strip()
    return f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/" if pmid else ""

def doi_url(doi: str) -> str:
    doi = (doi or "").strip()
    return f"https://doi.org/{doi}" if doi else ""

def pmc_url(pmcid: str) -> str:
    pmcid = (pmcid or "").strip()
    if not pmcid:
        return ""
    if not pmcid.upper().startswith("PMC"):
        pmcid = "PMC" + pmcid
    return f"https://pmc.ncbi.nlm.nih.gov/articles/{pmcid}/"

def extract_english_tokens(text: str) -> List[str]:
    toks = re.findall(r"[A-Za-z][A-Za-z0-9\-\+]{1,}", text or "")
    toks = [t.strip() for t in toks if t.strip()]
    seen = set()
    out = []
    for tok in toks:
        u = tok.upper()
        if u not in seen:
            seen.add(u)
            out.append(tok)
    return out

def split_vs(question: str) -> Tuple[str, str, bool]:
    q = re.sub(r"\s+", " ", (question or "").strip())
    parts = re.split(r"\b(vs\.?|versus|compared to|compared with|compare to|comparison)\b", q, flags=re.I)
    if len(parts) >= 3:
        left = parts[0].strip(" -:;,.")
        right = " ".join(parts[2:]).strip(" -:;,.")
        return left, right, True
    return q, "", False


# =========================
# LLM (BYOK; no secrets)
# =========================
def llm_available() -> bool:
    ss = st.session_state
    if not ss.get("USE_LLM", False):
        return False
    return bool((ss.get("LLM_BASE_URL") or "").strip()
                and (ss.get("LLM_API_KEY") or "").strip()
                and (ss.get("LLM_MODEL") or "").strip())

def llm_chat(messages: List[dict], temperature: float = 0.2, timeout: int = 120) -> Optional[str]:
    if not llm_available():
        return None
    ss = st.session_state
    base = (ss.get("LLM_BASE_URL") or "").strip().rstrip("/")
    key  = (ss.get("LLM_API_KEY") or "").strip()
    model= (ss.get("LLM_MODEL") or "").strip()

    url = base + "/v1/chat/completions"
    headers = {"Authorization": f"Bearer {key}", "Content-Type": "application/json"}
    payload = {"model": model, "messages": messages, "temperature": float(temperature)}

    try:
        r = requests.post(url, headers=headers, json=payload, timeout=timeout)
        r.raise_for_status()
        js = r.json()
        return js["choices"][0]["message"]["content"]
    except Exception:
        st.warning(t("warn_llm_failed"))
        return None


# =========================
# PubMed (robust)
# =========================
def pubmed_esearch_ids(query: str, retstart: int, retmax: int) -> Tuple[List[str], int, str]:
    params = {"db": "pubmed", "term": query, "retmode": "json", "retstart": retstart, "retmax": retmax}
    r = requests.get(NCBI_ESEARCH, params=params, timeout=30)
    r.raise_for_status()
    js = r.json().get("esearchresult", {})
    ids = js.get("idlist", []) or []
    total = int(js.get("count", 0) or 0)
    return ids, total, r.url

def pubmed_efetch_xml(pmids: List[str]) -> Tuple[str, str]:
    params = {"db": "pubmed", "id": ",".join(pmids), "retmode": "xml"}
    r = requests.get(NCBI_EFETCH, params=params, timeout=90)
    r.raise_for_status()
    return (r.text or ""), r.url

def parse_pubmed_xml(xml_text: str) -> pd.DataFrame:
    if not xml_text or "<" not in xml_text:
        return safe_empty_records_df().iloc[:0].copy()

    head = xml_text[:300].lower()
    if "<html" in head or "access denied" in head or "cloudflare" in head:
        return safe_empty_records_df().iloc[:0].copy()

    import xml.etree.ElementTree as ET
    try:
        root = ET.fromstring(xml_text)
    except Exception:
        return safe_empty_records_df().iloc[:0].copy()

    rows = []
    for art in root.findall(".//PubmedArticle"):
        pmid = (art.findtext(".//PMID") or "").strip()
        title = norm_text(art.findtext(".//ArticleTitle") or "")

        ab_parts = []
        for ab in art.findall(".//AbstractText"):
            if ab.text:
                ab_parts.append(norm_text(ab.text))
        abstract = " ".join([x for x in ab_parts if x]).strip()

        year = (art.findtext(".//PubDate/Year") or "").strip()
        journal = norm_text(art.findtext(".//Journal/Title") or "")

        first_author = ""
        a0 = art.find(".//AuthorList/Author[1]")
        if a0 is not None:
            last = (a0.findtext("LastName") or "").strip()
            ini  = (a0.findtext("Initials") or "").strip()
            first_author = f"{last} {ini}".strip()

        doi = ""
        pmcid = ""
        for aid in art.findall(".//ArticleIdList/ArticleId"):
            idt = (aid.get("IdType") or "").lower()
            val = (aid.text or "").strip()
            if idt == "doi" and val and not doi:
                doi = val
            if idt == "pmc" and val and not pmcid:
                pmcid = val

        rows.append({
            "record_id": f"PMID:{pmid}" if pmid else "",
            "pmid": pmid,
            "pmcid": pmcid,
            "doi": doi,
            "title": title,
            "abstract": abstract,
            "year": year,
            "journal": journal,
            "first_author": first_author,
            "url": pubmed_url(pmid),
            "doi_url": doi_url(doi),
            "pmc_url": pmc_url(pmcid) if pmcid else "",
            "source": "PubMed"
        })

    df = pd.DataFrame(rows)
    df = ensure_cols(df, EXPECTED_RECORD_COLS, "")
    if not df.empty:
        mask = ~df["record_id"].astype(str).str.strip().astype(bool)
        df.loc[mask, "record_id"] = df.loc[mask, "pmid"].map(lambda x: f"PMID:{x}" if str(x).strip() else "")
    return df

def fetch_pubmed(query: str, max_records: int = 300, batch_size: int = 200, polite_delay: float = 0.0) -> Tuple[pd.DataFrame, int, Dict]:
    diag = {"esearch_url": "", "efetch_urls": [], "warnings": []}
    query = (query or "").strip()
    if not query:
        return safe_empty_records_df().iloc[:0].copy(), 0, {"warnings": ["Empty query"]}

    try:
        ids, total, esearch_url = pubmed_esearch_ids(query, 0, min(batch_size, 500))
        diag["esearch_url"] = esearch_url
    except Exception as e:
        diag["warnings"].append(f"Esearch failed: {e}")
        return safe_empty_records_df().iloc[:0].copy(), 0, diag

    target = min(total, max_records) if max_records and max_records > 0 else total
    all_ids = list(ids)

    while len(all_ids) < target:
        retstart = len(all_ids)
        need = min(batch_size, target - len(all_ids))
        try:
            ids2, _, _ = pubmed_esearch_ids(query, retstart, need)
        except Exception as e:
            diag["warnings"].append(f"Esearch page failed @retstart={retstart}: {e}")
            break
        if not ids2:
            break
        all_ids.extend(ids2)
        if polite_delay:
            time.sleep(polite_delay)

    frames = []
    for i in range(0, len(all_ids), batch_size):
        chunk = all_ids[i:i+batch_size]
        try:
            xml, efetch_url = pubmed_efetch_xml(chunk)
            diag["efetch_urls"].append(efetch_url)
            df_chunk = parse_pubmed_xml(xml)
            if df_chunk.empty and chunk:
                head = (xml or "")[:200].lower()
                if "<html" in head:
                    diag["warnings"].append("Efetch returned HTML (possibly blocked).")
            frames.append(df_chunk)
        except Exception as e:
            diag["warnings"].append(f"Efetch failed for batch starting {i}: {e}")
        if polite_delay:
            time.sleep(polite_delay)

    if not frames:
        return safe_empty_records_df().iloc[:0].copy(), total, diag

    df = pd.concat(frames, ignore_index=True)
    df = ensure_cols(df, EXPECTED_RECORD_COLS, "")

    # de-dup robustly
    if not df.empty:
        df["title_norm"] = df["title"].fillna("").str.lower().str.replace(r"\s+"," ", regex=True).str.strip()
        df["doi_norm"] = df["doi"].fillna("").str.lower().str.strip()
        if df["doi_norm"].astype(bool).any():
            df = df.sort_values(["doi_norm","year"]).drop_duplicates(subset=["doi_norm"], keep="first")
        df = df.sort_values(["title_norm","year"]).drop_duplicates(subset=["title_norm","year"], keep="first")
        df = df.drop(columns=["title_norm","doi_norm"], errors="ignore").reset_index(drop=True)

        mask = ~df["record_id"].astype(str).str.strip().astype(bool)
        df.loc[mask, "record_id"] = df.loc[mask, "pmid"].map(lambda x: f"PMID:{x}" if str(x).strip() else "")

    return df, total, diag


# =========================
# PMC OA full text (optional)
# =========================
def fetch_pmc_fulltext_xml(pmcid: str) -> str:
    pmcid = (pmcid or "").strip()
    if not pmcid:
        return ""
    if not pmcid.upper().startswith("PMC"):
        pmcid = "PMC" + pmcid
    params = {"db": "pmc", "id": pmcid, "retmode": "xml"}
    r = requests.get(NCBI_EFETCH, params=params, timeout=90)
    r.raise_for_status()
    return r.text or ""

def pmc_xml_to_text(xml_text: str, max_chars: int = 180000) -> str:
    if not xml_text:
        return ""
    import xml.etree.ElementTree as ET
    try:
        root = ET.fromstring(xml_text)
    except Exception:
        return ""
    texts: List[str] = []
    for tag in ["article-title", "abstract", "p", "table-wrap", "fig"]:
        for el in root.findall(f".//{tag}"):
            t2 = norm_text("".join(el.itertext()))
            if t2:
                texts.append(t2)
            if sum(len(x) for x in texts) > max_chars:
                break
    return ("\n".join(texts))[:max_chars]


# =========================
# Protocol (LLM optional)
# =========================
def protocol_fallback(question: str, goal_mode: str) -> dict:
    left, right, has_vs = split_vs(question)
    pico = {
        "P": "",
        "I": left if left else question,
        "C": right if has_vs else "",
        "O": "",
        "NOT": "animal OR mice OR rat OR in vitro OR case report"
    }

    # If question contains lots of CJK, try to keep useful English tokens in I_synonyms
    toks = extract_english_tokens(question)
    I_syn = []
    if re.search(r"[\u4e00-\u9fff]", pico["I"]) and toks:
        I_syn = toks

    plan = {
        "principles": [
            "不要把 extraction 欄位寫死：在 PICO 層級由模型規劃 extraction sheet",
            "先檢查是否已有相關 SR/MA/NMA；若存在，對齊並延伸（包含新 RCT 或不同比較）",
            "納入所有 RCT primary + secondary outcomes（避免只抽你想要的那幾個）",
            "若資料在 figure/table 或 PDF 無文字層：標記 OCR REQUIRED，指明要 OCR 的頁/圖/表"
        ],
        "base_cols": ["First author","Year","Design","Population details","Intervention","Comparator","Follow-up","Notes (figure/table/section)"],
        "outcome_groups": [
            {"group_name":"Primary outcomes","suggested_items":["Primary outcome"]},
            {"group_name":"Secondary outcomes","suggested_items":["Secondary outcome 1","Secondary outcome 2"]}
        ],
        "effect_preference": ["RR","OR","HR","MD","SMD"]
    }

    return {
        "pico": pico,
        "inclusion_decision": {
            "scope_tradeoff": "gap_fill_fast" if goal_mode == "fast" else "rigorous_scope",
            "recommended_scope": "（降級模式）未啟用 LLM：建議用更具體的一句話（包含場景/介入/比較/結局）以提高搜尋召回與精準度。"
        },
        "inclusion_criteria": ["Human studies relevant to the question.", "Prefer RCTs if meta-analysis intended."],
        "exclusion_criteria": ["Animal/in vitro", "Case report/series only (unless required)"],
        "search_expansion": {
            "P_synonyms": [],
            "I_synonyms": I_syn,
            "C_synonyms": [],
            "O_synonyms": [],
            "NOT": ["animal","mice","rat","in vitro","case report"]
        },
        "mesh_candidates": {"P":[],"I":[],"C":[],"O":[]},
        "feasibility_plan": {
            "how_to_check_existing_srma": "Use PubMed filters for systematic review/meta-analysis/network meta-analysis and check overlap.",
            "what_to_do_if_srma_exists": "Narrow PICO boundaries (population/comparator/time), include newer RCTs, or switch focus to feasibility-high gap."
        },
        "recommended_extraction_schema_plan": plan,
        "analysis_plan": {
            "effect_measures": ["RR","OR","HR","MD","SMD"],
            "timepoint_preference": ["final follow-up"],
            "consider_nma": "maybe",
            "nma_node_definition": "Define nodes by intervention types/brands; ensure transitivity and connected network."
        }
    }

def protocol_from_question_llm(question: str, goal_mode: str) -> dict:
    sys = "You are a senior SR/MA methodologist. Output ONLY valid JSON."
    user = {
        "task": "From one question, produce protocol + inclusion criteria decision + search expansion + extraction schema plan (not hard-coded) + NMA considerations.",
        "goal_mode": "Fast / feasible (gap-fill)" if goal_mode == "fast" else "Rigorous / comprehensive",
        "question": question,
        "output_schema": {
            "pico": {"P":"","I":"","C":"","O":"","NOT":""},
            "inclusion_decision": {
                "scope_tradeoff": "gap_fill_fast OR rigorous_scope",
                "recommended_scope": "How to set PICO boundaries given existing evidence and feasibility."
            },
            "inclusion_criteria": ["..."],
            "exclusion_criteria": ["..."],
            "search_expansion": {
                "P_synonyms": ["..."], "I_synonyms": ["..."], "C_synonyms": ["..."], "O_synonyms": ["..."],
                "NOT": ["animal","case report","in vitro"]
            },
            "mesh_candidates": {"P":["..."],"I":["..."],"C":["..."],"O":["..."]},
            "feasibility_plan": {
                "how_to_check_existing_srma": "...",
                "what_to_do_if_srma_exists": "..."
            },
            "recommended_extraction_schema_plan": {
                "principles": [
                    "Do not hard-code columns; define at PICO level",
                    "Check whether prior SR/MA/NMA exists and align/extend",
                    "Enumerate ALL RCT primary and secondary outcomes for extraction",
                    "Include OCR/figure/table instructions"
                ],
                "base_cols": ["First author","Year","Design","Population details","Intervention","Comparator","Follow-up","Notes (figure/table/section)"],
                "outcome_groups": [
                    {"group_name":"Primary outcomes","suggested_items":["..."]},
                    {"group_name":"Secondary outcomes","suggested_items":["..."]}
                ],
                "effect_preference": ["RR","OR","HR","MD","SMD"]
            },
            "analysis_plan": {
                "effect_measures": ["RR","OR","HR","MD","SMD"],
                "timepoint_preference": ["final follow-up"],
                "consider_nma": "yes/no/maybe",
                "nma_node_definition": "how to define nodes for NMA"
            }
        },
        "constraints": [
            "Topic-agnostic; do not assume a specialty.",
            "Maximize recall with abbreviations/variant spellings.",
            "If question is non-English, translate/expand to English search terms.",
            "Avoid hallucinating; mark uncertainty explicitly."
        ]
    }
    txt = llm_chat(
        [{"role":"system","content":sys},{"role":"user","content":json.dumps(user, ensure_ascii=False)}],
        temperature=0.2,
        timeout=180
    )
    js = json_from_text(txt or "")
    if not js:
        return {"error":"bad_json","raw":(txt or "")}
    return js


# =========================
# Query builder
# =========================
def build_concept_group(term: str, mesh_list: List[str], synonyms: List[str], field_tag: str = "tiab") -> str:
    term = (term or "").strip()
    synonyms = [s.strip() for s in (synonyms or []) if (s or "").strip()]
    mesh_list = [m.strip() for m in (mesh_list or []) if (m or "").strip()]

    free_blocks = []
    if term:
        if has_advanced_syntax(term):
            free_blocks.append(f"({term})")
        else:
            free_blocks.append(f"{quote_if_needed(term)}[{field_tag}]")
    for s in synonyms:
        if has_advanced_syntax(s):
            free_blocks.append(s)
        else:
            free_blocks.append(f"{quote_if_needed(s)}[{field_tag}]")

    mesh_blocks = [f'{quote_if_needed(m)}[MeSH Terms]' for m in mesh_list[:8]]
    blocks = []
    if free_blocks:
        blocks.append("(" + " OR ".join(free_blocks) + ")")
    if mesh_blocks:
        blocks.append("(" + " OR ".join(mesh_blocks) + ")")
    if not blocks:
        return ""
    return "(" + " OR ".join(blocks) + ")" if len(blocks) > 1 else blocks[0]

def build_pubmed_query(protocol: dict, strict_CO: bool, include_rct_filter: bool) -> str:
    pico = protocol.get("pico", {}) or {}
    exp  = protocol.get("search_expansion", {}) or {}
    mesh = protocol.get("mesh_candidates", {}) or {}

    P = build_concept_group(pico.get("P",""), mesh.get("P", []), exp.get("P_synonyms", []))
    I = build_concept_group(pico.get("I",""), mesh.get("I", []), exp.get("I_synonyms", []))
    C = build_concept_group(pico.get("C",""), mesh.get("C", []), exp.get("C_synonyms", []))
    O = build_concept_group(pico.get("O",""), mesh.get("O", []), exp.get("O_synonyms", []))

    # Recall-first: use I (and C if present), even if P empty
    parts = []
    if P: parts.append(P)
    if I: parts.append(I)
    if C: parts.append(C)
    if strict_CO and O:
        parts.append(O)

    if not parts:
        return f'{quote_if_needed("health")}[tiab]'

    base = "(" + " AND ".join(parts) + ")"
    if include_rct_filter:
        rct = "(randomized controlled trial[pt] OR randomized[tiab] OR randomised[tiab] OR trial[tiab] OR placebo[tiab])"
        base = base + " AND " + rct

    NOT = (pico.get("NOT","") or "").strip()
    if not NOT:
        NOT = " OR ".join([x for x in (exp.get("NOT", []) or []) if x])
    if NOT:
        base = base + " NOT (" + NOT + ")"
    return base


# =========================
# Feasibility scan: SR/MA/NMA
# =========================
def scan_sr_ma_nma(pubmed_query: str, top_n: int = 25) -> Dict:
    q = (pubmed_query or "").strip()
    if not q:
        return {"summary":"(empty query)","hits":pd.DataFrame()}
    sr_filter = (
        '(systematic review[pt] OR meta-analysis[pt] '
        'OR "systematic review"[tiab] OR "meta analysis"[tiab] OR meta-analy*[tiab] '
        'OR "network meta-analysis"[tiab] OR "network meta analysis"[tiab] OR NMA[tiab] '
        'OR "mixed treatment comparison"[tiab] OR "indirect comparison"[tiab])'
    )
    q2 = "(" + q + ") AND " + sr_filter
    try:
        ids, count, _ = pubmed_esearch_ids(q2, 0, min(top_n, 50))
        hits = pd.DataFrame()
        if ids:
            xml, _ = pubmed_efetch_xml(ids)
            hits = parse_pubmed_xml(xml)[["pmid","doi","pmcid","title","year","first_author","journal","url","doi_url","pmc_url"]]
        return {"summary": f"PubMed count≈{count} (showing top {min(len(ids), top_n)}).", "hits": hits}
    except Exception as e:
        return {"summary": f"Scan failed: {e}", "hits": pd.DataFrame()}


# =========================
# Title/Abstract screening
# =========================
def screen_rule_based(row: pd.Series, protocol: dict) -> Dict:
    title = (row.get("title","") or "")
    abstract = (row.get("abstract","") or "")
    text = (title + " " + abstract).lower()

    pico = protocol.get("pico", {}) or {}
    I = (pico.get("I","") or "").lower()
    C = (pico.get("C","") or "").lower()
    NOT = (pico.get("NOT","") or "").lower()

    def hit(term: str) -> bool:
        term = (term or "").strip().lower()
        if not term:
            return False
        toks = [t for t in re.split(r"[,;\s]+", term) if t][:10]
        return any(t in text for t in toks)

    if NOT and hit(NOT):
        return {"label": t("ta_exclude"), "confidence": 0.8, "reason": "NOT keyword hit"}
    if I and hit(I) and (not C or hit(C)):
        return {"label": t("ta_include"), "confidence": 0.7, "reason": "Key terms present"}
    if any(w in text for w in ["randomized","randomised","trial","controlled"]) and hit(I):
        return {"label": t("ta_include"), "confidence": 0.65, "reason": "Trial-like + match"}
    return {"label": t("ta_unsure"), "confidence": 0.4, "reason": "Insufficient TA signal"}

def screen_llm(row: pd.Series, protocol: dict) -> Dict:
    sys = "You are an SR/MA screening assistant. Output ONLY valid JSON."
    user = {
        "task":"Title/Abstract screening",
        "pico": protocol.get("pico", {}) or {},
        "inclusion_criteria": protocol.get("inclusion_criteria", []) or [],
        "exclusion_criteria": protocol.get("exclusion_criteria", []) or [],
        "record": {"title": row.get("title",""), "abstract": row.get("abstract",""), "year": row.get("year",""), "journal": row.get("journal","")},
        "output_schema": {"label":"Include/Exclude/Unsure","confidence":"0..1","reason":"short"}
    }
    txt = llm_chat(
        [{"role":"system","content":sys},{"role":"user","content":json.dumps(user, ensure_ascii=False)}],
        temperature=0.1,
        timeout=90
    )
    js = json_from_text(txt or "")
    if not js:
        return screen_rule_based(row, protocol)
    lab = str(js.get("label","Unsure")).strip()
    # normalize to UI language labels
    if lab.lower().startswith("incl"):
        lab2 = t("ta_include")
    elif lab.lower().startswith("excl"):
        lab2 = t("ta_exclude")
    else:
        lab2 = t("ta_unsure")
    conf = js.get("confidence", 0.0)
    try:
        conf = float(conf)
    except Exception:
        conf = 0.0
    return {"label": lab2, "confidence": conf, "reason": str(js.get("reason",""))}


# =========================
# Full-text ingestion (upload or paste)
# =========================
def pdf_to_text(file_bytes: bytes, max_chars: int = 180000) -> str:
    if not HAS_PYPDF2:
        return ""
    try:
        reader = PdfReader(io.BytesIO(file_bytes))
        chunks = []
        for p in reader.pages[:50]:
            try:
                chunks.append(p.extract_text() or "")
            except Exception:
                continue
        txt = "\n".join(chunks)
        txt = re.sub(r"\s+", " ", txt).strip()
        return txt[:max_chars]
    except Exception:
        return ""

def choose_fulltext_source(record_id: str, records: pd.DataFrame, pmc_text_map: dict, uploaded_text: str) -> str:
    # Prefer uploaded/pasted (user explicit) over PMC (implicit), if provided
    if uploaded_text and uploaded_text.strip():
        return uploaded_text
    if pmc_text_map.get(record_id):
        return pmc_text_map[record_id]
    return ""


# =========================
# Extraction + ROB2 prompts (LLM)
# =========================
def build_extraction_prompt(protocol: dict) -> str:
    pico = protocol.get("pico", {}) or {}
    plan = protocol.get("recommended_extraction_schema_plan", {}) or {}
    base_cols = plan.get("base_cols", []) or []
    outcome_groups = plan.get("outcome_groups", []) or []

    criteria_text = "Inclusion:\n" + "\n".join([f"- {x}" for x in (protocol.get("inclusion_criteria", []) or [])])
    criteria_text += "\n\nExclusion:\n" + "\n".join([f"- {x}" for x in (protocol.get("exclusion_criteria", []) or [])])

    return f"""
You are performing systematic review FULL-TEXT data extraction.

Critical OCR / Figure / Table instructions:
- If PDF is scanned or text is missing/garbled: write 'OCR REQUIRED' in notes and specify what to OCR (pages/sections).
- Prefer extracting from outcome TABLES, figure legends, supplements.
- If a value is taken from a figure: write 'from Figure X (approx.)' and capture axis units/timepoint.
- NEVER fabricate numbers. If not found, leave blank and state where it might be.

PICO:
P={pico.get("P","")}
I={pico.get("I","")}
C={pico.get("C","")}
O={pico.get("O","")}
NOT={pico.get("NOT","")}

Criteria:
{criteria_text}

Extraction schema planning requirement:
- Do NOT hard-code columns. Follow this plan:
  principles={json.dumps(plan.get("principles", []), ensure_ascii=False)}
  base_cols={json.dumps(base_cols, ensure_ascii=False)}
  outcome_groups={json.dumps(outcome_groups, ensure_ascii=False)}
- Explicitly check/mention whether prior SR/MA exists (if referenced) and align/extend.
- Enumerate ALL RCT primary + secondary outcomes that are relevant.

Effect estimates (if available):
- effect_measure in {plan.get("effect_preference", ["RR","OR","HR","MD","SMD"])}
- effect, lower_CI, upper_CI, timepoint, unit

Output JSON ONLY with keys:
- fulltext_decision: "Include for meta-analysis" / "Exclude after full-text" / "Not reviewed yet"
- fulltext_reason
- extraction_schema: object
- extracted_fields: object
- meta: object (effect_measure,effect,lower_CI,upper_CI,timepoint,unit)
- notes: list of strings
""".strip()

def extract_llm(fulltext: str, protocol: dict) -> dict:
    sys = "You are an SR/MA full-text reviewer and extractor. Output ONLY valid JSON."
    prompt = build_extraction_prompt(protocol)
    txt = llm_chat(
        [{"role":"system","content":sys},
         {"role":"user","content":prompt + "\n\n[Full text]\n" + fulltext[:120000]}],
        temperature=0.1,
        timeout=180
    )
    js = json_from_text(txt or "")
    if not js:
        return {"error":"bad_json","raw":(txt or "")}
    return js

def build_rob2_prompt(protocol: dict) -> str:
    pico = protocol.get("pico", {}) or {}
    return f"""
You are applying Cochrane ROB 2.0 for randomized trials.

Rules:
- Do NOT guess. If insufficient info, choose 'Some concerns' and say what is missing.
- If OCR seems required, add 'OCR REQUIRED' in notes.
- Provide evidence pointers (Methods/Randomization, Table X, Figure Y).

Topic context:
P={pico.get("P","")}
I={pico.get("I","")}
C={pico.get("C","")}
O={pico.get("O","")}

Output JSON only:
{{
  "D1": {{"judgement":"Low risk/Some concerns/High risk","rationale":"...","evidence":["..."]}},
  "D2": {{"judgement":"Low risk/Some concerns/High risk","rationale":"...","evidence":["..."]}},
  "D3": {{"judgement":"Low risk/Some concerns/High risk","rationale":"...","evidence":["..."]}},
  "D4": {{"judgement":"Low risk/Some concerns/High risk","rationale":"...","evidence":["..."]}},
  "D5": {{"judgement":"Low risk/Some concerns/High risk","rationale":"...","evidence":["..."]}},
  "Overall": {{"judgement":"Low risk/Some concerns/High risk","rationale":"...","evidence":["..."]}},
  "notes": ["..."]
}}
""".strip()

def rob2_llm(fulltext: str, protocol: dict) -> dict:
    sys = "You are a risk-of-bias assessor. Output ONLY valid JSON."
    prompt = build_rob2_prompt(protocol)
    txt = llm_chat(
        [{"role":"system","content":sys},
         {"role":"user","content":prompt + "\n\n[Full text]\n" + fulltext[:120000]}],
        temperature=0.1,
        timeout=180
    )
    js = json_from_text(txt or "")
    if not js:
        return {"error":"bad_json","raw":(txt or "")}
    return js


# =========================
# Fixed-effect MA (needs Effect + CI)
# =========================
def _to_float(x) -> Optional[float]:
    try:
        if x is None:
            return None
        s = str(x).strip().replace("−","-")
        if not s or s.lower() in ["nan","none"]:
            return None
        return float(s)
    except Exception:
        return None

def fixed_effect_ma(df: pd.DataFrame) -> dict:
    if df is None or df.empty:
        return {"error":"no data"}

    rows = []
    for _, r in df.iterrows():
        m = (r.get("Effect_measure","") or "").strip().upper()
        eff = _to_float(r.get("Effect"))
        lo  = _to_float(r.get("Lower_CI"))
        hi  = _to_float(r.get("Upper_CI"))
        if eff is None or lo is None or hi is None:
            continue

        is_log = m in ["OR","RR","HR"]
        if is_log:
            if eff <= 0 or lo <= 0 or hi <= 0:
                continue
            y = math.log(eff)
            se = (math.log(hi) - math.log(lo)) / (2 * 1.96)
        else:
            y = eff
            se = (hi - lo) / (2 * 1.96)

        if se <= 0:
            continue
        w = 1.0 / (se * se)
        rows.append((y, se, w, m, eff, lo, hi, r.get("title",""), r.get("record_id","")))

    if not rows:
        return {"error":"insufficient numeric effect/CI"}

    W = sum(w for _,_,w,_,_,_,_,_,_ in rows)
    yhat = sum(y*w for y,_,w,_,_,_,_,_,_ in rows) / W
    sehat = math.sqrt(1.0 / W)
    lohat = yhat - 1.96 * sehat
    hihat = yhat + 1.96 * sehat

    m0 = rows[0][3] or "Effect"
    if m0 in ["OR","RR","HR"]:
        pooled, lo_p, hi_p = math.exp(yhat), math.exp(lohat), math.exp(hihat)
    else:
        pooled, lo_p, hi_p = yhat, lohat, hihat

    return {"k": len(rows), "measure": m0, "pooled": pooled, "lower": lo_p, "upper": hi_p, "rows": rows}

def plot_forest(ma: dict):
    if not HAS_MPL:
        return None
    rows = ma.get("rows", [])
    if not rows:
        return None
    m0 = ma.get("measure","Effect")
    is_log = m0 in ["OR","RR","HR"]

    labels, effects, lowers, uppers = [], [], [], []
    for _,_,_,_,eff,lo,hi,title,rid in rows:
        labels.append((title or rid or "")[:60])
        effects.append(eff); lowers.append(lo); uppers.append(hi)

    fig, ax = plt.subplots(figsize=(8, max(3.0, 0.35*len(rows)+1.8)))
    y_pos = list(range(len(rows)))[::-1]
    ax.errorbar(
        effects, y_pos,
        xerr=[[e-l for e,l in zip(effects,lowers)], [u-e for e,u in zip(effects,uppers)]],
        fmt="o", capsize=3
    )
    ax.set_yticks(y_pos); ax.set_yticklabels(labels)
    ax.axvline(1.0 if is_log else 0.0, linestyle="--")
    ax.set_xlabel(m0)
    ax.set_title(f"Forest plot (fixed effect), k={ma.get('k',0)}")
    ax.grid(True, axis="x", linestyle=":", linewidth=0.5)
    fig.tight_layout()
    return fig


# =========================
# Word export (optional)
# =========================
def export_docx(question: str, protocol: dict, pubmed_query: str, prisma: dict,
                wide: pd.DataFrame, rob_df: pd.DataFrame, ma: Optional[dict]) -> Optional[bytes]:
    if not HAS_DOCX:
        return None
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    doc.add_heading("SR/MA Draft (Auto)", level=1)
    doc.add_paragraph("Research question:")
    doc.add_paragraph(question)

    doc.add_heading("Protocol (auto)", level=2)
    doc.add_paragraph(json.dumps(protocol, ensure_ascii=False, indent=2))

    doc.add_heading("PubMed query", level=2)
    doc.add_paragraph(pubmed_query)

    doc.add_heading("PRISMA (draft counts)", level=2)
    for k, v in prisma.items():
        doc.add_paragraph(f"- {k}: {v}")

    doc.add_heading("Extraction (wide snapshot)", level=2)
    if isinstance(wide, pd.DataFrame) and not wide.empty:
        cols = list(wide.columns)[:min(10, len(wide.columns))]
        t2 = doc.add_table(rows=1, cols=len(cols))
        for i, c in enumerate(cols):
            t2.rows[0].cells[i].text = c
        for _, r in wide.head(25).iterrows():
            row = t2.add_row().cells
            for i, c in enumerate(cols):
                row[i].text = str(r.get(c, ""))
    else:
        doc.add_paragraph("(none)")

    doc.add_heading("ROB 2.0 (snapshot)", level=2)
    if isinstance(rob_df, pd.DataFrame) and not rob_df.empty:
        t3 = doc.add_table(rows=1, cols=len(rob_df.columns))
        for i, c in enumerate(rob_df.columns):
            t3.rows[0].cells[i].text = c
        for _, r in rob_df.iterrows():
            row = t3.add_row().cells
            for i, c in enumerate(rob_df.columns):
                row[i].text = str(r.get(c, ""))
    else:
        doc.add_paragraph("(none)")

    doc.add_heading("Meta-analysis (fixed effect)", level=2)
    if ma and not ma.get("error"):
        doc.add_paragraph(f"k={ma.get('k')} measure={ma.get('measure')}")
        doc.add_paragraph(f"Pooled: {ma.get('pooled'):.4g} (95% CI {ma.get('lower'):.4g} to {ma.get('upper'):.4g})")
    else:
        doc.add_paragraph("(Not available)")

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# =========================
# Session init
# =========================
def init_state():
    ss = st.session_state
    ss.setdefault("LANG", "zh-TW")

    ss.setdefault("question", "")
    ss.setdefault("goal_mode", "fast")  # fast/rigorous
    ss.setdefault("strict_CO", False)
    ss.setdefault("include_rct_filter", False)
    ss.setdefault("include_srma_scan", True)
    ss.setdefault("max_records", 300)
    ss.setdefault("polite_delay", 0.0)

    ss.setdefault("auto_fetch_pmc", True)
    ss.setdefault("auto_fulltext_limit", 8)

    ss.setdefault("USE_LLM", False)
    ss.setdefault("LLM_BASE_URL", "https://api.openai.com")
    ss.setdefault("LLM_MODEL", "gpt-4o-mini")
    ss.setdefault("LLM_API_KEY", "")

    ss.setdefault("protocol", None)
    ss.setdefault("pubmed_query", "")
    ss.setdefault("srma_scan", {"summary":"", "hits": pd.DataFrame()})
    ss.setdefault("records", safe_empty_records_df().iloc[:0].copy())
    ss.setdefault("pubmed_total", 0)
    ss.setdefault("diag", {})

    ss.setdefault("screening", pd.DataFrame())
    ss.setdefault("ta_final", {})       # record_id -> Include/Exclude/Unsure
    ss.setdefault("ft_decision", {})    # record_id -> Not reviewed / Include for MA / Exclude
    ss.setdefault("ft_reason", {})      # record_id -> reason
    ss.setdefault("pmc_fulltext", {})   # record_id -> text

    ss.setdefault("uploaded_fulltext", "")  # global text from upload/paste
    ss.setdefault("fulltext_source", "none")  # none/pmc/upload

    ss.setdefault("extraction_wide", pd.DataFrame())
    ss.setdefault("rob2_table", {})
    ss.setdefault("rob2_raw", {})
    ss.setdefault("ma_result", None)

init_state()


# =========================
# UI
# =========================
st.set_page_config(page_title=t("app_title"), layout="wide")
st.title(t("app_title"))
st.caption(t("app_caption"))
st.info(t("privacy_tip"))

# Language toggle
with st.sidebar:
    st.session_state["LANG"] = st.selectbox(
        t("lang"),
        ["zh-TW", "en"],
        index=0 if st.session_state.get("LANG","zh-TW") == "zh-TW" else 1,
        help=t("lang_help")
    )

# Advanced settings
with st.sidebar:
    with st.expander(t("advanced"), expanded=False):
        goal = st.selectbox(
            t("scope_pref"),
            [t("scope_fast"), t("scope_rigorous")],
            index=0 if st.session_state["goal_mode"] == "fast" else 1
        )
        st.session_state["goal_mode"] = "fast" if goal == t("scope_fast") else "rigorous"

        st.session_state["strict_CO"] = st.checkbox(t("require_CO"), value=st.session_state["strict_CO"])
        st.session_state["include_rct_filter"] = st.checkbox(t("apply_rct"), value=st.session_state["include_rct_filter"])
        st.session_state["include_srma_scan"] = st.checkbox(t("feasibility_scan"), value=st.session_state["include_srma_scan"])
        st.session_state["max_records"] = st.number_input(t("max_records"), 50, 5000, int(st.session_state["max_records"]), 50)
        st.session_state["polite_delay"] = st.slider(t("polite_delay"), 0.0, 1.0, float(st.session_state["polite_delay"]), 0.1)

        st.session_state["auto_fetch_pmc"] = st.checkbox(t("auto_pmc"), value=st.session_state["auto_fetch_pmc"])
        st.session_state["auto_fulltext_limit"] = st.number_input(t("auto_pmc_limit"), 0, 50, int(st.session_state["auto_fulltext_limit"]), 1)
        st.caption(t("ft_limit_hint"))

# BYOK LLM block
with st.sidebar:
    st.markdown("---")
    st.subheader(t("llm_block"))
    st.session_state["USE_LLM"] = st.checkbox(t("llm_toggle"), value=bool(st.session_state.get("USE_LLM", False)))
    st.info(t("llm_notice"))

    if st.session_state["USE_LLM"]:
        st.session_state["LLM_BASE_URL"] = st.text_input(t("base_url"), value=st.session_state.get("LLM_BASE_URL", "https://api.openai.com"))
        st.session_state["LLM_MODEL"] = st.text_input(t("model"), value=st.session_state.get("LLM_MODEL", "gpt-4o-mini"))
        st.session_state["LLM_API_KEY"] = st.text_input(t("api_key"), value=st.session_state.get("LLM_API_KEY",""), type="password")
        c1, c2 = st.columns(2)
        with c1:
            if st.button(t("clear_key")):
                st.session_state["LLM_API_KEY"] = ""
                st.success(t("cleared"))
        with c2:
            st.caption(t("server_side"))
    else:
        st.session_state["LLM_API_KEY"] = ""
        st.caption(t("no_key_degrade"))

# Full-text inputs (optional)
st.subheader(t("use_fulltext_from"))
ft_source = st.radio(
    t("use_fulltext_from"),
    options=[t("fulltext_from_none"), t("fulltext_from_pmc"), t("fulltext_from_upload")],
    index=0,
    horizontal=True
)
if ft_source == t("fulltext_from_none"):
    st.session_state["fulltext_source"] = "none"
elif ft_source == t("fulltext_from_pmc"):
    st.session_state["fulltext_source"] = "pmc"
else:
    st.session_state["fulltext_source"] = "upload"

if st.session_state["fulltext_source"] == "upload":
    up = st.file_uploader(t("upload_pdf"), type=["pdf"])
    pasted = st.text_area(t("manual_fulltext"), value="", height=160)
    fulltxt = ""

    if up is not None:
        file_bytes = up.read()
        fulltxt = pdf_to_text(file_bytes)
        if not fulltxt:
            st.warning(t("pdf_parse_fail"))
    if pasted.strip():
        fulltxt = pasted.strip()

    st.session_state["uploaded_fulltext"] = fulltxt

# Main question + Run
st.session_state["question"] = st.text_area(
    t("main_question"),
    value=st.session_state["question"],
    height=90,
    placeholder=t("q_placeholder")
)
run = st.button(t("run"), type="primary")


# =========================
# Pipeline run
# =========================
if run:
    q = (st.session_state["question"] or "").strip()
    if not q:
        st.error("Please enter a question.")
        st.stop()

    # Step 0: protocol
    with st.spinner(t("step0")):
        if llm_available():
            prot = protocol_from_question_llm(q, st.session_state["goal_mode"])
            if prot.get("error"):
                prot = protocol_fallback(q, st.session_state["goal_mode"])
        else:
            prot = protocol_fallback(q, st.session_state["goal_mode"])
        st.session_state["protocol"] = prot

    # Step 1: query
    with st.spinner(t("step1")):
        st.session_state["pubmed_query"] = build_pubmed_query(
            st.session_state["protocol"],
            st.session_state["strict_CO"],
            st.session_state["include_rct_filter"]
        )

    # Step 2: feasibility scan
    if st.session_state["include_srma_scan"]:
        with st.spinner(t("step2")):
            st.session_state["srma_scan"] = scan_sr_ma_nma(st.session_state["pubmed_query"], top_n=25)

    # Step 3: PubMed retrieval
    with st.spinner(t("step3")):
        df, total, diag = fetch_pubmed(
            st.session_state["pubmed_query"],
            max_records=int(st.session_state["max_records"]),
            batch_size=200,
            polite_delay=float(st.session_state["polite_delay"])
        )
        df = ensure_cols(df, EXPECTED_RECORD_COLS, "")
        st.session_state["records"] = df
        st.session_state["pubmed_total"] = int(total or 0)
        st.session_state["diag"] = diag

        for rid in df["record_id"].tolist():
            st.session_state["ta_final"].setdefault(rid, t("ta_unsure"))
            st.session_state["ft_decision"].setdefault(rid, t("ft_not"))
            st.session_state["ft_reason"].setdefault(rid, "")
            st.session_state["pmc_fulltext"].setdefault(rid, "")

    # Step 4: TA screening
    with st.spinner(t("step4")):
        out = []
        for _, r in st.session_state["records"].iterrows():
            res = screen_llm(r, st.session_state["protocol"]) if llm_available() else screen_rule_based(r, st.session_state["protocol"])
            rid = r["record_id"]
            st.session_state["ta_final"][rid] = res["label"]
            out.append({"record_id": rid, "AI_label": res["label"], "AI_confidence": res["confidence"], "AI_reason": res["reason"]})
        st.session_state["screening"] = pd.DataFrame(out)

    # Step 5: Auto-fetch PMC OA full text
    if st.session_state["auto_fetch_pmc"]:
        with st.spinner(t("step5")):
            merged = st.session_state["records"].merge(st.session_state["screening"], on="record_id", how="left")
            merged = ensure_cols(merged, ["AI_label"], "")
            cand = merged[merged["AI_label"].isin([t("ta_include"), t("ta_unsure")])].copy()
            cand = cand[cand["pmcid"].astype(str).str.strip().astype(bool)].head(int(st.session_state["auto_fulltext_limit"]))
            for _, r in cand.iterrows():
                rid = r["record_id"]
                if (st.session_state["pmc_fulltext"].get(rid) or "").strip():
                    continue
                try:
                    xml = fetch_pmc_fulltext_xml(r.get("pmcid",""))
                    txt = pmc_xml_to_text(xml)
                    if txt.strip():
                        st.session_state["pmc_fulltext"][rid] = txt
                except Exception:
                    continue

    # Step 6: Extraction + ROB2 (needs LLM + FT)
    with st.spinner(t("step6")):
        prot = st.session_state["protocol"]
        merged = st.session_state["records"].merge(st.session_state["screening"], on="record_id", how="left")
        merged = ensure_cols(merged, ["AI_label","AI_reason","AI_confidence"], "")

        # Wide table skeleton
        wide = merged[["record_id","pmid","doi","pmcid","year","first_author","title","url","doi_url","pmc_url","AI_label","AI_reason"]].copy()
        wide["TA_final"] = wide["record_id"].map(lambda rid: st.session_state["ta_final"].get(rid, t("ta_unsure")))
        wide["FT_decision"] = wide["record_id"].map(lambda rid: st.session_state["ft_decision"].get(rid, t("ft_not")))
        wide["FT_reason"] = wide["record_id"].map(lambda rid: st.session_state["ft_reason"].get(rid, ""))

        plan = prot.get("recommended_extraction_schema_plan", {}) or {}
        base_cols = plan.get("base_cols", []) or []
        outcome_items = []
        for g in (plan.get("outcome_groups", []) or []):
            outcome_items.extend(g.get("suggested_items", []) or [])

        for c in base_cols:
            if c not in wide.columns:
                wide[c] = ""
        for o in outcome_items:
            if o and o not in wide.columns:
                wide[o] = ""

        for c in ["Effect_measure","Effect","Lower_CI","Upper_CI","Timepoint","Unit"]:
            if c not in wide.columns:
                wide[c] = ""

        # Only do LLM extraction if enabled + key exists + fulltext source permits
        if llm_available():
            for _, r in wide.iterrows():
                rid = r["record_id"]
                if st.session_state["ta_final"].get(rid) not in [t("ta_include"), t("ta_unsure")]:
                    continue

                # select full text based on user choice
                if st.session_state["fulltext_source"] == "none":
                    ft = ""
                elif st.session_state["fulltext_source"] == "pmc":
                    ft = (st.session_state["pmc_fulltext"].get(rid) or "").strip()
                else:
                    ft = (st.session_state["uploaded_fulltext"] or "").strip()

                if not ft:
                    continue

                ex = extract_llm(ft, prot)
                if ex.get("error"):
                    continue

                dec = str(ex.get("fulltext_decision","")).strip()
                if dec in ["Include for meta-analysis","Exclude after full-text","Not reviewed yet"]:
                    # map to UI lang strings for storage/display
                    if dec.startswith("Include"):
                        dec_ui = t("ft_include")
                    elif dec.startswith("Exclude"):
                        dec_ui = t("ft_exclude")
                    else:
                        dec_ui = t("ft_not")
                    st.session_state["ft_decision"][rid] = dec_ui
                    wide.loc[wide["record_id"] == rid, "FT_decision"] = dec_ui

                reason = str(ex.get("fulltext_reason","") or "").strip()
                if reason:
                    st.session_state["ft_reason"][rid] = reason
                    wide.loc[wide["record_id"] == rid, "FT_reason"] = reason

                schema_obj = ex.get("extraction_schema")
                if schema_obj:
                    if "_extraction_schema_json" not in wide.columns:
                        wide["_extraction_schema_json"] = ""
                    wide.loc[wide["record_id"] == rid, "_extraction_schema_json"] = json.dumps(schema_obj, ensure_ascii=False)

                fields = ex.get("extracted_fields") or {}
                # Put known columns; unknown go into _extra_fields_json
                for k, v in fields.items():
                    if k in wide.columns:
                        wide.loc[wide["record_id"] == rid, k] = "" if v is None else str(v)
                    else:
                        if "_extra_fields_json" not in wide.columns:
                            wide["_extra_fields_json"] = ""
                        cur = wide.loc[wide["record_id"] == rid, "_extra_fields_json"].values[0]
                        bag = json.loads(cur) if cur else {}
                        bag[k] = v
                        wide.loc[wide["record_id"] == rid, "_extra_fields_json"] = json.dumps(bag, ensure_ascii=False)

                meta = ex.get("meta") or {}
                wide.loc[wide["record_id"] == rid, "Effect_measure"] = str(meta.get("effect_measure",""))
                wide.loc[wide["record_id"] == rid, "Effect"] = str(meta.get("effect",""))
                wide.loc[wide["record_id"] == rid, "Lower_CI"] = str(meta.get("lower_CI",""))
                wide.loc[wide["record_id"] == rid, "Upper_CI"] = str(meta.get("upper_CI",""))
                wide.loc[wide["record_id"] == rid, "Timepoint"] = str(meta.get("timepoint",""))
                wide.loc[wide["record_id"] == rid, "Unit"] = str(meta.get("unit",""))

                # ROB2 only for FT included
                if st.session_state["ft_decision"].get(rid) == t("ft_include"):
                    rb = rob2_llm(ft, prot)
                    if not rb.get("error"):
                        st.session_state["rob2_raw"][rid] = rb
                        st.session_state["rob2_table"][rid] = {k: (rb.get(k, {}) or {}).get("judgement","") for k in ROB2_DOMAINS}

        st.session_state["extraction_wide"] = wide

    # Step 7: MA attempt
    with st.spinner(t("step7")):
        wide = st.session_state["extraction_wide"]
        inc = wide[wide["FT_decision"] == t("ft_include")].copy() if isinstance(wide, pd.DataFrame) and not wide.empty else pd.DataFrame()
        ma_df = inc[["record_id","title","Effect_measure","Effect","Lower_CI","Upper_CI","Timepoint","Unit"]].copy() if not inc.empty else pd.DataFrame()
        st.session_state["ma_result"] = fixed_effect_ma(ma_df) if not ma_df.empty else {"error":"no included studies"}

    st.success(t("done"))


# =========================
# Outputs (persistent)
# =========================
st.header(t("outputs"))

prot = st.session_state.get("protocol") or {}
pub_q = st.session_state.get("pubmed_query","")
df = st.session_state.get("records", safe_empty_records_df().iloc[:0].copy())
scr = st.session_state.get("screening", pd.DataFrame())
scan = st.session_state.get("srma_scan", {"summary":"", "hits": pd.DataFrame()})
wide = st.session_state.get("extraction_wide", pd.DataFrame())
ma = st.session_state.get("ma_result")
diag = st.session_state.get("diag", {}) or {}

colA, colB = st.columns([2,1])
with colA:
    st.subheader(t("protocol"))
    st.code(json.dumps(prot, ensure_ascii=False, indent=2), language="json")
with colB:
    st.subheader(t("pubmed_query"))
    st.code(pub_q or "", language="text")

st.subheader(t("diagnostics"))
with st.expander(t("show_diag")):
    st.write({"pubmed_total_count": int(st.session_state.get("pubmed_total") or 0)})
    st.write(diag)

if st.session_state.get("include_srma_scan", True) and scan and scan.get("summary"):
    st.subheader(t("srma_title"))
    st.info(scan["summary"])
    hits = scan.get("hits", pd.DataFrame())
    if isinstance(hits, pd.DataFrame) and not hits.empty:
        st.dataframe(hits, use_container_width=True)
        st.download_button(t("download_csv"), data=to_csv_bytes(hits), file_name="srma_nma_hits.csv")

st.subheader(t("records"))
if isinstance(df, pd.DataFrame) and df.empty:
    st.info(t("no_records"))
else:
    merged = df.merge(scr, on="record_id", how="left")
    merged = ensure_cols(merged, ["AI_label","AI_reason","AI_confidence"], "")
    merged["TA_final"] = merged["record_id"].map(lambda rid: st.session_state["ta_final"].get(rid, t("ta_unsure")))
    merged["FT_decision"] = merged["record_id"].map(lambda rid: st.session_state["ft_decision"].get(rid, t("ft_not")))
    merged["FT_reason"] = merged["record_id"].map(lambda rid: st.session_state["ft_reason"].get(rid, ""))

    st.dataframe(
        merged[["record_id","year","first_author","title","AI_label","AI_reason","TA_final","FT_decision","pmid","doi","pmcid","url","pmc_url"]],
        use_container_width=True
    )
    st.download_button(t("download_csv"), data=to_csv_bytes(merged), file_name="records_with_screening.csv")

st.subheader(t("extraction_wide"))
if isinstance(wide, pd.DataFrame) and not wide.empty:
    edited = st.data_editor(wide, use_container_width=True, hide_index=True, num_rows="dynamic")
    st.session_state["extraction_wide"] = edited
    st.download_button(t("download_csv"), data=to_csv_bytes(edited), file_name="extraction_wide.csv")
else:
    st.caption(t("no_key_degrade"))

st.subheader(t("rob2"))
rob_rows = []
wide2 = st.session_state.get("extraction_wide", pd.DataFrame())
if isinstance(wide2, pd.DataFrame) and not wide2.empty:
    inc2 = wide2[wide2["FT_decision"] == t("ft_include")].copy()
    for _, r in inc2.iterrows():
        rid = r["record_id"]
        rb = st.session_state["rob2_table"].get(rid, {}) or {}
        name = f"{r.get('first_author','')} ({r.get('year','')})".strip() or rid
        rob_rows.append({
            "Study": name,
            "D1": rb.get("D1",""),
            "D2": rb.get("D2",""),
            "D3": rb.get("D3",""),
            "D4": rb.get("D4",""),
            "D5": rb.get("D5",""),
            "Overall": rb.get("Overall",""),
        })
rob_df = pd.DataFrame(rob_rows)
if not rob_df.empty:
    st.dataframe(rob_df, use_container_width=True)
    st.download_button(t("download_csv"), data=to_csv_bytes(rob_df), file_name="rob2.csv")
else:
    st.caption(t("rob2_note"))

st.subheader(t("ma"))
if ma:
    if ma.get("error"):
        st.warning(t("ma_not_ready"))
    else:
        st.success(f"Pooled ({ma.get('measure')}): {ma.get('pooled'):.4g} (95% CI {ma.get('lower'):.4g} to {ma.get('upper'):.4g}), k={ma.get('k')}")
        if HAS_MPL:
            fig = plot_forest(ma)
            if fig is not None:
                st.pyplot(fig, clear_figure=True)

st.subheader(t("prisma"))
df3 = st.session_state.get("records", pd.DataFrame())
ta_vals = [st.session_state["ta_final"].get(rid, t("ta_unsure")) for rid in (df3["record_id"].tolist() if isinstance(df3, pd.DataFrame) and not df3.empty else [])]
prisma = {
    "Records identified (PubMed total count)": int(st.session_state.get("pubmed_total") or 0),
    "Records retrieved (limited by max_records)": int(len(df3)) if isinstance(df3, pd.DataFrame) else 0,
    "TA Include": int(sum(1 for x in ta_vals if x == t("ta_include"))),
    "TA Exclude": int(sum(1 for x in ta_vals if x == t("ta_exclude"))),
    "TA Unsure": int(sum(1 for x in ta_vals if x == t("ta_unsure"))),
    "FT Include for meta-analysis": int(sum(1 for rid in (df3["record_id"].tolist() if isinstance(df3, pd.DataFrame) and not df3.empty else []) if st.session_state["ft_decision"].get(rid) == t("ft_include"))),
}
st.json(prisma)

st.subheader(t("export_word"))
if HAS_DOCX and prot and isinstance(st.session_state.get("extraction_wide", pd.DataFrame()), pd.DataFrame):
    docx_bytes = export_docx(
        question=st.session_state.get("question",""),
        protocol=prot,
        pubmed_query=pub_q,
        prisma=prisma,
        wide=st.session_state.get("extraction_wide", pd.DataFrame()),
        rob_df=rob_df,
        ma=ma if ma and not ma.get("error") else None
    )
    if docx_bytes:
        st.download_button(
            "Download draft_report.docx",
            data=docx_bytes,
            file_name="draft_report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
else:
    st.caption(t("word_missing"))
