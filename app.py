# -*- coding: utf-8 -*-
from __future__ import annotations

import html
import io
import json
import math
import re
import time
from typing import Dict, List, Optional, Tuple

import pandas as pd
import requests
import streamlit as st

# Optional deps
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
# Config
# =========================
NCBI_ESEARCH = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
NCBI_EFETCH  = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"

EXPECTED_RECORD_COLS = [
    "record_id","pmid","pmcid","doi",
    "title","abstract","year","journal","first_author",
    "url","doi_url","pmc_url","source"
]


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


# =========================
# Institution access (no passwords stored)
# =========================
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
            "source": "PubMed",
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

    df = pd.concat(frames, ignore_index=True) if frames else safe_empty_records_df().iloc[:0].copy()
    df = ensure_cols(df, EXPECTED_RECORD_COLS, "")
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
# PMC full text (OA)
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

def pmc_xml_to_text(xml_text: str, max_chars: int = 160000) -> str:
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
            t = norm_text("".join(el.itertext()))
            if t:
                texts.append(t)
            if sum(len(x) for x in texts) > max_chars:
                break
    return ("\n".join(texts))[:max_chars]


# =========================
# LLM (OpenAI-compatible)
# =========================
def llm_available() -> bool:
    ss = st.session_state
    return bool((ss.get("LLM_BASE_URL") or "").strip() and (ss.get("LLM_API_KEY") or "").strip() and (ss.get("LLM_MODEL") or "").strip())

def llm_chat(messages: List[dict], temperature: float = 0.2, timeout: int = 120) -> Optional[str]:
    ss = st.session_state
    base = (ss.get("LLM_BASE_URL") or "").strip().rstrip("/")
    key  = (ss.get("LLM_API_KEY") or "").strip()
    model= (ss.get("LLM_MODEL") or "").strip()
    if not (base and key and model):
        return None
    url = base + "/v1/chat/completions"
    headers = {"Authorization": f"Bearer {key}", "Content-Type": "application/json"}
    payload = {"model": model, "messages": messages, "temperature": float(temperature)}
    try:
        r = requests.post(url, headers=headers, json=payload, timeout=timeout)
        r.raise_for_status()
        js = r.json()
        return js["choices"][0]["message"]["content"]
    except Exception as e:
        st.warning(f"LLM call failed; fallback mode. {e}")
        return None


# =========================
# Question -> Protocol (PICO/criteria/schema/feasibility)
# =========================
def extract_english_tokens(text: str) -> List[str]:
    # pulls useful abbreviations/english device names if user asks in Chinese
    toks = re.findall(r"[A-Za-z][A-Za-z0-9\-\+]{1,}", text or "")
    toks = [t.strip() for t in toks if t.strip()]
    # de-dup preserving order
    seen = set()
    out = []
    for t in toks:
        u = t.upper()
        if u not in seen:
            seen.add(u)
            out.append(t)
    return out

def parse_question_to_pico(question: str) -> dict:
    q = (question or "").strip()
    q_norm = re.sub(r"\s+", " ", q)
    m = re.split(r"\b(vs\.?|versus|compared to|compared with|compare to|comparison)\b", q_norm, flags=re.I)
    if len(m) >= 3:
        left = m[0].strip(" -:;,.")
        right = " ".join(m[2:]).strip(" -:;,.")
        return {"P": "", "I": left, "C": right, "O": "", "NOT": "animal OR mice OR rat OR in vitro OR case report"}
    return {"P": "", "I": q_norm, "C": "", "O": "", "NOT": "animal OR mice OR rat OR in vitro OR case report"}

def protocol_from_question_llm(question: str, goal_mode: str) -> dict:
    sys = "You are a senior SR/MA methodologist. Output ONLY valid JSON."
    user = {
        "task": "From one clinical question, produce protocol + search expansion + feasibility plan + extraction schema plan (not hard-coded).",
        "goal_mode": goal_mode,
        "question": question,
        "output_schema": {
            "pico": {"P":"","I":"","C":"","O":"","NOT":""},
            "inclusion_decision": {
                "scope_tradeoff": "gap_fill_fast OR rigorous_scope",
                "recommended_scope": "How to set PICO boundaries based on existing evidence and feasibility."
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
                    "Check whether prior SR/MA exists and align/extend",
                    "Enumerate ALL RCT primary and secondary outcomes for extraction"
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
                "consider_nma": "yes/no",
                "nma_node_definition": "how to define nodes"
            }
        },
        "constraints": [
            "Be topic-agnostic; do not assume specialty.",
            "Maximize recall: include abbreviations, variant spellings, and common synonyms.",
            "If question is non-English, translate/expand to English search terms."
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

def protocol_fallback(question: str, goal_mode: str) -> dict:
    pico = parse_question_to_pico(question)
    toks = extract_english_tokens(question)
    # If user wrote Chinese sentence, use extracted tokens as I-synonyms to avoid nonsense [tiab]
    I0 = (pico.get("I") or "").strip()
    C0 = (pico.get("C") or "").strip()

    def expand(term: str) -> List[str]:
        t = (term or "").strip()
        if not t:
            return []
        syn = [t, t.replace("-", " "), t.replace(" ", "-")]
        return list(dict.fromkeys([s for s in syn if s]))

    I_syn = expand(I0)
    C_syn = expand(C0)
    # If term is long and contains CJK, prefer tokens
    if re.search(r"[\u4e00-\u9fff]", I0) and toks:
        I_syn = toks
        pico["I"] = toks[0]
    if re.search(r"[\u4e00-\u9fff]", C0) and toks:
        C_syn = toks

    return {
        "pico": pico,
        "inclusion_decision": {"scope_tradeoff":"gap_fill_fast" if goal_mode.startswith("Fast") else "rigorous_scope",
                               "recommended_scope":"Fallback mode. Configure LLM for best PICO boundary + criteria."},
        "inclusion_criteria": ["Human studies relevant to the question."],
        "exclusion_criteria": ["Animal/in vitro", "Case reports/series only (unless specifically needed)."],
        "search_expansion": {"P_synonyms": [], "I_synonyms": I_syn, "C_synonyms": C_syn, "O_synonyms": [],
                             "NOT":["animal","mice","rat","in vitro","case report"]},
        "mesh_candidates": {"P": [], "I": [], "C": [], "O": []},
        "feasibility_plan": {"how_to_check_existing_srma":"Scan SR/MA/NMA in PubMed first.",
                             "what_to_do_if_srma_exists":"Consider narrowing population, newer RCTs, or different comparator."},
        "recommended_extraction_schema_plan": {
            "principles": [
                "Do not hard-code columns; define at PICO level",
                "Check whether prior SR/MA exists and align/extend",
                "Enumerate ALL RCT primary and secondary outcomes for extraction"
            ],
            "base_cols": ["First author","Year","Design","Population details","Intervention","Comparator","Follow-up","Notes (figure/table/section)"],
            "outcome_groups": [{"group_name":"Primary outcomes","suggested_items":["Primary outcome"]},
                               {"group_name":"Secondary outcomes","suggested_items":["Secondary outcome 1","Secondary outcome 2"]}],
            "effect_preference": ["RR","OR","HR","MD","SMD"]
        },
        "analysis_plan": {"effect_measures":["RR","OR","HR","MD","SMD"],"timepoint_preference":["final follow-up"],"consider_nma":"no","nma_node_definition":""},
        "goal_mode": goal_mode
    }


# =========================
# Query builder (recall-friendly; uses I/C even if P empty)
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
        return {"summary": f"Existing SR/MA/NMA scan: PubMed count≈{count} (showing top {min(len(ids), top_n)}).", "hits": hits}
    except Exception as e:
        return {"summary": f"Scan failed: {e}", "hits": pd.DataFrame()}


# =========================
# Title/Abstract screening (LLM optional)
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
        return {"label":"Exclude","confidence":0.8,"reason":"NOT keyword hit"}
    if I and hit(I) and (not C or hit(C)):
        return {"label":"Include","confidence":0.7,"reason":"Key terms present"}
    if any(w in text for w in ["randomized","randomised","trial","controlled"]) and hit(I):
        return {"label":"Include","confidence":0.65,"reason":"Trial-like + match"}
    return {"label":"Unsure","confidence":0.4,"reason":"Insufficient TA signal"}

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
    return {"label": js.get("label","Unsure"), "confidence": js.get("confidence",0.0), "reason": js.get("reason","")}


# =========================
# Full-text extraction + ROB2 (LLM)
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
- Explicitly check/mention whether prior SR/MA exists (if the text references it) and align/extend.
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
# Fixed-effect MA (requires Effect + CI)
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
# Word export
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
        cols = list(wide.columns)[:min(8, len(wide.columns))]
        t = doc.add_table(rows=1, cols=len(cols))
        for i, c in enumerate(cols):
            t.rows[0].cells[i].text = c
        for _, r in wide.head(25).iterrows():
            row = t.add_row().cells
            for i, c in enumerate(cols):
                row[i].text = str(r.get(c, ""))
    else:
        doc.add_paragraph("(none)")

    doc.add_heading("ROB 2.0 (snapshot)", level=2)
    if isinstance(rob_df, pd.DataFrame) and not rob_df.empty:
        t = doc.add_table(rows=1, cols=len(rob_df.columns))
        for i, c in enumerate(rob_df.columns):
            t.rows[0].cells[i].text = c
        for _, r in rob_df.iterrows():
            row = t.add_row().cells
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
# Streamlit UI
# =========================
st.set_page_config(page_title="SR/MA One-question", layout="wide")
st.title("SR/MA (One-question prototype)")
st.caption("一句問題 → protocol/feasibility/query → TA screening → (可選) PMC 全文 → extraction/ROB2 → MA → Word")

def init_state():
    ss = st.session_state
    ss.setdefault("question", "")
    ss.setdefault("goal_mode", "Fast / feasible (gap-fill)")
    ss.setdefault("strict_CO", False)
    ss.setdefault("include_rct_filter", False)
    ss.setdefault("include_srma_scan", True)
    ss.setdefault("max_records", 300)
    ss.setdefault("polite_delay", 0.0)

    ss.setdefault("auto_fetch_pmc", True)
    ss.setdefault("auto_fulltext_limit", 10)

    ss.setdefault("RESOLVER_BASE", "")
    ss.setdefault("EZPROXY", "")

    ss.setdefault("LLM_BASE_URL", "")
    ss.setdefault("LLM_API_KEY", "")
    ss.setdefault("LLM_MODEL", "")

    ss.setdefault("protocol", None)
    ss.setdefault("pubmed_query", "")
    ss.setdefault("srma_scan", {"summary":"", "hits": pd.DataFrame()})
    ss.setdefault("records", safe_empty_records_df().iloc[:0].copy())
    ss.setdefault("pubmed_total", 0)
    ss.setdefault("diag", {})

    ss.setdefault("screening", pd.DataFrame())
    ss.setdefault("ta_final", {})     # record_id -> Include/Exclude/Unsure
    ss.setdefault("ft_decision", {})  # record_id -> Not reviewed / Include for MA / Exclude
    ss.setdefault("ft_reason", {})    # record_id -> reason
    ss.setdefault("fulltext", {})     # record_id -> text

    ss.setdefault("extraction_wide", pd.DataFrame())
    ss.setdefault("rob2_table", {})
    ss.setdefault("rob2_raw", {})
    ss.setdefault("ma_result", None)

init_state()

with st.sidebar:
    st.subheader("Advanced (optional)")
    st.session_state["goal_mode"] = st.selectbox(
        "Scope preference",
        ["Fast / feasible (gap-fill)", "Rigorous / comprehensive"],
        index=0 if st.session_state["goal_mode"].startswith("Fast") else 1
    )
    st.session_state["include_srma_scan"] = st.checkbox("Feasibility: scan existing SR/MA/NMA", value=st.session_state["include_srma_scan"])
    st.session_state["strict_CO"] = st.checkbox("Search requires C/O (lower recall)", value=st.session_state["strict_CO"])
    st.session_state["include_rct_filter"] = st.checkbox("Apply RCT filter (may miss studies)", value=st.session_state["include_rct_filter"])
    st.session_state["max_records"] = st.number_input("Max PubMed records", 50, 5000, int(st.session_state["max_records"]), 50)
    st.session_state["polite_delay"] = st.slider("Polite delay (sec)", 0.0, 1.0, float(st.session_state["polite_delay"]), 0.1)

    st.markdown("---")
    st.session_state["auto_fetch_pmc"] = st.checkbox("Auto-fetch PMC OA full text", value=st.session_state["auto_fetch_pmc"])
    st.session_state["auto_fulltext_limit"] = st.number_input("Auto fulltext fetch limit", 0, 50, int(st.session_state["auto_fulltext_limit"]), 1)

    st.markdown("---")
    st.markdown("Institution access (optional; no passwords stored)")
    st.session_state["RESOLVER_BASE"] = st.text_input("OpenURL resolver base", value=st.session_state["RESOLVER_BASE"])
    st.session_state["EZPROXY"] = st.text_input("EZproxy prefix", value=st.session_state["EZPROXY"])

    st.markdown("---")
    st.markdown("LLM (OpenAI-compatible; required for extraction/ROB2 auto)")
    st.session_state["LLM_BASE_URL"] = st.text_input("Base URL", value=st.session_state["LLM_BASE_URL"])
    st.session_state["LLM_API_KEY"] = st.text_input("API Key", value=st.session_state["LLM_API_KEY"], type="password")
    st.session_state["LLM_MODEL"] = st.text_input("Model", value=st.session_state["LLM_MODEL"])
    st.caption("沒填：仍會抓文獻 + 規則初篩；但不會自動 full-text extraction/ROB2。")

st.session_state["question"] = st.text_area(
    "Research question (one sentence)",
    value=st.session_state["question"],
    height=80,
    placeholder="e.g., Does intervention A improve outcome B vs comparator C in population P?"
)

run = st.button("Run (one-click pipeline)", type="primary")

if run:
    q = (st.session_state["question"] or "").strip()
    if not q:
        st.error("Please enter a research question.")
        st.stop()

    # 0) Protocol
    with st.spinner("0/7 Building protocol (PICO/criteria/feasibility/schema)…"):
        if llm_available():
            prot = protocol_from_question_llm(q, st.session_state["goal_mode"])
            if prot.get("error"):
                st.warning("Protocol JSON parse failed; fallback used.")
                prot = protocol_fallback(q, st.session_state["goal_mode"])
        else:
            prot = protocol_fallback(q, st.session_state["goal_mode"])
        st.session_state["protocol"] = prot

    # 1) Query
    with st.spinner("1/7 Building PubMed query…"):
        st.session_state["pubmed_query"] = build_pubmed_query(
            st.session_state["protocol"],
            st.session_state["strict_CO"],
            st.session_state["include_rct_filter"]
        )

    # 2) Feasibility scan
    if st.session_state["include_srma_scan"]:
        with st.spinner("2/7 Feasibility scan (SR/MA/NMA)…"):
            st.session_state["srma_scan"] = scan_sr_ma_nma(st.session_state["pubmed_query"], top_n=25)

    # 3) Fetch PubMed
    with st.spinner("3/7 Fetching PubMed records…"):
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
            st.session_state["ta_final"].setdefault(rid, "Unsure")
            st.session_state["ft_decision"].setdefault(rid, "Not reviewed yet")
            st.session_state["ft_reason"].setdefault(rid, "")
            st.session_state["fulltext"].setdefault(rid, "")

    # 4) TA screening
    with st.spinner("4/7 Title/Abstract screening…"):
        out = []
        for _, r in st.session_state["records"].iterrows():
            res = screen_llm(r, st.session_state["protocol"]) if llm_available() else screen_rule_based(r, st.session_state["protocol"])
            rid = r["record_id"]
            st.session_state["ta_final"][rid] = res["label"]
            out.append({"record_id": rid, "AI_label": res["label"], "AI_confidence": res["confidence"], "AI_reason": res["reason"]})
        st.session_state["screening"] = pd.DataFrame(out)

    # 5) Auto-fetch PMC OA full text (candidates)
    if st.session_state["auto_fetch_pmc"]:
        with st.spinner("5/7 Auto-fetching PMC OA full text…"):
            merged = st.session_state["records"].merge(st.session_state["screening"], on="record_id", how="left")
            merged = ensure_cols(merged, ["AI_label"], "")
            cand = merged[merged["AI_label"].isin(["Include","Unsure"])].copy()
            cand = cand[cand["pmcid"].astype(str).str.strip().astype(bool)].head(int(st.session_state["auto_fulltext_limit"]))
            for _, r in cand.iterrows():
                rid = r["record_id"]
                if (st.session_state["fulltext"].get(rid) or "").strip():
                    continue
                try:
                    xml = fetch_pmc_fulltext_xml(r.get("pmcid",""))
                    txt = pmc_xml_to_text(xml)
                    if txt.strip():
                        st.session_state["fulltext"][rid] = txt
                except Exception:
                    pass

    # 6) Extraction + ROB2 (LLM only) + wide table
    with st.spinner("6/7 Extraction + ROB2 (when full text available)…"):
        prot = st.session_state["protocol"]
        merged = st.session_state["records"].merge(st.session_state["screening"], on="record_id", how="left")
        merged = ensure_cols(merged, ["AI_label","AI_reason","AI_confidence"], "")

        # add institution links
        resolver = st.session_state["RESOLVER_BASE"]
        ezp = st.session_state["EZPROXY"]
        if resolver:
            merged["openurl"] = merged.apply(
                lambda x: apply_ezproxy(ezp, build_openurl(resolver, doi=x.get("doi",""), pmid=x.get("pmid",""), title=x.get("title",""))),
                axis=1
            )
        else:
            merged["openurl"] = ""

        plan = prot.get("recommended_extraction_schema_plan", {}) or {}
        base_cols = plan.get("base_cols", []) or []
        outcome_items = []
        for g in (plan.get("outcome_groups", []) or []):
            outcome_items.extend(g.get("suggested_items", []) or [])

        wide = merged[["record_id","pmid","doi","year","first_author","title","url","doi_url","pmc_url","openurl","AI_label","AI_reason"]].copy()
        wide["TA_final"] = wide["record_id"].map(lambda rid: st.session_state["ta_final"].get(rid, "Unsure"))
        wide["FT_decision"] = wide["record_id"].map(lambda rid: st.session_state["ft_decision"].get(rid, "Not reviewed yet"))
        wide["FT_reason"] = wide["record_id"].map(lambda rid: st.session_state["ft_reason"].get(rid, ""))

        for c in base_cols:
            if c not in wide.columns:
                wide[c] = ""
        for o in outcome_items:
            if o and o not in wide.columns:
                wide[o] = ""

        for c in ["Effect_measure","Effect","Lower_CI","Upper_CI","Timepoint","Unit"]:
            if c not in wide.columns:
                wide[c] = ""

        if llm_available():
            for _, r in wide.iterrows():
                rid = r["record_id"]
                if st.session_state["ta_final"].get(rid) not in ["Include","Unsure"]:
                    continue
                ft = (st.session_state["fulltext"].get(rid) or "").strip()
                if not ft:
                    continue

                ex = extract_llm(ft, prot)
                if ex.get("error"):
                    continue

                dec = ex.get("fulltext_decision")
                if dec in ["Include for meta-analysis","Exclude after full-text","Not reviewed yet"]:
                    st.session_state["ft_decision"][rid] = dec
                    wide.loc[wide["record_id"] == rid, "FT_decision"] = dec

                if ex.get("fulltext_reason"):
                    st.session_state["ft_reason"][rid] = str(ex.get("fulltext_reason"))
                    wide.loc[wide["record_id"] == rid, "FT_reason"] = str(ex.get("fulltext_reason"))

                schema_obj = ex.get("extraction_schema")
                if schema_obj:
                    if "_extraction_schema_json" not in wide.columns:
                        wide["_extraction_schema_json"] = ""
                    wide.loc[wide["record_id"] == rid, "_extraction_schema_json"] = json.dumps(schema_obj, ensure_ascii=False)

                fields = ex.get("extracted_fields") or {}
                for k, v in fields.items():
                    if k not in wide.columns:
                        if "_extra_fields_json" not in wide.columns:
                            wide["_extra_fields_json"] = ""
                        cur = wide.loc[wide["record_id"] == rid, "_extra_fields_json"].values[0]
                        bag = json.loads(cur) if cur else {}
                        bag[k] = v
                        wide.loc[wide["record_id"] == rid, "_extra_fields_json"] = json.dumps(bag, ensure_ascii=False)
                    else:
                        wide.loc[wide["record_id"] == rid, k] = "" if v is None else str(v)

                meta = ex.get("meta") or {}
                wide.loc[wide["record_id"] == rid, "Effect_measure"] = str(meta.get("effect_measure",""))
                wide.loc[wide["record_id"] == rid, "Effect"] = str(meta.get("effect",""))
                wide.loc[wide["record_id"] == rid, "Lower_CI"] = str(meta.get("lower_CI",""))
                wide.loc[wide["record_id"] == rid, "Upper_CI"] = str(meta.get("upper_CI",""))
                wide.loc[wide["record_id"] == rid, "Timepoint"] = str(meta.get("timepoint",""))
                wide.loc[wide["record_id"] == rid, "Unit"] = str(meta.get("unit",""))

                if st.session_state["ft_decision"].get(rid) == "Include for meta-analysis":
                    rb = rob2_llm(ft, prot)
                    if not rb.get("error"):
                        st.session_state["rob2_raw"][rid] = rb
                        st.session_state["rob2_table"][rid] = {k: (rb.get(k, {}) or {}).get("judgement","") for k in ["D1","D2","D3","D4","D5","Overall"]}

        st.session_state["extraction_wide"] = wide

    # 7) MA attempt
    with st.spinner("7/7 Meta-analysis (fixed effect) attempt…"):
        wide = st.session_state["extraction_wide"]
        inc = wide[wide["FT_decision"] == "Include for meta-analysis"].copy() if isinstance(wide, pd.DataFrame) else pd.DataFrame()
        ma_df = inc[["record_id","title","Effect_measure","Effect","Lower_CI","Upper_CI","Timepoint","Unit"]].copy() if not inc.empty else pd.DataFrame()
        st.session_state["ma_result"] = fixed_effect_ma(ma_df) if not ma_df.empty else {"error":"no included studies"}

    st.success("Done. Scroll down for outputs.")


# =========================
# Outputs (persistent)
# =========================
st.header("Outputs")

prot = st.session_state.get("protocol")
pub_q = st.session_state.get("pubmed_query","")
df = st.session_state.get("records", safe_empty_records_df().iloc[:0].copy())
scr = st.session_state.get("screening", pd.DataFrame())
scan = st.session_state.get("srma_scan", {"summary":"", "hits": pd.DataFrame()})
wide = st.session_state.get("extraction_wide", pd.DataFrame())
ma = st.session_state.get("ma_result")
diag = st.session_state.get("diag", {}) or {}

colA, colB = st.columns([2,1])
with colA:
    st.subheader("Protocol (auto)")
    st.code(json.dumps(prot or {}, ensure_ascii=False, indent=2), language="json")
with colB:
    st.subheader("PubMed query (auto)")
    st.code(pub_q or "", language="text")

st.subheader("Diagnostics")
with st.expander("Show diagnostics"):
    st.write({"pubmed_total_count": int(st.session_state.get("pubmed_total") or 0)})
    st.write(diag)

if scan and scan.get("summary"):
    st.subheader("Feasibility: Existing SR/MA/NMA scan")
    st.info(scan["summary"])
    hits = scan.get("hits", pd.DataFrame())
    if isinstance(hits, pd.DataFrame) and not hits.empty:
        st.dataframe(hits, use_container_width=True)
        st.download_button("Download SR/MA/NMA hits (CSV)", data=to_csv_bytes(hits), file_name="srma_nma_hits.csv")

if isinstance(df, pd.DataFrame) and not df.empty:
    st.subheader("Search summary")
    ta_vals = [st.session_state["ta_final"].get(rid, "Unsure") for rid in df["record_id"].tolist()]
    st.write({
        "PubMed total count (query)": int(st.session_state.get("pubmed_total") or 0),
        "Retrieved (dedup, limited by max_records)": int(len(df)),
        "TA Include": int(sum(1 for x in ta_vals if x == "Include")),
        "TA Exclude": int(sum(1 for x in ta_vals if x == "Exclude")),
        "TA Unsure": int(sum(1 for x in ta_vals if x == "Unsure")),
    })

    st.subheader("Records (TA screening results)")
    merged = df.merge(scr, on="record_id", how="left")
    merged = ensure_cols(merged, ["AI_label","AI_reason","AI_confidence"], "")
    merged["TA_final"] = merged["record_id"].map(lambda rid: st.session_state["ta_final"].get(rid, "Unsure"))
    st.dataframe(merged[["record_id","year","first_author","title","AI_label","AI_reason","pmid","doi","pmcid","url","pmc_url"]], use_container_width=True)
    st.download_button("Download records_with_screening.csv", data=to_csv_bytes(merged), file_name="records_with_screening.csv")

if isinstance(wide, pd.DataFrame) and not wide.empty:
    st.subheader("Extraction wide table (editable)")
    edited = st.data_editor(wide, use_container_width=True, hide_index=True, num_rows="dynamic")
    st.session_state["extraction_wide"] = edited
    st.download_button("Download extraction_wide.csv", data=to_csv_bytes(edited), file_name="extraction_wide.csv")

st.subheader("ROB 2.0 (table)")
rob_rows = []
if isinstance(wide, pd.DataFrame) and not wide.empty:
    inc = wide[wide["FT_decision"] == "Include for meta-analysis"].copy()
    for _, r in inc.iterrows():
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
    st.download_button("Download rob2.csv", data=to_csv_bytes(rob_df), file_name="rob2.csv")
else:
    st.caption("No FT included studies yet, or LLM not configured. ROB2 is drafted after FT inclusion.")

st.subheader("Meta-analysis (fixed effect)")
if ma:
    if ma.get("error"):
        st.warning(f"MA not available: {ma.get('error')}")
        st.caption("Needs: FT_decision=Include for meta-analysis and Effect + CI in the wide table (usually from LLM extraction or manual entry).")
    else:
        st.success(f"Pooled ({ma.get('measure')}): {ma.get('pooled'):.4g} (95% CI {ma.get('lower'):.4g} to {ma.get('upper'):.4g}), k={ma.get('k')}")
        if HAS_MPL:
            fig = plot_forest(ma)
            if fig is not None:
                st.pyplot(fig, clear_figure=True)

st.subheader("PRISMA (draft counts)")
prisma = {
    "Records identified (PubMed total count)": int(st.session_state.get("pubmed_total") or 0),
    "Records retrieved (limited by max_records)": int(len(df)) if isinstance(df, pd.DataFrame) else 0,
    "TA Include": int(sum(1 for rid in (df["record_id"].tolist() if isinstance(df, pd.DataFrame) and not df.empty else []) if st.session_state["ta_final"].get(rid) == "Include")),
    "FT Include for MA": int(sum(1 for rid in (df["record_id"].tolist() if isinstance(df, pd.DataFrame) and not df.empty else []) if st.session_state["ft_decision"].get(rid) == "Include for meta-analysis")),
}
st.json(prisma)

st.subheader("Export Word (draft)")
if HAS_DOCX and prot and isinstance(wide, pd.DataFrame):
    docx_bytes = export_docx(
        question=st.session_state.get("question",""),
        protocol=prot,
        pubmed_query=pub_q,
        prisma=prisma,
        wide=wide,
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
    st.caption("Word export requires python-docx.")
