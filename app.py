# app.py
# =========================================================
# 一句話帶你完成 Meta-analysis（BYOK / Traditional Chinese）
# Author: Ya Hsin Yao
#
# 免責聲明：本工具僅供學術研究/教學用途，不構成醫療建議或法律意見；
# 使用者須自行驗證所有結果、數值、引用與全文內容；請勿上傳可識別之病人資訊。
#
# 校內資源/授權提醒（重要）：
# - 若文章來自校內訂閱（付費期刊/EZproxy/館藏系統），請勿將受版權保護之全文上傳至任何第三方服務或公開部署之網站。
# - 請遵守圖書館授權條款：避免大量下載/自動化批次擷取、避免共享全文給未授權者。
# - 若不確定是否可上傳：建議改用「本機版」或僅上傳你有權分享的開放取用全文（OA/PMC）。
#
# Privacy notice (BYOK):
# - Key only used for this session; do not use on untrusted deployments; do not upload identifiable patient info.
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

# Optional: Plotly (for forest plot)
try:
    import plotly.graph_objects as go  # type: ignore
    HAS_PLOTLY = True
except Exception:
    HAS_PLOTLY = False

# Set Streamlit page config
st.set_page_config(layout="wide")

# =========================================================
# Constants and Globals
# =========================================================
MODEL_NAME = "gemini-2.5-flash"
MODEL_NAME_ROB = "gemini-2.5-pro" # RoB 2.0 needs stronger reasoning
MAX_CHAR_COUNT = 16000 # Max characters for LLM processing
MAX_ABSTRACTS = 500 # Max number of abstracts to fetch
# Used for Step 4/5/6: list of exclusion reasons for the dropdown
exclusion_reasons_options = [
    'None',
    'Wrong P', 'Wrong I', 'Wrong C', 'Wrong O', 'Wrong S (Study Design)',
    'Duplicate', 'Not full text article (e.g., Letter, Protocol, Review)',
    'Other (Specify in notes)'
]

# =========================================================
# Session State Initialization
# =========================================================
if 'pico' not in st.session_state:
    st.session_state.pico = {
        'P': 'Patients/Population (e.g., Type II Diabetes, Adults > 65)',
        'I': 'Intervention (e.g., Drug X 10mg daily, Cognitive Behavioral Therapy)',
        'C': 'Comparison (e.g., Placebo, Standard care, Drug Y)',
        'O': 'Outcome (e.g., HbA1c, Mortality, Quality of Life)',
        'S': 'Study Design (e.g., RCTs, Cohort studies)',
        'L': 'LLM System Prompt Language (e.g., Traditional Chinese)'
    }
if 'settings' not in st.session_state:
    st.session_state.settings = {
        'api_key': '',
        'search_terms': '',
        'max_results': 100,
        'email': 'my.email@example.com'
    }
if 'studies' not in st.session_state:
    st.session_state.studies = {}
if 'diagnostics' not in st.session_state:
    st.session_state.diagnostics = {}
if 'locale' not in st.session_state:
    st.session_state.locale = {}
if 'manuscript' not in st.session_state:
    st.session_state.manuscript = {
        'Title': 'Effect of X vs Y on Z: A Systematic Review and Meta-analysis',
        'Introduction': '',
        'Methods': '',
        'Results': '',
        'Discussion': '',
    }
if 'meta_analysis_results' not in st.session_state:
    st.session_state.meta_analysis_results = {}
if 'meta_analysis_data' not in st.session_state:
    st.session_state.meta_analysis_data = {}


# =========================================================
# Helper Functions
# =========================================================

def format_abstract(abstract: str) -> str:
    """
    格式化摘要，將常見的學術標題（如 Purpose: Methods:）轉換為粗體並斷行。
    **修改目標 1: 增加 RESULT 和 DISCUSSION 斷句**
    """
    # 斷句：把 PURPOSE: METHODS: ... 變成新的一行
    abstract = re.sub(
        r"(?<!\n)\b(PURPOSE|METHODS|RESULTS|RESULT|DISCUSSION|CONCLUSIONS|CONCLUSION|BACKGROUND|DESIGN|SETTING|PATIENTS|INTERVENTION|MAIN OUTCOME MEASURES|IMPORTANCE|OBJECTIVE|DATA SOURCES|STUDY SELECTION|DATA EXTRACTION|LIMITATIONS)\s*:\s*",
        r"\n\n**\1:** ",
        abstract,
        flags=re.IGNORECASE
    )
    # 移除摘要開頭可能的空格
    abstract = abstract.lstrip()
    return abstract

def parse_pubmed_result(xml_string: str) -> List[Dict[str, str]]:
    """
    解析 PubMed EFetch 服務返回的 XML 字串，提取關鍵資訊。
    """
    results: List[Dict[str, str]] = []
    try:
        root = ET.fromstring(xml_string)
    except ET.ParseError as e:
        st.error(f"PubMed XML 解析錯誤: {e}")
        st.code(xml_string)
        return results

    # 命名空間（防止命名空間影響尋找）
    for article in root.findall('.//PubmedArticle'):
        try:
            # PMCID/PMID
            pmid = article.findtext('.//PMID')
            doi_element = article.find(".//ELocationID[@EIdType='doi']")
            doi = doi_element.text if doi_element is not None else ''

            if not pmid:
                continue

            # Title
            title = article.findtext('.//ArticleTitle') or 'No Title Available'

            # Year
            year = article.findtext('.//PubDate/Year') or 'N/A'
            if year == 'N/A':
                medline_date = article.findtext('.//PubDate/MedlineDate')
                if medline_date:
                    match = re.search(r'\d{4}', medline_date)
                    year = match.group(0) if match else 'N/A'
            
            # Authors (up to 3, et al.)
            author_list = article.findall('.//AuthorList/Author')
            authors: List[str] = []
            for i, author in enumerate(author_list):
                if i >= 3:
                    authors.append("et al.")
                    break
                lastname = author.findtext('./LastName')
                initials = author.findtext('./Initials')
                if lastname and initials:
                    authors.append(f"{lastname} {initials}")
                elif lastname:
                    authors.append(lastname)
                elif author.findtext('./CollectiveName'):
                    authors.append(author.findtext('./CollectiveName'))
            
            # Journal
            journal = article.findtext('.//Title') or article.findtext('.//MedlineTA') or 'N/A'
            
            # Abstract
            abstract_texts: List[str] = []
            for abstract_text in article.findall('.//AbstractText'):
                label = abstract_text.get('Label')
                text = abstract_text.text or ''
                if label:
                    abstract_texts.append(f"{label}: {text}")
                else:
                    abstract_texts.append(text)
            
            abstract = ' '.join(abstract_texts).strip()

            if abstract == '':
                abstract = 'No Abstract Available'
                
            results.append({
                'pmid': pmid,
                'doi': doi,
                'title': html.unescape(title),
                'year': year,
                'journal': html.unescape(journal),
                'authors': ', '.join(authors),
                'abstract': html.unescape(abstract),
                'lbl': 'Unsure', # Initial label for Step 2
                'final_lbl': 'Unsure', # Initial label for Step 3+4
                'exclusion_reasons': ['None'],
                'data_extracted': False,
                'full_text_clean': '',
                'pico_json': {},
                'rob_json': {},
            })

        except Exception as e:
            st.warning(f"Error parsing article with PMID/DOI: {pmid or 'N/A'} / {doi or 'N/A'}. Error: {e}")
            continue

    return results

# =========================================================
# LLM Integration (Stubs for full app)
# =========================================================

# (Stubs for get_model and call_llm_with_prompt are omitted for brevity,
# as they are not directly related to the NameError, but the user should
# keep them in their full script. The locale and t() function are essential.)

@st.cache_resource
def get_model(api_key: str, model_name: str):
    """Placeholder for get_model"""
    return None

def call_llm_with_prompt(client, model_name: str, system_prompt: str, user_prompt: str, max_output_tokens: int = 4096) -> Optional[str]:
    """Placeholder for call_llm_with_prompt"""
    return None

# =========================================================
# Localize / Traditional Chinese
# =========================================================
def get_locale():
    # zh-TW
    return {
        "app_title": "一句話帶你完成 Meta-analysis",
        "menu_about": "關於 / 設定",
        "tab_settings": "設定 (API & PICO)",
        "tab_guide": "使用指南",
        
        "tabs_search": "Step 1 Search",
        "tabs_abstract": "Step 2 Abstract Screening",
        "tabs_step34": "Step 3+4 Records + 粗篩", # <--- MODIFIED LINE (目標 2a)
        "tabs_rob_a": "Step 5 Qualitative Synthesis",
        "tabs_rob_b": "Step 6a Data Extraction",
        "tabs_rob_c": "Step 6b RoB 2.0 (RCT)",
        "tabs_manuscript": "Step 7 Manuscript Draft",
        "tabs_diag": "Diagnostics",

        "step1_instructions": "輸入 PubMed 檢索詞。僅支援 PubMed ESearch/EFetch。單次限制 100 篇。",
        "step2_instructions": "請 AI 根據 PICO 快速篩選摘要，您可覆寫 AI 建議。",
        "step34_instructions": "依據納排準則（PICO-S）進行粗篩。請將所有研究結果納入考慮。",
        "step5_instructions": "彙整納入研究的關鍵資訊（如：特徵、設計）。",
        "step6a_instructions": "針對納入文獻進行資料擷取。AI 將嘗試提取 PICO-S 資訊及結果。",
        "step6b_instructions": "針對隨機對照試驗 (RCT) 進行 RoB 2.0 偏倚風險評估。",
        
        "lbl_msg_include": "納入：文獻符合 PICO 準則。",
        "lbl_msg_exclude": "排除：文獻不符合 PICO 準則。",
        "lbl_msg_unsure": "不確定：無法僅憑摘要判斷，需看全文。",
        "lbl_msg_not_reviewed": "尚待審查。",

        "lbl_msg_include_for_meta_analysis": "納入 Meta 分析：文獻設計和結果適合量化分析。",
        
        "err_no_api": "請在「設定」頁籤中輸入 Google AI API Key。",
        "err_search_fail": "檢索失敗：請檢查 API Key、搜尋詞和網路連線。",
        "err_no_abstracts": "找不到符合條件的摘要。",
        "err_full_text_too_long": "全文太長（>16,000 字元），AI 無法處理。",
        "err_pdf_fail": "PDF 解析失敗。請確認檔案是有效的文字 PDF。",
        
        "btn_search": "1. 執行 PubMed 檢索 (ESearch/EFetch)",
        "btn_screen_abstracts": "2. 執行摘要篩選 (LLM)",
        "btn_clean_full_text": "3. 僅清理與結構化全文 (LLM)",
        "btn_extract_data": "4. 提取 PICO-S & 結果資料 (LLM)",
        "btn_generate_rob2": "5. 產生 RoB 2.0 建議 (LLM)",
        "btn_generate_synthesis": "6. 產生定性綜合摘要 (LLM)",
        "btn_generate_meta_analysis": "7. 執行 Meta 分析（固定效應）並產生報告草稿",
        "btn_save_ms": "儲存手稿草稿",
        "btn_draft_section": "產生章節草稿",
        "btn_upload_pdf": "上傳 PDF/TXT/DOCX 全文檔案",
        "btn_reset_study": "重置研究資料 (清空所有資料擷取結果)",
        
        "roba_system_prompt": "你是一位資深系統性回顧專家，請根據提供的納入文獻 PICO 資訊和定性數據，撰寫一份結構化、客觀的定性綜合報告。請確保報告內容基於數據而非主觀推測。",
        "pico_system_prompt": "你是一位資深系統性回顧專家。請根據 PICO 準則，判斷這篇摘要是否應納入。然後，根據全文（若有）提取結構化數據。務必以 JSON 格式回傳，且 JSON 內容不含任何前言、解釋或Markdown格式。",
        "rob2_system_prompt": "你是一位資深系統性回顧專家。你的任務是評估一篇隨機對照試驗 (RCT) 的偏倚風險 (RoB 2.0)。請仔細閱讀提供的全文，並對每個領域 (Domain) 進行風險判斷 (Low/Some concerns/High/Unclear)，並提供詳細的理由 (Rationale)。請務必以 JSON 格式回傳，且 JSON 內容不含任何前言、解釋或Markdown格式。",
        "ms_system_prompt": "你是一位資深系統性回顧專家。請根據提供的所有資料，撰寫手稿草稿的 {section} 章節。請確保內容準確、結構化、專業且客觀。請使用繁體中文。",
    }

def t(key: str) -> str:
    """
    本地化查找函數
    """
    return st.session_state.get("locale", get_locale()).get(key, key)

# =========================================================
# Custom Sidebar (Stub)
# =========================================================
with st.sidebar:
    st.title(t("app_title"))
    st.subheader("PICO Statement")
    # ... (Sidebar content)
    st.info("Sidebar content placeholder.")


# =========================================================
# Main Tabs Layout (Crucial missing part)
# =========================================================
tab_names = [
    t("tab_settings"),
    t("tab_guide"),
    t("tabs_search"),
    t("tabs_abstract"),
    t("tabs_step34"),
    t("tabs_rob_a"),
    t("tabs_rob_b"),
    t("tabs_rob_c"),
    t("tabs_manuscript"),
    t("tabs_diag"),
]

# Define the 'tabs' list here
tabs = st.tabs(tab_names)


# =========================================================
# Tab 0: Settings & PICO (Stub)
# =========================================================
with tabs[0]:
    st.subheader(t("tab_settings"))
    # ... (Tab content)
    st.info("Tab 0 (Settings) content placeholder.")

# =========================================================
# Tab 1: Guide (Stub)
# =========================================================
with tabs[1]:
    st.subheader(t("tab_guide"))
    # ... (Tab content)
    st.info("Tab 1 (Guide) content placeholder.")

# =========================================================
# Tab 2: Step 1 Search (Stub)
# =========================================================
with tabs[2]:
    st.subheader(t("tabs_search"))
    st.caption(t("step1_instructions"))
    # ... (Tab content)
    st.info("Tab 2 (Search) content placeholder.")

# =========================================================
# Tab 3: Step 2 Abstract Screening (Stub)
# =========================================================
with tabs[3]:
    st.subheader(t("tabs_abstract"))
    st.caption(t("step2_instructions"))
    # ... (Tab content)
    st.info("Tab 3 (Abstract Screening) content placeholder.")


# =========================================================
# Tab 4: Step 3+4 Records + 粗篩 (MODIFIED LOGIC)
# =========================================================
with tabs[4]:
    st.subheader(t("tabs_step34"))
    st.caption(t("step34_instructions"))
    
    # 處理數據
    studies_df = pd.DataFrame(st.session_state.studies).T.fillna({
        'final_lbl': 'Unsure', 
        'exclusion_reasons': ['None'],
        'data_extracted': False,
        'full_text_clean': '',
        'pico_json': {},
        'rob_json': {},
    })

    if not studies_df.empty:
        st.write(f"共 {len(studies_df)} 篇文獻。")

        # START OF MODIFIED LOGIC: Group studies by final_lbl (目標 2b)
        grouped_studies = {
            "INCLUDE": [],
            "EXCLUDE": [],
            "UNSURE": []
        }

        # Populate groups
        for pmcid, study in studies_df.iterrows():
            final_lbl = study.get('final_lbl') or study.get('lbl') or 'Unsure'
            if final_lbl in ("Include", "Include for Meta-analysis"):
                grouped_studies["INCLUDE"].append((pmcid, study))
            elif final_lbl == "Exclude":
                grouped_studies["EXCLUDE"].append((pmcid, study))
            else: # "Unsure", "Not reviewed"
                grouped_studies["UNSURE"].append((pmcid, study))

        # Define display order and labels
        order = ["INCLUDE", "EXCLUDE", "UNSURE"]
        labels = {
            "INCLUDE": f"✅ INCLUDE / 納入分析 ({len(grouped_studies['INCLUDE'])})",
            "EXCLUDE": f"❌ EXCLUDE / 排除 ({len(grouped_studies['EXCLUDE'])})",
            "UNSURE": f"❓ UNSURE / 尚待決定 ({len(grouped_studies['UNSURE'])})"
        }
        
        # Display each group in order
        for group_key in order:
            studies_in_group = grouped_studies[group_key]
            if studies_in_group:
                # Display the group header (subheader)
                st.subheader(labels[group_key])
                
                # Iterate and display each study in the group
                for pmcid, study in studies_in_group:
                    
                    final_lbl = study.get('final_lbl') or study.get('lbl') or 'Unsure'
                    title_abstract = f"({study['year']}) {study['title']}"
                    
                    # Determine Expander color/icon based on final_lbl (used for visual only)
                    expander_icon = " "
                    if final_lbl == 'Exclude':
                        expander_icon = "❌"
                    elif final_lbl in ('Include', 'Include for Meta-analysis'):
                        expander_icon = "✅"
                    elif final_lbl == 'Unsure':
                        expander_icon = "❓"
                    
                    # Use st.expander for each study
                    with st.expander(f"{expander_icon} {pmcid} - {title_abstract}", expanded=False):
                        
                        st.markdown(f"**PMCID/DOI:** `{pmcid}` ([View in PubMed](https://pubmed.ncbi.nlm.nih.gov/{pmcid}/))")
                        st.markdown(f"**Authors:** {study['authors']}")
                        st.markdown(f"**Journal:** {study['journal']} ({study['year']})")

                        # AI Screening Result (Step 2)
                        st.markdown("---")
                        st.markdown("**Step 2 (Abstract) AI Recommendation:**")
                        lbl = study.get('lbl') or 'Not reviewed'
                        lbl_msg = t(f"lbl_msg_{lbl.lower().replace(' ', '_')}")
                        st.markdown(f"**Status: `{lbl}`**")
                        st.caption(f"{lbl_msg}")
                        
                        # Full Text/PICO info
                        full_text = study.get('full_text_clean')
                        if full_text:
                            # AI Full Text Suggestion (Step 4)
                            if 'lbl_ft' in study:
                                st.markdown("---")
                                st.markdown("**Step 4 (Full-Text) AI Recommendation:**")
                                lbl_ft = study.get('lbl_ft')
                                lbl_ft_msg = t(f"lbl_msg_{lbl_ft.lower().replace(' ', '_')}")
                                st.markdown(f"**Status: `{lbl_ft}`**")
                                st.caption(f"{lbl_ft_msg}")
                            
                            # PICO Extraction
                            if 'pico_json' in study and study['pico_json']:
                                pico_json = study['pico_json']
                                st.markdown("---")
                                st.markdown("**Step 4 (PICO) Extraction:**")
                                st.json(pico_json)
                        else:
                            st.caption("No full text available to run Step 4 (Full-Text Screening/PICO).")
                        
                        # Step 3+4 Manual Decision
                        st.markdown("---")
                        st.markdown("**Step 3+4 Manual Decision (粗篩):**")
                        
                        # Dropdown for manual label
                        new_lbl = st.selectbox(
                            "Your Decision:",
                            options=['Unsure', 'Include', 'Include for Meta-analysis', 'Exclude'],
                            index=['Unsure', 'Include', 'Include for Meta-analysis', 'Exclude'].index(final_lbl),
                            key=f"step34_lbl_{pmcid}"
                        )
                        
                        # Dropdown for exclusion reason
                        exclusion_reasons = study.get('exclusion_reasons') or ['None']
                        exclusion_index = 0
                        
                        if final_lbl == 'Exclude' and exclusion_reasons[0] != 'None':
                            try:
                                exclusion_index = exclusion_reasons_options.index(exclusion_reasons[0])
                            except ValueError:
                                pass # Default to None if not found

                        new_reason = st.selectbox(
                            "Exclusion Reason (if Exclude):",
                            options=exclusion_reasons_options,
                            index=exclusion_index,
                            key=f"step34_reason_{pmcid}"
                        )
                        
                        # Logic to save the decision
                        if st.button("Save Decision", key=f"step34_save_{pmcid}"):
                            # Update session state with manual decision
                            st.session_state.studies[pmcid]['final_lbl'] = new_lbl
                            if new_lbl == 'Exclude':
                                st.session_state.studies[pmcid]['exclusion_reasons'] = [new_reason]
                            else:
                                st.session_state.studies[pmcid]['exclusion_reasons'] = ['None']
                            st.success(f"Decision saved for {pmcid} as `{new_lbl}`.")
                            st.experimental_rerun()
                        
                        # Display Abstract
                        st.markdown("---")
                        st.markdown("**Abstract:**")
                        abstract = study.get('abstract') or "N/A"
                        st.markdown(format_abstract(abstract))

                        # Optional Full Text Display
                        if st.checkbox("Show Full Text", key=f"step34_show_ft_{pmcid}"):
                            st.markdown("---")
                            st.markdown("**Full Text (AI Parsed):**")
                            if full_text:
                                st.markdown(full_text)
                            else:
                                st.warning("Full text not available in the database.")
        # END OF MODIFIED LOGIC

    else:
        st.info("請先在 Step 1 執行 PubMed 檢索。")

# =========================================================
# Tab 5: Step 5 Qualitative Synthesis (Stub)
# =========================================================
with tabs[5]:
    st.subheader(t("tabs_rob_a"))
    st.caption(t("step5_instructions"))
    st.info("Tab 5 (Qualitative Synthesis) content placeholder.")

# =========================================================
# Tab 6: Step 6a Data Extraction (Stub)
# =========================================================
with tabs[6]:
    st.subheader(t("tabs_rob_b"))
    st.caption(t("step6a_instructions"))
    st.info("Tab 6 (Data Extraction) content placeholder.")

# =========================================================
# Tab 7: Step 6b RoB 2.0 (RCT) (Stub)
# =========================================================
with tabs[7]:
    st.subheader(t("tabs_rob_c"))
    st.caption(t("step6b_instructions"))
    st.info("Tab 7 (RoB 2.0) content placeholder.")

# =========================================================
# Tab 8: Step 7 Manuscript Draft (Stub)
# =========================================================
with tabs[8]:
    st.subheader(t("tabs_manuscript"))
    st.info("Tab 8 (Manuscript Draft) content placeholder.")

# =========================================================
# Tab 9: Diagnostics (Stub)
# =========================================================
with tabs[9]:
    st.subheader(t("tabs_diag"))
    st.info("Tab 9 (Diagnostics) content placeholder.")
