import streamlit as st
import zipfile
import io
from typing import List, Dict
from anonymizer import MQXLIFFAnonymizer, load_dictionary_terms
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

st.set_page_config(
    page_title="Anonymizer",
    page_icon="🔒",
    layout="wide",
    initial_sidebar_state="expanded"
)

CUSTOM_CSS = """
<style>
    .main {
        background-color: #ffffff !important;
    }
    .stApp {
        background-color: #ffffff !important;
    }
    .stApp, .stApp p, .stApp span, .stApp label, .stApp div {
        color: #130e45 !important;
    }
    
    /* Global scrollbar styles */
    *::-webkit-scrollbar {
        width: 16px !important;
        height: 16px !important;
    }
    *::-webkit-scrollbar-track {
        background: #d0d3d4 !important;
        border-radius: 8px !important;
    }
    *::-webkit-scrollbar-thumb {
        background: linear-gradient(180deg, #1a5488, #0e7bc0) !important;
        border-radius: 8px !important;
        border: 2px solid #d0d3d4 !important;
    }
    *::-webkit-scrollbar-thumb:hover {
        background: linear-gradient(180deg, #134277, #1a5488) !important;
    }
    
    /* Firefox scrollbar */
    * {
        scrollbar-width: auto;
        scrollbar-color: #1a5488 #d0d3d4;
    }
    [data-testid="stSidebar"] {
        background-color: #e0e3e4 !important;
    }
    [data-testid="stSidebar"] p, [data-testid="stSidebar"] span, [data-testid="stSidebar"] label {
        color: #130e45 !important;
    }
    h1, h2, h3 {
        color: #1a5488 !important;
    }
    .stButton > button {
        background-color: #0e7bc0;
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.5rem 1rem;
        font-weight: 800;
    }
    .stButton > button[kind="primary"] {
        font-weight: 900;
        font-size: 1.1rem;
    }
    .stButton > button:hover {
        background-color: #134277;
    }
    .stat-card {
        background-color: #e0e3e4;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
        border-left: 4px solid #1a5488;
    }
    .stat-number {
        font-size: 2rem;
        font-weight: bold;
        color: #130e45;
    }
    .stat-label {
        color: #5e5f6b;
        font-size: 0.9rem;
    }
    .preview-box {
        background-color: #f8f9fa;
        border: 1px solid #bcbdbe;
        border-radius: 8px;
        padding: 1rem;
        margin: 0.5rem 0;
    }
    .preview-container {
        max-height: 500px;
        overflow-y: auto;
        padding-right: 10px;
        margin: 1rem 0;
        border: 1px solid #e0e3e4;
        border-radius: 8px;
        background-color: #fafafa;
    }
    .preview-container::-webkit-scrollbar {
        width: 18px;
    }
    .preview-container::-webkit-scrollbar-track {
        background: #d0d3d4;
        border-radius: 9px;
        box-shadow: inset 0 0 3px rgba(0,0,0,0.2);
    }
    .preview-container::-webkit-scrollbar-thumb {
        background: linear-gradient(180deg, #1a5488, #0e7bc0);
        border-radius: 9px;
        border: 3px solid #d0d3d4;
        box-shadow: 0 2px 4px rgba(0,0,0,0.3);
    }
    .preview-container::-webkit-scrollbar-thumb:hover {
        background: linear-gradient(180deg, #134277, #1a5488);
    }
    div[data-testid="stExpander"] {
        max-height: 600px;
        overflow-y: auto;
    }
    div[data-testid="stExpander"]::-webkit-scrollbar {
        width: 20px;
    }
    div[data-testid="stExpander"]::-webkit-scrollbar-track {
        background: #c8cbcc;
        border-radius: 10px;
        box-shadow: inset 0 0 4px rgba(0,0,0,0.25);
    }
    div[data-testid="stExpander"]::-webkit-scrollbar-thumb {
        background: linear-gradient(180deg, #1a5488, #0e7bc0);
        border-radius: 10px;
        border: 3px solid #c8cbcc;
        box-shadow: 0 2px 5px rgba(0,0,0,0.35);
    }
    div[data-testid="stExpander"]::-webkit-scrollbar-thumb:hover {
        background: linear-gradient(180deg, #134277, #1a5488);
    }
    .exclude-badge {
        background-color: #dc3545;
        color: white;
        padding: 0.2rem 0.5rem;
        border-radius: 4px;
        font-size: 0.75rem;
        font-weight: bold;
        margin-left: 0.5rem;
    }
    .excluded-segment {
        border: 2px dashed #dc3545 !important;
        background-color: #fff0f0 !important;
    }
    .before-text {
        color: #dc3545;
        background-color: #ffe6e6;
        padding: 0.5rem;
        border-radius: 4px;
    }
    .after-text {
        color: #28a745;
        background-color: #e6ffe6;
        padding: 0.5rem;
        border-radius: 4px;
    }
    .section-header {
        background-color: #7cb4db;
        color: white;
        padding: 0.5rem 1rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #e0e3e4;
        border-left: 4px solid #0e7bc0;
        padding: 1rem;
        border-radius: 0 8px 8px 0;
        margin: 1rem 0;
    }
    div[data-testid="stFileUploader"] {
        background-color: #f8f9fa;
        border: 2px dashed #7cb4db;
        border-radius: 10px;
        padding: 1rem;
    }
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def count_words(text: str) -> int:
    """Count words in text, ignoring XML tags and whitespace."""
    if not text:
        return 0
    import re
    clean = re.sub(r'<[^>]+>', ' ', text)
    clean = re.sub(r'\s+', ' ', clean).strip()
    words = [w for w in clean.split() if len(w) > 0]
    return len(words)


def segment_word_count(preview: dict) -> int:
    """Returns the max word count between source and target."""
    source_words = count_words(preview.get('source_before', ''))
    target_words = count_words(preview.get('target_before', ''))
    return max(source_words, target_words)


def render_stat_card(label: str, value: int, col):
    with col:
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-number">{value}</div>
            <div class="stat-label">{label}</div>
        </div>
        """, unsafe_allow_html=True)


def generate_changes_excel() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Anonymization Report"
    
    header_font = Font(bold=True, color="FFFFFF", size=12)
    header_fill = PatternFill(start_color="1a5488", end_color="1a5488", fill_type="solid")
    title_font = Font(bold=True, size=14, color="1a5488")
    border = Border(
        left=Side(style='thin', color='bcbdbe'),
        right=Side(style='thin', color='bcbdbe'),
        top=Side(style='thin', color='bcbdbe'),
        bottom=Side(style='thin', color='bcbdbe')
    )
    alt_fill = PatternFill(start_color="e8f4fc", end_color="e8f4fc", fill_type="solid")
    total_fill = PatternFill(start_color="d4edda", end_color="d4edda", fill_type="solid")
    total_font = Font(bold=True, size=11)
    
    ws.merge_cells('A1:F1')
    ws['A1'] = "ANONYMIZATION REPORT - MQXLIFF/TMX Anonymizer v6.1"
    ws['A1'].font = title_font
    ws['A1'].alignment = Alignment(horizontal='center')
    
    row = 3
    ws.merge_cells(f'A{row}:F{row}')
    ws[f'A{row}'] = "STATISTICS SUMMARY"
    ws[f'A{row}'].font = Font(bold=True, size=12)
    
    row += 1
    stats_headers = ["File", "Regex CT IDs", "Presidio", "ScispaCy", "Dictionary", "Total"]
    for col, header in enumerate(stats_headers, 1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='center')
    
    row += 1
    grand_total = 0
    for filename, stats in st.session_state.all_stats.items():
        file_total = stats.get("regex_ct", 0) + stats.get("presidio_pii", 0) + stats.get("biomedical", 0) + stats.get("dictionary", 0)
        grand_total += file_total
        data = [filename, stats["regex_ct"], stats["presidio_pii"], stats["biomedical"], stats["dictionary"], file_total]
        for col, value in enumerate(data, 1):
            cell = ws.cell(row=row, column=col, value=value)
            cell.border = border
            if col > 1:
                cell.alignment = Alignment(horizontal='center')
        row += 1
    
    total_data = ["GRAND TOTAL", "", "", "", "", grand_total]
    for col, value in enumerate(total_data, 1):
        cell = ws.cell(row=row, column=col, value=value)
        cell.font = total_font
        cell.fill = total_fill
        cell.border = border
        if col == 1:
            cell.alignment = Alignment(horizontal='left')
        else:
            cell.alignment = Alignment(horizontal='center')
    
    row += 3
    ws.merge_cells(f'A{row}:E{row}')
    ws[f'A{row}'] = "CHANGES DETAIL"
    ws[f'A{row}'].font = Font(bold=True, size=12)
    
    row += 1
    detail_headers = ["File", "Segment", "Type", "Original Text", "Anonymized Text"]
    for col, header in enumerate(detail_headers, 1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='center')
    
    row += 1
    alt_row = False
    filter_short = st.session_state.get('filter_short_segments', False)
    min_words = st.session_state.get('min_words', 5)
    
    for filename, file_previews in st.session_state.previews.items():
        for preview in file_previews:
            if filter_short and segment_word_count(preview) < min_words:
                continue
            
            if preview['source_before'] != preview['source_after']:
                data = [filename, preview['segment'], "Source", preview['source_before'], preview['source_after']]
                for col, value in enumerate(data, 1):
                    cell = ws.cell(row=row, column=col, value=value)
                    cell.border = border
                    if alt_row:
                        cell.fill = alt_fill
                    cell.alignment = Alignment(wrap_text=True, vertical='top')
                row += 1
                alt_row = not alt_row
            if preview['target_before'] != preview['target_after']:
                data = [filename, preview['segment'], "Target", preview['target_before'], preview['target_after']]
                for col, value in enumerate(data, 1):
                    cell = ws.cell(row=row, column=col, value=value)
                    cell.border = border
                    if alt_row:
                        cell.fill = alt_fill
                    cell.alignment = Alignment(wrap_text=True, vertical='top')
                row += 1
                alt_row = not alt_row
    
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 40
    ws.column_dimensions['E'].width = 40
    ws.column_dimensions['F'].width = 10
    
    ws2 = wb.create_sheet(title="Excluded Segments")
    
    ws2.merge_cells('A1:F1')
    ws2['A1'] = "EXCLUDED SEGMENTS - Anonymizer v6.1"
    ws2['A1'].font = title_font
    ws2['A1'].alignment = Alignment(horizontal='center')
    
    ex_row = 3
    
    short_fill = PatternFill(start_color="fff3cd", end_color="fff3cd", fill_type="solid")
    tm_fill = PatternFill(start_color="f8d7da", end_color="f8d7da", fill_type="solid")
    
    replacement_token = st.session_state.get('replacement_token', 'REDACTED')
    threshold = st.session_state.get('exclusion_threshold', 20)
    exclude_enabled = st.session_state.get('exclude_modified_targets', False)
    excluded_segments = st.session_state.get('excluded_segments', {})
    
    has_short = filter_short
    has_tm = exclude_enabled
    
    if has_short:
        ws2.merge_cells(f'A{ex_row}:F{ex_row}')
        ws2[f'A{ex_row}'] = f"SHORT SEGMENTS - Excluded from TM (less than {min_words} words)"
        ws2[f'A{ex_row}'].font = Font(bold=True, size=12, color="856404")
        ws2[f'A{ex_row}'].fill = short_fill
        ex_row += 1
        
        short_headers = ["File", "Segment", "Source (original)", "Target (original)", "Source (anonymized)", "Target (anonymized)"]
        for col, header in enumerate(short_headers, 1):
            cell = ws2.cell(row=ex_row, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border
            cell.alignment = Alignment(horizontal='center')
        ex_row += 1
        
        short_count = 0
        for filename, file_previews in st.session_state.previews.items():
            for preview in file_previews:
                if segment_word_count(preview) < min_words:
                    short_count += 1
                    data = [
                        filename,
                        preview['segment'],
                        preview['source_before'],
                        preview['target_before'],
                        preview['source_after'],
                        preview['target_after'],
                    ]
                    for col, value in enumerate(data, 1):
                        cell = ws2.cell(row=ex_row, column=col, value=value)
                        cell.border = border
                        cell.fill = PatternFill(start_color="fffbe6", end_color="fffbe6", fill_type="solid")
                        cell.alignment = Alignment(wrap_text=True, vertical='top')
                    ex_row += 1
        
        if short_count == 0:
            ws2.cell(row=ex_row, column=1, value="No short segments found.").font = Font(italic=True, color="666666")
            ex_row += 1
        else:
            ws2.merge_cells(f'A{ex_row}:F{ex_row}')
            ws2[f'A{ex_row}'] = f"Total: {short_count} short segments excluded from TM"
            ws2[f'A{ex_row}'].font = total_font
            ws2[f'A{ex_row}'].fill = short_fill
            ex_row += 1
        
        ex_row += 2
    
    if has_tm:
        ws2.merge_cells(f'A{ex_row}:F{ex_row}')
        ws2[f'A{ex_row}'] = f"TM EXCLUDED SEGMENTS (redaction >= {threshold}%)"
        ws2[f'A{ex_row}'].font = Font(bold=True, size=12, color="721c24")
        ws2[f'A{ex_row}'].fill = tm_fill
        ex_row += 1
        
        tm_headers = ["File", "Segment", "Redaction %", "Source (anonymized)", "Target (anonymized)", "Reason"]
        for col, header in enumerate(tm_headers, 1):
            cell = ws2.cell(row=ex_row, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border
            cell.alignment = Alignment(horizontal='center')
        ex_row += 1
        
        tm_count = 0
        for filename, file_previews in st.session_state.previews.items():
            for preview in file_previews:
                has_token = (replacement_token in preview.get('source_after', '') or 
                            replacement_token in preview.get('target_after', ''))
                if not has_token:
                    continue
                if filter_short and segment_word_count(preview) < min_words:
                    continue
                
                segment_key = f"{filename}_{preview['segment']}"
                target_text = preview.get('target_after', '')
                target_pct = 0
                if target_text:
                    words = target_text.split()
                    if words:
                        target_pct = sum(1 for w in words if replacement_token in w) / len(words) * 100
                
                is_excluded = excluded_segments.get(segment_key, True) if target_pct >= threshold else False
                
                if is_excluded:
                    tm_count += 1
                    data = [
                        filename,
                        preview['segment'],
                        f"{target_pct:.1f}%",
                        preview['source_after'],
                        preview['target_after'],
                        f"Redaction >= {threshold}%",
                    ]
                    for col, value in enumerate(data, 1):
                        cell = ws2.cell(row=ex_row, column=col, value=value)
                        cell.border = border
                        cell.fill = PatternFill(start_color="fce4ec", end_color="fce4ec", fill_type="solid")
                        cell.alignment = Alignment(wrap_text=True, vertical='top')
                    ex_row += 1
        
        if tm_count == 0:
            ws2.cell(row=ex_row, column=1, value="No TM-excluded segments found.").font = Font(italic=True, color="666666")
            ex_row += 1
        else:
            ws2.merge_cells(f'A{ex_row}:F{ex_row}')
            ws2[f'A{ex_row}'] = f"Total: {tm_count} segments excluded from TM"
            ws2[f'A{ex_row}'].font = total_font
            ws2[f'A{ex_row}'].fill = tm_fill
            ex_row += 1
    
    if not has_short and not has_tm:
        ws2.cell(row=ex_row, column=1, value="No filters active. No segments were excluded.").font = Font(italic=True, color="666666")
    
    ws2.column_dimensions['A'].width = 40
    ws2.column_dimensions['B'].width = 12
    ws2.column_dimensions['C'].width = 14
    ws2.column_dimensions['D'].width = 45
    ws2.column_dimensions['E'].width = 45
    ws2.column_dimensions['F'].width = 20
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def main():
    st.markdown("""
    <div style="background: linear-gradient(135deg, #1a5488 0%, #0e7bc0 100%); 
                padding: 2rem; 
                border-radius: 12px; 
                margin-bottom: 2rem;
                box-shadow: 0 4px 15px rgba(26, 84, 136, 0.3);">
        <h1 style="color: white !important; margin: 0; font-size: 2.5rem; font-weight: 700; letter-spacing: 8px;">
            Anonymizer <span style="background: rgba(255,255,255,0.2); color: rgba(255,255,255,0.95); font-size: 0.9rem; font-weight: 500; padding: 3px 10px; border-radius: 20px; vertical-align: middle; letter-spacing: 1px; border: 1px solid rgba(255,255,255,0.3);">v6.1</span>
        </h1>
        <p style="color: rgba(255,255,255,0.9) !important; margin: 0.5rem 0 0 0; font-size: 1.1rem;">
            Anonymize bilingual memoQ documents (EN > ES)
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    with st.sidebar:
        st.markdown("### ⚙️ Settings")
        replacement_token = st.text_input("Replacement token", value="REDACTED")
        st.session_state['replacement_token'] = replacement_token
        process_source = True
        process_target = True
        st.markdown("")
        st.markdown("#### Anonymization layers")
        use_regex = st.checkbox("Clinical ID Regex", value=True, help="NCT IDs, EudraCT, Protocol IDs, Subject IDs, etc.")
        use_presidio = st.checkbox("Presidio", value=True, help="Emails, phone numbers, person names, addresses")
        use_biomedical = st.checkbox("ScispaCy", value=True, help="Drugs, pharmaceutical organizations")
        use_dictionary = st.checkbox("Custom dictionary", value=True, help="Manually loaded sensitive terms")
        st.markdown("")
        st.markdown("#### Filters")
        filter_short_segments = st.checkbox("Exclude short segments", value=True, help="Excludes short anonymized segments")
        min_words = st.slider("Min. words", 2, 10, 5, 1) if filter_short_segments else 5
        st.session_state['filter_short_segments'] = filter_short_segments
        st.session_state['min_words'] = min_words
        exclude_modified_targets = st.checkbox("Exclude heavily anonymized", value=True)
        if exclude_modified_targets:
            exclusion_threshold = st.slider("Threshold (%)", 10, 90, 50, 5, help="Segments above X% anonymized excluded")
            exclude_source_too = True
        else:
            exclusion_threshold = 20
            exclude_source_too = False
        st.session_state['exclusion_threshold'] = exclusion_threshold
        st.session_state['exclude_source_too'] = exclude_source_too
        st.session_state['exclude_modified_targets'] = exclude_modified_targets
    
    tab1, tab2, tab3 = st.tabs(["📤 Upload", "📝 Preview", "📥 Download"])
    
    with tab1:
        st.markdown("### Upload MQXLIFF / TMX files")
        
        mqxliff_files = st.file_uploader(
            "Select one or more .mqxliff or .tmx files",
            type=["mqxliff", "tmx"],
            accept_multiple_files=True,
            help="You can upload multiple files for batch processing"
        )
        
        if mqxliff_files:
            st.success(f"✅ {len(mqxliff_files)} file(s) loaded")
            for f in mqxliff_files:
                ext = f.name.rsplit(".", 1)[-1].upper() if "." in f.name else "?"
                st.write(f"- {f.name} ({f.size / 1024:.1f} KB) — {ext}")
        
        st.markdown("---")
        st.markdown("### Sensitive terms dictionary")
        
        st.markdown("""
        <div class="info-box">
            <strong>ℹ️ Terms that will be anonymized:</strong><br>
            These terms will be forcefully anonymized even if not detected automatically.<br>
            One term per line or separated by commas:<br>
            <code>Example</code><br>
            <code>Example project, info@company.com</code><br>
            <code>Brand®, www.example.com</code>
        </div>
        """, unsafe_allow_html=True)
        
        dictionary_file = st.file_uploader(
            "Upload TXT file with sensitive terms (optional)",
            type=["txt"],
            help="Terms not well detected by Presidio or ScispaCy"
        )
        
        dictionary_terms = set()
        if dictionary_file:
            dict_raw = dictionary_file.read()
            if dict_raw[:3] == b'\xef\xbb\xbf':
                dict_raw = dict_raw[3:]
            elif dict_raw[:2] in (b'\xff\xfe', b'\xfe\xff'):
                dict_raw = dict_raw.decode('utf-16').encode('utf-8')
            try:
                content = dict_raw.decode("utf-8")
            except UnicodeDecodeError:
                content = dict_raw.decode("latin-1")
            dictionary_terms = load_dictionary_terms(content)
            st.success(f"✅ {len(dictionary_terms)} unique terms loaded")
            
            with st.expander("View loaded terms"):
                for term in sorted(dictionary_terms):
                    st.write(f"- {term}")
        
        st.markdown("### Protected terms (whitelist)")
        
        st.markdown("""
        <div class="info-box">
            <strong>ℹ️ Terms that will NOT be anonymized:</strong><br>
            These terms will be preserved even if detected.<br>
            One term per line or separated by commas.
        </div>
        """, unsafe_allow_html=True)
        
        whitelist_file = st.file_uploader(
            "Upload TXT file with protected terms (optional)",
            type=["txt"],
            help="Terms that should never be anonymized",
            key="whitelist_uploader"
        )
        
        whitelist_terms = set()
        if whitelist_file:
            wl_raw = whitelist_file.read()
            if wl_raw[:3] == b'\xef\xbb\xbf':
                wl_raw = wl_raw[3:]
            elif wl_raw[:2] in (b'\xff\xfe', b'\xfe\xff'):
                wl_raw = wl_raw.decode('utf-16').encode('utf-8')
            try:
                wl_content = wl_raw.decode("utf-8")
            except UnicodeDecodeError:
                wl_content = wl_raw.decode("latin-1")
            whitelist_terms = load_dictionary_terms(wl_content)
            st.success(f"🛡️ {len(whitelist_terms)} protected terms loaded")
            
            with st.expander("View protected terms"):
                for term in sorted(whitelist_terms):
                    st.write(f"- {term}")
        
        if mqxliff_files:
            st.markdown("---")
            if st.button("🚀 Process files", type="primary", use_container_width=True):
                process_files(
                    mqxliff_files, replacement_token, process_source, process_target,
                    use_regex, use_presidio, use_biomedical, use_dictionary, dictionary_terms,
                    whitelist_terms
                )
    
    with tab2:
        if "previews" in st.session_state and st.session_state.previews:
            st.markdown("### Changes preview")
            
            search_term = st.text_input(
                "🔍 Search in preview",
                placeholder="Type to filter changes...",
                key="preview_search"
            )
            
            def calc_redaction_pct(text: str, token: str) -> float:
                if not text:
                    return 0
                words = text.split()
                if not words:
                    return 0
                redacted_count = sum(1 for w in words if token in w)
                return (redacted_count / len(words)) * 100
            
            replacement_token = st.session_state.get('replacement_token', 'REDACTED')
            threshold = st.session_state.get('exclusion_threshold', 20)
            exclude_enabled = st.session_state.get('exclude_modified_targets', False)
            exclude_source = st.session_state.get('exclude_source_too', False)
            
            if 'excluded_segments' not in st.session_state:
                st.session_state['excluded_segments'] = {}
            if 'no_anon_segments' not in st.session_state:
                st.session_state['no_anon_segments'] = {}
            
            filter_short = st.session_state.get('filter_short_segments', False)
            min_words = st.session_state.get('min_words', 5)
            
            total_matches = 0
            candidates_count = 0
            short_segments_count = 0
            
            for filename, file_previews in st.session_state.previews.items():
                filtered_previews = []
                for preview in file_previews:
                    has_token = (replacement_token in preview.get('source_after', '') or 
                                replacement_token in preview.get('target_after', ''))
                    if not has_token:
                        continue
                    
                    if filter_short and segment_word_count(preview) < min_words:
                        short_segments_count += 1
                        continue
                    
                    if search_term:
                        search_lower = search_term.lower()
                        if (search_lower in preview['source_before'].lower() or
                            search_lower in preview['source_after'].lower() or
                            search_lower in preview['target_before'].lower() or
                            search_lower in preview['target_after'].lower()):
                            filtered_previews.append(preview)
                    else:
                        filtered_previews.append(preview)
                
                if filtered_previews:
                    total_matches += len(filtered_previews)
                    with st.expander(f"📄 {filename} ({len(filtered_previews)} affected segments)", expanded=True):
                        for preview in filtered_previews:
                            target_pct = calc_redaction_pct(preview.get('target_after', ''), replacement_token)
                            source_pct = calc_redaction_pct(preview.get('source_after', ''), replacement_token)
                            is_candidate = target_pct >= threshold
                            
                            if is_candidate:
                                candidates_count += 1
                            
                            segment_key = f"{filename}_{preview['segment']}"
                            
                            has_changes = (preview['source_before'] != preview['source_after'] or 
                                         preview['target_before'] != preview['target_after'])
                            
                            col_header, col_no_anon, col_exclude = st.columns([3, 1, 1])
                            with col_header:
                                header_placeholder = st.empty()
                            
                            with col_no_anon:
                                if has_changes:
                                    no_anon_val = st.session_state['no_anon_segments'].get(segment_key, False)
                                    no_anon = st.checkbox(
                                        "Skip",
                                        value=no_anon_val,
                                        key=f"noanon_{segment_key}",
                                        help="Keep original text without anonymizing"
                                    )
                                    st.session_state['no_anon_segments'][segment_key] = no_anon
                            
                            with col_exclude:
                                if is_candidate and exclude_enabled:
                                    default_val = st.session_state['excluded_segments'].get(segment_key, True)
                                    exclude_this = st.checkbox(
                                        "Exclude TM",
                                        value=default_val,
                                        key=f"excl_{segment_key}"
                                    )
                                    st.session_state['excluded_segments'][segment_key] = exclude_this
                            
                            is_no_anon = st.session_state['no_anon_segments'].get(segment_key, False)
                            
                            segment_header = f"**Segment {preview['segment']}**"
                            if is_no_anon:
                                segment_header += ' <span style="background:#28a745;color:white;padding:2px 6px;border-radius:4px;font-size:0.75rem;">✓ Not anonymized</span>'
                            elif is_candidate and exclude_enabled:
                                segment_header += f' <span class="exclude-badge">⚠️ {target_pct:.0f}% → exclusion</span>'
                            elif target_pct > 0:
                                segment_header += f' <small style="color:#666">({target_pct:.0f}%)</small>'
                            header_placeholder.markdown(segment_header, unsafe_allow_html=True)
                            
                            if preview['source_before'] != preview['source_after']:
                                col1, col2 = st.columns(2)
                                with col1:
                                    st.markdown("**Source - Before:**")
                                    st.markdown(f'<div class="before-text">{preview["source_before"]}</div>', 
                                              unsafe_allow_html=True)
                                with col2:
                                    if is_no_anon:
                                        st.markdown("**Source - After** _(no changes)_:")
                                        st.markdown(f'<div class="before-text" style="border-color:#28a745;">{preview["source_before"]}</div>', 
                                                  unsafe_allow_html=True)
                                    else:
                                        label_src = "**Source - After:**"
                                        show_excluded = is_candidate and exclude_enabled and exclude_source and st.session_state['excluded_segments'].get(segment_key, True)
                                        if show_excluded:
                                            label_src = "**Source - After** _(will be excluded)_:"
                                        st.markdown(label_src)
                                        css_class = "after-text excluded-segment" if show_excluded else "after-text"
                                        st.markdown(f'<div class="{css_class}">{preview["source_after"]}</div>', 
                                                  unsafe_allow_html=True)
                            
                            if preview['target_before'] != preview['target_after']:
                                col1, col2 = st.columns(2)
                                with col1:
                                    st.markdown("**Target - Before:**")
                                    st.markdown(f'<div class="before-text">{preview["target_before"]}</div>', 
                                              unsafe_allow_html=True)
                                with col2:
                                    if is_no_anon:
                                        st.markdown("**Target - After** _(no changes)_:")
                                        st.markdown(f'<div class="before-text" style="border-color:#28a745;">{preview["target_before"]}</div>', 
                                                  unsafe_allow_html=True)
                                    else:
                                        label = "**Target - After:**"
                                        show_excluded = is_candidate and exclude_enabled and st.session_state['excluded_segments'].get(segment_key, True)
                                        if show_excluded:
                                            label = "**Target - After** _(will be excluded)_:"
                                        st.markdown(label)
                                        css_class = "after-text excluded-segment" if show_excluded else "after-text"
                                        st.markdown(f'<div class="{css_class}">{preview["target_after"]}</div>', 
                                                  unsafe_allow_html=True)
                            
                            st.markdown("---")
            
            if filter_short and short_segments_count > 0:
                st.info(f"✂️ {short_segments_count} short segments excluded from TM (less than {min_words} words)")
            
            if exclude_enabled and candidates_count > 0:
                excluded_count = sum(1 for v in st.session_state['excluded_segments'].values() if v)
                st.info(f"📊 {excluded_count} of {candidates_count} candidate segments will be excluded from TM (threshold: {threshold}%)")
            
            if search_term and total_matches == 0:
                st.warning(f"No changes found containing '{search_term}'")
        else:
            st.info("Process the files to see the changes preview")
    
    with tab3:
        if "results" in st.session_state and st.session_state.results:
            st.markdown("### Anonymization statistics")
            
            total_stats = {"regex_ct": 0, "presidio_pii": 0, "biomedical": 0, "dictionary": 0}
            for stats in st.session_state.all_stats.values():
                for key in total_stats:
                    total_stats[key] += stats[key]
            
            col1, col2, col3, col4 = st.columns(4)
            render_stat_card("Regex CT IDs", total_stats["regex_ct"], col1)
            render_stat_card("Presidio", total_stats["presidio_pii"], col2)
            render_stat_card("Biomedical", total_stats["biomedical"], col3)
            render_stat_card("Dictionary", total_stats["dictionary"], col4)
            
            st.markdown("---")
            st.markdown("### Download anonymized files")
            
            exclude_modified_targets = st.session_state.get('exclude_modified_targets', False)
            exclusion_threshold = st.session_state.get('exclusion_threshold', 20)
            exclude_source_too = st.session_state.get('exclude_source_too', False)
            excluded_segments = st.session_state.get('excluded_segments', {})
            
            if exclude_modified_targets:
                replacement_token = st.session_state.get('replacement_token', 'REDACTED')
                
                def calc_pct(text, token):
                    if not text:
                        return 0
                    words = text.split()
                    if not words:
                        return 0
                    return sum(1 for w in words if token in w) / len(words) * 100
                
                excluded_count = 0
                for filename, file_previews in st.session_state.get('previews', {}).items():
                    for preview in file_previews:
                        target_pct = calc_pct(preview.get('target_after', ''), replacement_token)
                        if target_pct >= exclusion_threshold:
                            segment_key = f"{filename}_{preview['segment']}"
                            if excluded_segments.get(segment_key, True):
                                excluded_count += 1
                
                if excluded_count > 0:
                    st.markdown(f"""
                    <div style="background-color: #fff3cd; border: 1px solid #ffc107; border-radius: 8px; padding: 1rem; margin: 0.5rem 0;">
                        <strong>⚠️ {excluded_count} segments will be excluded from TM</strong><br>
                        <small>Threshold: {exclusion_threshold}% | You can uncheck them in the Preview tab</small>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.info("No segments marked for exclusion")
            
            def get_short_segment_keys(filename: str, file_previews: list, min_words: int) -> set:
                """Gets the keys of short segments that should be restored."""
                short_keys = set()
                for preview in file_previews:
                    if segment_word_count(preview) < min_words:
                        short_keys.add(f"{filename}_{preview['segment']}")
                return short_keys
            
            def apply_no_anon_segments(anon_content: bytes, orig_content: bytes, filename: str, no_anon_segs: dict) -> bytes:
                if not no_anon_segs:
                    return anon_content
                
                from lxml import etree
                is_tmx = filename.lower().endswith(".tmx")
                
                try:
                    anon_tree = etree.fromstring(anon_content)
                    orig_tree = etree.fromstring(orig_content)
                    
                    if is_tmx:
                        anon_units = anon_tree.xpath('//tu')
                        orig_units = orig_tree.xpath('//tu')
                    else:
                        nsmap = anon_tree.nsmap
                        default_ns = nsmap.get(None, '')
                        
                        if default_ns:
                            ns = {'x': default_ns}
                            anon_units = anon_tree.xpath('//x:trans-unit', namespaces=ns)
                            orig_units = orig_tree.xpath('//x:trans-unit', namespaces=ns)
                        else:
                            anon_units = anon_tree.xpath('//trans-unit')
                            orig_units = orig_tree.xpath('//trans-unit')
                    
                    replacement_token = st.session_state.get('replacement_token', 'REDACTED')
                    
                    for idx, (anon_tu, orig_tu) in enumerate(zip(anon_units, orig_units)):
                        segment_num = idx + 1
                        segment_key = f"{filename}_{segment_num}"
                        
                        should_restore_only = no_anon_segs.get(segment_key, False)
                        
                        if should_restore_only:
                            if is_tmx:
                                anon_tuvs = anon_tu.xpath('tuv')
                                orig_tuvs = orig_tu.xpath('tuv')
                                for anon_tuv, orig_tuv in zip(anon_tuvs, orig_tuvs):
                                    anon_segs = anon_tuv.xpath('seg')
                                    orig_segs = orig_tuv.xpath('seg')
                                    for anon_seg, orig_seg in zip(anon_segs, orig_segs):
                                        anon_seg.getparent().replace(anon_seg, orig_seg)
                            else:
                                if default_ns:
                                    anon_sources = anon_tu.xpath('.//x:source', namespaces=ns)
                                    orig_sources = orig_tu.xpath('.//x:source', namespaces=ns)
                                    anon_targets = anon_tu.xpath('.//x:target', namespaces=ns)
                                    orig_targets = orig_tu.xpath('.//x:target', namespaces=ns)
                                else:
                                    anon_sources = anon_tu.xpath('.//source')
                                    orig_sources = orig_tu.xpath('.//source')
                                    anon_targets = anon_tu.xpath('.//target')
                                    orig_targets = orig_tu.xpath('.//target')
                                
                                for anon_src, orig_src in zip(anon_sources, orig_sources):
                                    anon_src.getparent().replace(anon_src, orig_src)
                                
                                for anon_tgt, orig_tgt in zip(anon_targets, orig_targets):
                                    anon_tgt.getparent().replace(anon_tgt, orig_tgt)
                    
                    return etree.tostring(anon_tree, encoding='utf-8', xml_declaration=True)
                except:
                    return anon_content
            
            def _clear_element(elem):
                for child in list(elem):
                    elem.remove(child)
                elem.text = None
                elem.tail = None

            def prepare_download_content(content: bytes, filename: str, exclude_targets: bool, threshold: float, exclude_source: bool, excluded_segs: dict, short_segs: set = None) -> bytes:
                if short_segs is None:
                    short_segs = set()
                
                if not exclude_targets and not short_segs:
                    return content
                
                from lxml import etree
                is_tmx = filename.lower().endswith(".tmx")
                
                try:
                    tree = etree.fromstring(content)
                    
                    if is_tmx:
                        trans_units = tree.xpath('//tu')
                    else:
                        nsmap = tree.nsmap
                        default_ns = nsmap.get(None, '')
                        if default_ns:
                            ns = {'x': default_ns}
                            trans_units = tree.xpath('//x:trans-unit', namespaces=ns)
                        else:
                            trans_units = tree.xpath('//trans-unit')
                    
                    replacement_token = st.session_state.get('replacement_token', 'REDACTED')
                    
                    for idx, tu in enumerate(trans_units):
                        segment_num = idx + 1
                        segment_key = f"{filename}_{segment_num}"
                        
                        is_short = segment_key in short_segs
                        
                        if is_short:
                            if is_tmx:
                                for tuv in tu.xpath('tuv'):
                                    for seg in tuv.xpath('seg'):
                                        _clear_element(seg)
                            else:
                                if default_ns:
                                    targets = tu.xpath('.//x:target', namespaces=ns)
                                    sources = tu.xpath('.//x:source', namespaces=ns)
                                else:
                                    targets = tu.xpath('.//target')
                                    sources = tu.xpath('.//source')
                                for elem in sources + targets:
                                    _clear_element(elem)
                            continue
                        
                        if not exclude_targets:
                            continue
                        
                        should_exclude_this = excluded_segs.get(segment_key, None)
                        if should_exclude_this is False:
                            continue
                        
                        if is_tmx:
                            target_tuvs = []
                            source_tuvs = []
                            for tuv in tu.xpath('tuv'):
                                tuv_lang = tuv.get("{http://www.w3.org/XML/1998/namespace}lang", tuv.get("lang", ""))
                                if tuv_lang.lower().startswith("es"):
                                    target_tuvs.extend(tuv.xpath('seg'))
                                elif tuv_lang.lower().startswith("en") and exclude_source:
                                    source_tuvs.extend(tuv.xpath('seg'))
                            
                            for seg in target_tuvs:
                                text_content = ''.join(seg.itertext())
                                words = text_content.split()
                                if replacement_token in text_content and words:
                                    redacted_pct = sum(1 for w in words if replacement_token in w) / len(words) * 100
                                    if redacted_pct >= threshold:
                                        _clear_element(seg)
                                        if exclude_source:
                                            for src_seg in source_tuvs:
                                                _clear_element(src_seg)
                        else:
                            if default_ns:
                                targets = tu.xpath('.//x:target', namespaces=ns)
                                sources = tu.xpath('.//x:source', namespaces=ns) if exclude_source else []
                            else:
                                targets = tu.xpath('.//target')
                                sources = tu.xpath('.//source') if exclude_source else []
                            
                            for target in targets:
                                text_content = ''.join(target.itertext())
                                words = text_content.split()
                                if replacement_token in text_content and words:
                                    redacted_pct = sum(1 for w in words if replacement_token in w) / len(words) * 100
                                    
                                    if redacted_pct >= threshold:
                                        _clear_element(target)
                                        
                                        if exclude_source:
                                            for source in sources:
                                                _clear_element(source)
                    
                    return etree.tostring(tree, encoding='utf-8', xml_declaration=True)
                except:
                    return content
            
            no_anon_segments = st.session_state.get('no_anon_segments', {})
            originals = st.session_state.get('originals', {})
            
            no_anon_count = sum(1 for v in no_anon_segments.values() if v)
            if no_anon_count > 0:
                st.markdown(f"""
                <div style="background-color: #d4edda; border: 1px solid #28a745; border-radius: 8px; padding: 1rem; margin: 0.5rem 0;">
                    <strong>✓ {no_anon_count} segments will keep their original text</strong><br>
                    <small>Marked as "Skip" in the Preview tab</small>
                </div>
                """, unsafe_allow_html=True)
            
            filter_short = st.session_state.get('filter_short_segments', False)
            min_words = st.session_state.get('min_words', 5)
            previews = st.session_state.get('previews', {})
            
            if filter_short:
                short_count = sum(
                    1 for file_previews in previews.values()
                    for preview in file_previews
                    if segment_word_count(preview) < min_words
                )
                if short_count > 0:
                    st.markdown(f"""
                    <div style="background-color: #e8f4fc; border: 1px solid #0e7bc0; border-radius: 8px; padding: 1rem; margin: 0.5rem 0;">
                        <strong>✂️ {short_count} short segments excluded from TM</strong><br>
                        <small>Less than {min_words} words → anonymized but excluded from TM update</small>
                    </div>
                    """, unsafe_allow_html=True)
            
            col_mqxliff, col_excel = st.columns(2)
            
            with col_mqxliff:
                if len(st.session_state.results) == 1:
                    filename, content = list(st.session_state.results.items())[0]
                    orig_content = originals.get(filename, content)
                    short_segs = get_short_segment_keys(filename, previews.get(filename, []), min_words) if filter_short else set()
                    content = apply_no_anon_segments(content, orig_content, filename, no_anon_segments)
                    download_content = prepare_download_content(content, filename, exclude_modified_targets, exclusion_threshold, exclude_source_too, excluded_segments, short_segs)
                    file_ext = filename.rsplit(".", 1)[-1].upper() if "." in filename else "XML"
                    st.download_button(
                        label=f"📥 Download {file_ext}",
                        data=download_content,
                        file_name=filename,
                        mime="application/xml",
                        use_container_width=True
                    )
                else:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                        for filename, content in st.session_state.results.items():
                            orig_content = originals.get(filename, content)
                            short_segs = get_short_segment_keys(filename, previews.get(filename, []), min_words) if filter_short else set()
                            content = apply_no_anon_segments(content, orig_content, filename, no_anon_segments)
                            download_content = prepare_download_content(content, filename, exclude_modified_targets, exclusion_threshold, exclude_source_too, excluded_segments, short_segs)
                            zf.writestr(filename, download_content)
                    
                    st.download_button(
                        label=f"📥 Download ZIP ({len(st.session_state.results)} files)",
                        data=zip_buffer.getvalue(),
                        file_name="anonymized_files.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
            
            with col_excel:
                excel_data = generate_changes_excel()
                st.download_button(
                    label="📊 Download changes report (Excel)",
                    data=excel_data,
                    file_name="anonymization_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.info("Process the files to download them")


def process_files(files, replacement_token, process_source, process_target,
                  use_regex, use_presidio, use_biomedical, use_dictionary, dictionary_terms,
                  whitelist_terms=None):
    
    anonymizer = MQXLIFFAnonymizer(replacement_token=replacement_token)
    
    results = {}
    originals = {}
    all_stats = {}
    previews = {}
    
    progress_bar = st.progress(0, text="Processing files...")
    
    for i, file in enumerate(files):
        progress_bar.progress((i + 1) / len(files), text=f"Processing {file.name}...")
        
        try:
            content = file.read()
            originals[file.name] = content
            
            is_tmx = file.name.lower().endswith(".tmx")
            
            if is_tmx:
                result_xml, stats, file_previews = anonymizer.anonymize_tmx(
                    xml_content=content,
                    process_source=process_source,
                    process_target=process_target,
                    use_regex=use_regex,
                    use_presidio=use_presidio,
                    use_biomedical=use_biomedical,
                    use_dictionary=use_dictionary,
                    dictionary_terms=dictionary_terms,
                    whitelist_terms=whitelist_terms
                )
            else:
                result_xml, stats, file_previews = anonymizer.anonymize_mqxliff(
                    xml_content=content,
                    process_source=process_source,
                    process_target=process_target,
                    use_regex=use_regex,
                    use_presidio=use_presidio,
                    use_biomedical=use_biomedical,
                    use_dictionary=use_dictionary,
                    dictionary_terms=dictionary_terms,
                    whitelist_terms=whitelist_terms
                )
            
            results[file.name] = result_xml
            all_stats[file.name] = stats
            previews[file.name] = file_previews
            
        except Exception as e:
            st.error(f"Error processing {file.name}: {str(e)}")
    
    progress_bar.empty()
    
    st.session_state.results = results
    st.session_state.originals = originals
    st.session_state.all_stats = all_stats
    st.session_state.previews = previews
    
    st.success(f"✅ Successfully processed {len(results)} file(s)")
    st.info("Go to the **Preview** and **Download** tabs to see the results")


if __name__ == "__main__":
    main()
