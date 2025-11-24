###############################################
#  CPM REFERENCE CHECK APP - v4.0
###############################################

import os
import re
import zipfile
from io import BytesIO
from datetime import datetime
from xml.etree import ElementTree as ET

import streamlit as st
import pandas as pd
import pdfplumber

# =====================================================
#  GLOBAL CONSTANTS
# =====================================================

CPM_DIR = "cpms"
APP_VERSION = "4.0 (Cloud)"

# These are the Weight Management IDs that should be replaced by "WM Bundle"
WM_BUNDLE_IDS = {'18-C', '1190-C', '794-C', '1227-C', '4774-C', '250-C', '6192-C'}

# ID pattern that handles spaces around hyphen
ID_PATTERN_SPACED = re.compile(r"\bC?(\d{1,5})\s*-\s*([A-Z]{1,3})\b")

# XML namespaces for DOCX/DOCM parsing
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml'
}

# =====================================================
#  HELPER FUNCTIONS
# =====================================================

def ensure_cpm_dir():
    os.makedirs(CPM_DIR, exist_ok=True)

def make_new_cpm_filename(sfc, original_name, ext):
    name_without_ext = original_name.rsplit(".", 1)[0]
    return f"{sfc}__{name_without_ext}.{ext}"

def list_cpm_files():
    ensure_cpm_dir()
    out = []
    for fn in os.listdir(CPM_DIR):
        if not fn.lower().endswith((".pdf", ".doc", ".docx", ".docm")):
            continue
        p = os.path.join(CPM_DIR, fn)
        base = fn.rsplit(".", 1)[0]
        ext = fn.rsplit(".", 1)[1].lower()
        parts = base.split("__", 1)
        sfc = parts[0]
        label = parts[1] if len(parts) > 1 else ""
        out.append({
            "sfc": sfc, "label": label, "filename": fn,
            "extension": ext, "path": p,
            "modified": datetime.fromtimestamp(os.path.getmtime(p)),
        })
    out.sort(key=lambda x: (x["sfc"], x["label"], x["modified"]))
    return out

def list_existing_sfc_numbers():
    return sorted({f["sfc"] for f in list_cpm_files()}, key=lambda x: int(x) if x.isdigit() else x)

def normalize_id(match_tuple):
    """Convert spaced ID to normalized form (e.g., ('123', 'A') -> '123-A')"""
    return f"{match_tuple[0]}-{match_tuple[1]}"

# =====================================================
#  XML PARSING HELPERS
# =====================================================

def load_xml_from_doc(path):
    """Load and parse XML from DOCX/DOCM file"""
    with zipfile.ZipFile(path, "r") as z:
        xml_content = z.read("word/document.xml").decode("utf-8")
    return ET.fromstring(xml_content)

def get_text_from_element(elem):
    """Extract all text from an XML element and its children"""
    texts = []
    for t in elem.findall('.//w:t', NAMESPACES):
        if t.text:
            texts.append(t.text)
    return ' '.join(texts)

# =====================================================
#  SPECIALTY/SGM DETECTION
# =====================================================

def is_specialty_only(text):
    """
    Returns True only if 'specialty' appears WITHOUT 'non-specialty' prefix.
    Non-Specialty tables should NOT be skipped.
    """
    text_lower = text.lower()
    # Remove all "non-specialty" variants first
    cleaned = re.sub(r'non[\s-]*specialty', '', text_lower)
    # Now check if 'specialty' still exists
    return 'specialty' in cleaned

# =====================================================
#  ACTION/CHECKBOX DETECTION
# =====================================================

def get_action_status_from_text(text):
    """
    Check for Add/Delete/Change action status based on checkbox symbols.
    Returns: 'add', 'delete', 'change', or None
    """
    parts = re.split(r'(☒|☐)', text)
    
    for i in range(1, len(parts), 2):
        if i+1 < len(parts):
            checkbox = parts[i]
            following = parts[i+1].lower().strip()
            if checkbox == '☒':
                if following.startswith('add'):
                    return 'add'
                elif following.startswith('delete'):
                    return 'delete'
                elif following.startswith('change'):
                    return 'change'
    return None

def is_reference_text(text, id_start, id_end):
    """
    Check if an ID is mentioned in reference/instructional text
    rather than being a selectable item.
    
    Patterns detected:
    - "Select ... ID in the ... section above"
    - "see ... above"
    - "in place currently (ID, ID, ID)"
    - "must be deselected"
    - "Criteria Drugs: Ref #ID:" (with colon after = definition, not selection)
    """
    before = text[:id_start].lower()
    after = text[id_end:].lower()
    
    # Pattern 1: "Select ... ID in the ... section above"
    has_select_before = 'select' in before[-100:]
    has_section_after = 'section' in after[:100] or 'above' in after[:80]
    if has_select_before and has_section_after:
        return True
    
    # Pattern 2: "see ... above"
    if re.search(r'see\s+.{0,50}above', before[-100:]):
        return True
    
    # Pattern 3: "in place currently (ID, ID, ID)"
    if re.search(r'in\s+place\s+currently\s*\([^)]*$', before[-150:]):
        return True
    
    # Pattern 4: "should select" + section after
    if 'should select' in before[-150:] and has_section_after:
        return True
    
    # Pattern 5: "must be deselected"
    if 'must be deselected' in before[-100:]:
        return True
    
    # Pattern 6: "Criteria Drugs: Ref #ID:" - reference to criteria details
    # Only skip if there's a colon AFTER the ID (meaning it's defining that ID)
    if 'criteria drugs' in before[-50:] and ':' in after[:10]:
        return True
    
    return False

def is_auto_included_table(text):
    """
    Check if table has patterns indicating IDs should be included by default.
    Examples: "automatically added", "the standard ... program"
    """
    text_lower = text.lower()
    patterns = [
        'automatically added',
        'criteria are automatically added',
        'the standard',
        'program is inclusive',
    ]
    return any(p in text_lower for p in patterns)

def is_wm_bundle_row(text):
    """Check if this is the Weight Management row specifically"""
    text_lower = text.lower()
    if 'weight management' in text_lower:
        if '☒' in text or '☐' in text:
            if 'add' in text_lower and ('delete' in text_lower or 'change' in text_lower):
                return True
    return False

# =====================================================
#  ID EXTRACTION FUNCTIONS
# =====================================================

def extract_ids_with_checkbox_awareness(text, default_include=False, skip_wm_ids=False):
    """
    Extract IDs from text, respecting individual checkboxes where present.
    
    Logic:
    1. Find all IDs and checkbox positions
    2. For each ID, find the controlling checkbox (if any)
    3. If checkbox is ☒, include the ID
    4. If no checkbox controls this ID, use default_include
    
    Args:
        text: The text to extract IDs from
        default_include: Whether to include IDs without individual checkboxes
        skip_wm_ids: Whether to skip Weight Management Bundle IDs
        
    Returns:
        tuple: (list of extracted IDs, bool indicating if WM Bundle IDs were found)
    """
    result_ids = []
    wm_bundle_found = False
    
    # Find all IDs with positions
    all_ids = []
    for m in ID_PATTERN_SPACED.finditer(text):
        normalized = normalize_id(m.groups())
        all_ids.append({
            'id': normalized,
            'start': m.start(),
            'end': m.end()
        })
    
    if not all_ids:
        return [], False
    
    # Find all checkbox positions
    checkboxes = []
    for m in re.finditer(r'(☒|☐)', text):
        checkboxes.append({
            'symbol': m.group(1),
            'pos': m.start(),
            'checked': m.group(1) == '☒'
        })
    
    for id_info in all_ids:
        id_start = id_info['start']
        id_end = id_info['end']
        the_id = id_info['id']
        
        # Skip WM Bundle IDs if requested (they'll be replaced by "WM Bundle")
        if skip_wm_ids and the_id in WM_BUNDLE_IDS:
            wm_bundle_found = True
            continue
        
        # Skip if this ID is in reference text
        if is_reference_text(text, id_start, id_end):
            continue
        
        # Find the most recent checkbox before this ID
        controlling_checkbox = None
        for cb in reversed(checkboxes):
            if cb['pos'] < id_start:
                # Check if this is an action checkbox (Add/Delete/Change)
                between_text = text[cb['pos']:id_start].lower()
                action_match = re.search(r'^\s*(add|delete|change)\b', between_text[1:])
                if action_match:
                    # This is an action checkbox, not an ID checkbox
                    controlling_checkbox = None
                    break
                else:
                    # This checkbox controls this ID
                    controlling_checkbox = cb
                    break
        
        if controlling_checkbox is not None:
            if controlling_checkbox['checked']:
                result_ids.append(the_id)
        else:
            if default_include:
                result_ids.append(the_id)
    
    return result_ids, wm_bundle_found

def extract_ids_generic_step_therapy(text):
    """
    Extract IDs from Generic Step Therapy tables.
    These use format: ☒ ☐ ... Ref# XXX-X
    First checkbox = Add, Second = Delete
    """
    result_ids = []
    cb_pairs = re.finditer(r'(☒|☐)\s*(☒|☐)', text)
    
    for match in cb_pairs:
        first_cb = match.group(1)
        pos_after = match.end()
        
        if first_cb == '☒':  # Add is selected
            remaining = text[pos_after:]
            ref_match = re.search(r'Ref#?\s*(\d{1,5})\s*-\s*([A-Z]{1,3})\b', remaining, re.IGNORECASE)
            if ref_match:
                normalized = normalize_id(ref_match.groups())
                result_ids.append(normalized)
    
    return result_ids

def extract_ids_auto_included_table(text):
    """
    Extract IDs from tables where IDs are included by default.
    (e.g., "automatically added" tables)
    """
    result_ids = []
    
    for m in ID_PATTERN_SPACED.finditer(text):
        normalized = normalize_id(m.groups())
        id_start = m.start()
        id_end = m.end()
        
        # Skip reference text
        if is_reference_text(text, id_start, id_end):
            continue
        
        result_ids.append(normalized)
    
    return result_ids

# =====================================================
#  MAIN DOCX/DOCM EXTRACTION
# =====================================================

def extract_ids_from_doc_tables(path, debug=False):
    """
    Extract Criteria IDs from DOCX/DOCM tables using proper XML parsing.
    
    Handles:
    - Standard tables with checkbox symbols
    - WM Bundle rows (outputs "WM Bundle" instead of individual WM IDs)
    - Generic Step Therapy tables (☒ ☐ format)
    - Auto-included tables (no checkboxes, IDs included by default)
    - Non-Specialty tables (not skipped, unlike Specialty/SGM tables)
    - Per-ID checkboxes within rows
    - Reference text (instructional text mentioning IDs that shouldn't be extracted)
    """
    collected_ids = []
    wm_bundle_added = False
    
    try:
        root = load_xml_from_doc(path)
        all_tables = root.findall('.//w:tbl', NAMESPACES)
        
        if debug:
            st.write(f"Found {len(all_tables)} tables")
        
        for tidx, tbl in enumerate(all_tables):
            tbl_text = get_text_from_element(tbl)
            tbl_lower = tbl_text.lower()
            
            # Skip true Specialty/SGM tables (but NOT Non-Specialty)
            if is_specialty_only(tbl_text) or 'sgm' in tbl_lower:
                if debug:
                    st.write(f"Table {tidx}: SKIPPED (specialty/SGM)")
                continue
            
            # Check for Generic Step Therapy table
            is_gstp = 'generic step therapy' in tbl_lower
            if is_gstp:
                if debug:
                    st.write(f"Table {tidx}: Generic Step Therapy")
                rows = tbl.findall('.//w:tr', NAMESPACES)
                for row in rows:
                    row_text = get_text_from_element(row)
                    gstp_ids = extract_ids_generic_step_therapy(row_text)
                    collected_ids.extend(gstp_ids)
                continue
            
            # Check for auto-included table (no checkboxes, IDs included by default)
            is_auto = is_auto_included_table(tbl_text)
            has_checkboxes = '☒' in tbl_text or '☐' in tbl_text
            
            if is_auto and not has_checkboxes:
                if debug:
                    st.write(f"Table {tidx}: AUTO-INCLUDED")
                auto_ids = extract_ids_auto_included_table(tbl_text)
                collected_ids.extend(auto_ids)
                continue
            
            # Standard processing for tables with checkboxes
            rows = tbl.findall('.//w:tr', NAMESPACES)
            current_action = None
            
            for ridx, row in enumerate(rows):
                row_text = get_text_from_element(row)
                
                # Skip specialty-only rows
                if is_specialty_only(row_text) or 'sgm' in row_text.lower():
                    continue
                
                # Check for action status in this row
                row_action = get_action_status_from_text(row_text)
                if row_action:
                    current_action = row_action
                
                # Check if this is the WM Bundle row
                is_wm_row = is_wm_bundle_row(row_text)
                
                # Extract IDs - default_include is True if current action is 'add'
                default_include = (current_action == 'add')
                row_ids, found_wm = extract_ids_with_checkbox_awareness(
                    row_text, 
                    default_include, 
                    skip_wm_ids=is_wm_row
                )
                
                # If we found WM Bundle IDs in a WM row with Add checked, add "WM Bundle"
                if found_wm and current_action == 'add' and not wm_bundle_added:
                    collected_ids.append("WM Bundle")
                    wm_bundle_added = True
                    if debug:
                        st.write(f"  Table {tidx}, Row {ridx}: Added 'WM Bundle'")
                
                if row_ids and debug:
                    st.write(f"  Table {tidx}, Row {ridx}: {row_ids}")
                
                collected_ids.extend(row_ids)
    
    except Exception as e:
        st.error(f"Error extracting IDs: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        return []
    
    return collected_ids

# =====================================================
#  PDF EXTRACTION
# =====================================================

def is_sgm_specialty_row(row):
    row_text = " ".join(str(cell) if cell else "" for cell in row).lower()
    indicators = ["specialty", "sgm", "client administered", "custom supplemental specialty"]
    return any(ind in row_text for ind in indicators)

def extract_ids_from_pdf(path):
    ids = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if not tables:
                continue
            for tbl in tables:
                for row in tbl:
                    if not row:
                        continue
                    if is_sgm_specialty_row(row):
                        continue
                    cell = row[-1]
                    if not cell:
                        continue
                    cell_upper = str(cell).strip().upper()
                    if "WM" in cell_upper and "BUNDLE" in cell_upper:
                        ids.append("WM Bundle")
                        continue
                    if "DERM" in cell_upper and "BUNDLE" in cell_upper:
                        ids.append("Derm Bundle")
                        continue
                    found = ID_PATTERN_SPACED.findall(str(cell).strip())
                    for f in found:
                        ids.append(normalize_id(f))
    return ids

# =====================================================
#  FILTERING
# =====================================================

def apply_final_filters(id_list):
    """Apply final filters: dedupe and remove -H suffix IDs"""
    if not id_list:
        return []
    clean = []
    seen = set()
    for cid in id_list:
        # Skip IDs ending in just -H (single letter H)
        if re.match(r"^C?\d{1,5}-H$", cid):
            continue
        if cid not in seen:
            seen.add(cid)
            clean.append(cid)
    return clean

# =====================================================
#  MAIN EXTRACTION
# =====================================================

def extract_criteria_ids(path, debug=False):
    ext = path.lower().split(".")[-1]
    
    if debug:
        st.write(f"Extension: {ext}, Path: {path}")

    if ext in ("docx", "docm"):
        table_ids = extract_ids_from_doc_tables(path, debug)
        return apply_final_filters(table_ids) if table_ids else []

    if ext == "pdf":
        pdf_ids = extract_ids_from_pdf(path)
        return apply_final_filters(pdf_ids)

    return []

# =====================================================
#  STREAMLIT UI
# =====================================================

st.set_page_config(page_title="Reference Check", page_icon="🔎")

if "criteria_ids" not in st.session_state:
    st.session_state["criteria_ids"] = None
if "selected_sfc" not in st.session_state:
    st.session_state["selected_sfc"] = None
if "selected_file" not in st.session_state:
    st.session_state["selected_file"] = None
if "error" not in st.session_state:
    st.session_state["error"] = None

def clear_error():
    st.session_state["error"] = None

previous_page = st.session_state.get("previous_page", None)
page = st.sidebar.radio("Go to:", ["Reference Check", "CPM Library"])
if previous_page != page:
    clear_error()
st.session_state["previous_page"] = page

st.markdown("""
<div style="background-color:#B00020; padding:14px 18px; border-radius:10px;">
    <span style="color:white; font-size:26px; font-weight:700;">Reference Check</span><br>
    <span style="color:#FFE5EA; font-size:13px;">CPM Criteria ID Extractor</span>
</div>
""", unsafe_allow_html=True)


# Cloud storage warning
st.info("⚠️ **Note:** Files saved to the CPM Library are temporary and will be deleted when the app restarts. Always download your extracted IDs.")

if st.session_state["error"]:
    st.error(st.session_state["error"])
    if st.button("Dismiss Error"):
        clear_error()
        st.rerun()

# =====================================================
#  PAGE 1: REFERENCE CHECK
# =====================================================

if page == "Reference Check":
    clear_error()
    st.write("Upload a **CPM** and extract **Criteria IDs**.")
    
    # Debug toggle
    debug_mode = st.checkbox("Enable Debug Mode")

    tab_upload, tab_existing = st.tabs(["Upload new CPM", "Use existing CPM"])

    with tab_upload:
        file = st.file_uploader("Upload CPM", type=["pdf", "doc", "docx", "docm"])
        sfc = st.text_input("Enter SFC (1-10 digits)", max_chars=10)

        col1, col2 = st.columns(2)
        save_only = col1.button("Save CPM")
        save_extract = col2.button("Save CPM & Extract IDs")

        def valid():
            if not file:
                st.session_state["error"] = "Upload a CPM."
                return False
            if not (sfc.isdigit() and 1 <= len(sfc) <= 10):
                st.session_state["error"] = "SFC must be 1-10 digits."
                return False
            return True

        if save_only or save_extract:
            clear_error()
            if valid():
                try:
                    ensure_cpm_dir()
                    ext = file.name.split(".")[-1]
                    fn = make_new_cpm_filename(sfc, file.name, ext)
                    path = os.path.join(CPM_DIR, fn)
                    with open(path, "wb") as f:
                        f.write(file.getbuffer())

                    if save_only:
                        st.success("Saved.")
                    else:
                        ids = extract_criteria_ids(path, debug=debug_mode)
                        st.session_state["criteria_ids"] = ids
                        st.session_state["selected_sfc"] = sfc
                        st.session_state["selected_file"] = fn
                        st.success(f"Extracted {len(ids)} IDs.")
                except Exception as e:
                    st.error(f"Error: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())

    with tab_existing:
        sfcs = list_existing_sfc_numbers()
        if not sfcs:
            st.info("No CPMs saved.")
        else:
            sfc_sel = st.text_input("Search SFC (Exact Match Required)")
            if sfc_sel and sfc_sel in sfcs:
                cpm_files = [f for f in list_cpm_files() if f["sfc"] == sfc_sel]
                file_options = {f["filename"]: f for f in cpm_files}
                chosen = st.selectbox("Select CPM file", list(file_options.keys()))
                chosen_file = file_options[chosen]

                col1, col2 = st.columns(2)
                if col1.button("Extract IDs"):
                    try:
                        clear_error()
                        ids = extract_criteria_ids(chosen_file["path"], debug=debug_mode)
                        st.session_state["criteria_ids"] = ids
                        st.session_state["selected_sfc"] = sfc_sel
                        st.session_state["selected_file"] = chosen_file["filename"]
                        st.success(f"Extracted {len(ids)} IDs.")
                    except Exception as e:
                        st.error(f"Error: {str(e)}")

                if col2.button("Download CPM"):
                    try:
                        with open(chosen_file["path"], "rb") as f:
                            st.download_button(
                                label="⬇️ Download",
                                data=f,
                                file_name=chosen_file["filename"],
                                mime="application/octet-stream"
                            )
                    except Exception as e:
                        st.error(f"Error: {str(e)}")

    # Results
    ids = st.session_state["criteria_ids"]
    if ids:
        st.markdown("---")
        st.subheader("Copy & Paste List")
        st.text_area("Extracted IDs", value="\n".join(ids), height=200, label_visibility="collapsed")

        df = pd.DataFrame({"criteria_id": ids})
        buf = BytesIO()
        df.to_excel(buf, index=False, engine="openpyxl")
        buf.seek(0)
        st.download_button("Download Excel", buf, "criteria_ids.xlsx")

        if st.button("Clear IDs"):
            st.session_state["criteria_ids"] = None
            st.session_state["selected_sfc"] = None
            st.session_state["selected_file"] = None
            st.rerun()

# =====================================================
#  PAGE 2: CPM LIBRARY
# =====================================================

elif page == "CPM Library":
    clear_error()
    st.subheader("CPM Library")
    
    files = list_cpm_files()
    
    if not files:
        st.info("No CPMs stored.")
    else:
        # Display total count
        st.write(f"**Total CPMs: {len(files)}**")
        st.write("")
        
        # Search/filter section
        col1, col2 = st.columns([2, 1])
        with col1:
            sfc_search = st.text_input("🔍 Search by SFC", placeholder="Enter SFC number")
        with col2:
            ext_filter = st.selectbox("Filter by Type", ["All", "PDF", "DOCX", "DOCM"])
        
        # Apply filters
        filtered_files = files
        if sfc_search:
            filtered_files = [f for f in filtered_files if sfc_search in f["sfc"]]
        if ext_filter != "All":
            filtered_files = [f for f in filtered_files if f["extension"] == ext_filter.lower()]
        
        st.write(f"Showing {len(filtered_files)} CPM(s)")
        st.markdown("---")
        
        # Display list of CPMs
        if filtered_files:
            for idx, file_info in enumerate(filtered_files):
                with st.expander(f"📄 **SFC {file_info['sfc']}** - {file_info['label'] or 'Untitled'} (.{file_info['extension']})"):
                    col_a, col_b = st.columns([3, 1])
                    
                    with col_a:
                        st.write(f"**Filename:** {file_info['filename']}")
                        st.write(f"**SFC:** {file_info['sfc']}")
                        st.write(f"**Type:** {file_info['extension'].upper()}")
                        st.write(f"**Modified:** {file_info['modified'].strftime('%Y-%m-%d %I:%M %p')}")
                    
                    with col_b:
                        if st.button("Extract IDs", key=f"extract_{idx}"):
                            try:
                                ids = extract_criteria_ids(file_info["path"], debug=False)
                                st.session_state["criteria_ids"] = ids
                                st.session_state["selected_sfc"] = file_info["sfc"]
                                st.session_state["selected_file"] = file_info["filename"]
                                st.success(f"Extracted {len(ids)} IDs")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error: {str(e)}")
                        
                        if st.button("Download", key=f"download_{idx}"):
                            try:
                                with open(file_info["path"], "rb") as f:
                                    st.download_button(
                                        label="⬇️ Download",
                                        data=f,
                                        file_name=file_info["filename"],
                                        mime="application/octet-stream",
                                        key=f"dl_btn_{idx}"
                                    )
                            except Exception as e:
                                st.error(f"Error: {str(e)}")
                        
                        if st.button("Delete", key=f"delete_{idx}"):
                            try:
                                os.remove(file_info["path"])
                                st.success("Deleted!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error: {str(e)}")
        else:
            st.info("No CPMs match your search criteria.")

st.markdown(f"<div style='text-align:center; color:#888; margin-top:30px;'>Reference Check v{APP_VERSION}</div>", unsafe_allow_html=True)