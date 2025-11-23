###############################################
#  CPM REFERENCE CHECK APP - v4.0 (Cloud)
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
    Returns True only if this is a dedicated Specialty-only section.
    Non-Specialty tables should NOT be skipped.
    
    For PDFs: Only skip if it's in the "Specialty Programs" section
    AND doesn't contain standard PA criteria.
    """
    text_lower = text.lower()
    
    # Remove all "non-specialty" variants first
    cleaned = re.sub(r'non[\s-]*specialty', '', text_lower)
    
    # If "specialty" doesn't appear at all, it's not specialty-only
    if 'specialty' not in cleaned:
        return False
    
    # Check if this is a standard PA/QL section (has criteria)
    # These should NOT be skipped even if they mention "specialty"
    has_criteria_indicators = any(indicator in text_lower for indicator in [
        'prior authorization required',
        'target drugs:',
        'criteria id',
        'therapeutic class',
        'quantity limit',
        'duration of approval'
    ])
    
    # If it has criteria indicators, it's NOT specialty-only
    if has_criteria_indicators:
        return False
    
    # Check if this is in the dedicated "Specialty Programs" section
    is_specialty_section = any(header in text_lower for header in [
        'specialty programs',
        'specialty program management',
        'enhanced sgm'
    ])
    
    # Only skip if it's in the specialty section AND has no criteria
    return is_specialty_section

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
    
    NOT detected as reference (these are legitimate):
    - "Therapy Name Ref #ID" (criteria listing format)
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
    # Only skip if there's "Criteria Drugs:" before AND a colon AFTER the ID
    # "Ref #ID" by itself (or with just "*") is a legitimate criteria reference
    if 'criteria drugs:' in before[-50:] and ':' in after[:10]:
        return True
    
    # Pattern 7: Instructional text like "see criteria" or "see table above"
    if re.search(r'\b(see|refer to|reference)\s+(criteria|table|section)', before[-100:]):
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
    4. If checkbox is ☐, exclude the ID
    5. If no controlling checkbox, use default_include
    6. Skip IDs ending in single -H (quantity limits without post-PA)
    """
    # Find all checkboxes with their positions
    checkboxes = []
    for match in re.finditer(r'☒|☐', text):
        checkboxes.append({
            'char': match.group(),
            'pos': match.start(),
            'checked': match.group() == '☒'
        })
    
    # Find all IDs with their positions
    id_matches = []
    for match in ID_PATTERN_SPACED.finditer(text):
        id_str = normalize_id(match.groups())
        id_matches.append({
            'id': id_str,
            'start': match.start(),
            'end': match.end()
        })
    
    extracted_ids = []
    
    for id_match in id_matches:
        # Skip if this ID appears in reference text
        if is_reference_text(text, id_match['start'], id_match['end']):
            continue
        
        # Skip Weight Management IDs if requested
        if skip_wm_ids and id_match['id'] in WM_BUNDLE_IDS:
            continue
        
        # Skip IDs ending in single -H (quantity limits without post-PA)
        # Keep IDs ending in -HJ or other combinations
        if id_match['id'].endswith('-H') and not any(id_match['id'].endswith(f'-H{letter}') for letter in 'ABCDEFGIJKLMNOPQRSTUVWXYZ'):
            continue
        
        # Find the closest checkbox BEFORE this ID
        controlling_checkbox = None
        min_distance = float('inf')
        
        for cb in checkboxes:
            if cb['pos'] < id_match['start']:
                distance = id_match['start'] - cb['pos']
                if distance < min_distance:
                    min_distance = distance
                    controlling_checkbox = cb
        
        # Decision logic
        if controlling_checkbox and min_distance < 200:
            if controlling_checkbox['checked']:
                extracted_ids.append(id_match['id'])
        elif default_include:
            extracted_ids.append(id_match['id'])
    
    return extracted_ids

def has_xml_checkboxes(table_elem):
    """Check if table contains XML-based checkboxes (not Unicode symbols)"""
    checkboxes = table_elem.findall('.//w:checkBox', NAMESPACES)
    return len(checkboxes) > 0

def get_xml_checkbox_state(table_elem):
    """
    For tables with XML checkboxes, determine if they indicate 'add' action.
    Returns True if any checkbox in an 'Add' context is checked.
    """
    # Get all checkbox elements
    checkboxes = table_elem.findall('.//w:checkBox', NAMESPACES)
    
    if not checkboxes:
        return False
    
    # Get table text to check context
    table_text = get_text_from_element(table_elem).lower()
    
    # Check each checkbox
    for cb_elem in checkboxes:
        # Find the checked state
        checked_elem = cb_elem.find('.//w:checked', NAMESPACES)
        if checked_elem is not None:
            val = checked_elem.get(f'{{{NAMESPACES["w"]}}}val')
            is_checked = val != '0' if val else True
            
            if is_checked:
                # This checkbox is checked - check if it's in an "Add" context
                # by looking at the parent paragraph text
                parent = cb_elem
                for _ in range(5):  # Go up a few levels
                    parent = parent.getparent() if parent is not None else None
                    if parent is None:
                        break
                    parent_text = get_text_from_element(parent).lower()
                    if 'add' in parent_text and not 'delete' in parent_text[:parent_text.find('add') + 10]:
                        return True
    
    return False

def process_table_xml(table_elem, debug=False):
    """
    Process a table element from Word XML.
    Returns list of extracted IDs based on checkbox states.
    """
    table_text = get_text_from_element(table_elem)
    
    if debug:
        st.write("---")
        st.write("**TABLE FOUND**")
        st.text(table_text[:500])
    
    # Skip specialty-only tables
    if is_specialty_only(table_text):
        if debug:
            st.write("❌ **SKIPPED:** Specialty Only table")
        return []
    
    # Check if this is a Weight Management Bundle table
    is_wm_bundle_table = is_wm_bundle_row(table_text)
    
    # Determine default behavior
    default_include = is_auto_included_table(table_text)
    action_status = get_action_status_from_text(table_text)
    
    # For XML checkboxes (DOCM/DOCX), check their state directly
    has_xml_cb = has_xml_checkboxes(table_elem)
    if has_xml_cb and action_status is None:
        # If we have XML checkboxes but no Unicode checkbox symbols in text,
        # check the XML checkbox state
        if get_xml_checkbox_state(table_elem):
            action_status = 'add'
            default_include = True
    
    if action_status == 'add':
        default_include = True
    elif action_status == 'delete':
        default_include = False
    
    if debug:
        st.write(f"**Auto-include:** {default_include}")
        st.write(f"**Action Status:** {action_status}")
        st.write(f"**WM Bundle:** {is_wm_bundle_table}")
        if has_xml_cb:
            st.write(f"**Has XML checkboxes:** True")
    
    # Extract IDs
    ids = extract_ids_with_checkbox_awareness(
        table_text, 
        default_include=default_include,
        skip_wm_ids=is_wm_bundle_table
    )
    
    # Add WM Bundle if applicable
    if is_wm_bundle_table and action_status == 'add':
        ids.append("WM Bundle")
    
    if debug:
        st.write(f"**Extracted IDs ({len(ids)}):** {ids}")
    
    return ids

def extract_criteria_ids_from_doc(path, debug=False):
    """Extract criteria IDs from DOCX/DOCM file"""
    root = load_xml_from_doc(path)
    tables = root.findall('.//w:tbl', NAMESPACES)
    
    if debug:
        st.write(f"**Total tables found:** {len(tables)}")
    
    all_ids = []
    for table in tables:
        ids = process_table_xml(table, debug=debug)
        all_ids.extend(ids)
    
    return list(dict.fromkeys(all_ids))  # Remove duplicates while preserving order

def extract_criteria_ids_from_pdf(path, debug=False):
    """Extract criteria IDs from PDF file"""
    all_ids = []
    
    with pdfplumber.open(path) as pdf:
        if debug:
            st.write(f"**Total pages:** {len(pdf.pages)}")
        
        for page_num, page in enumerate(pdf.pages, 1):
            text = page.extract_text()
            if not text:
                continue
            
            if debug:
                st.write(f"**Page {page_num}:**")
                st.text(text[:300])
            
            # Skip specialty-only content
            if is_specialty_only(text):
                if debug:
                    st.write("❌ **SKIPPED:** Specialty Only")
                continue
            
            # Check for Weight Management Bundle
            is_wm_bundle = is_wm_bundle_row(text)
            
            # Determine default behavior
            default_include = is_auto_included_table(text)
            action_status = get_action_status_from_text(text)
            
            if action_status == 'add':
                default_include = True
            elif action_status == 'delete':
                default_include = False
            
            if debug:
                st.write(f"**Auto-include:** {default_include}")
                st.write(f"**Action Status:** {action_status}")
                st.write(f"**WM Bundle:** {is_wm_bundle}")
            
            # Extract IDs
            ids = extract_ids_with_checkbox_awareness(
                text,
                default_include=default_include,
                skip_wm_ids=is_wm_bundle
            )
            
            # Add WM Bundle if applicable
            if is_wm_bundle and action_status == 'add':
                ids.append("WM Bundle")
            
            if debug:
                st.write(f"**Extracted IDs ({len(ids)}):** {ids}")
            
            all_ids.extend(ids)
    
    return list(dict.fromkeys(all_ids))  # Remove duplicates while preserving order

def extract_criteria_ids(path, debug=False):
    """Main extraction function that routes to appropriate handler"""
    ext = path.split(".")[-1].lower()
    
    if ext in ["docx", "docm"]:
        return extract_criteria_ids_from_doc(path, debug=debug)
    elif ext == "pdf":
        return extract_criteria_ids_from_pdf(path, debug=debug)
    else:
        raise ValueError(f"Unsupported file type: {ext}")

# =====================================================
#  STREAMLIT APP
# =====================================================

st.set_page_config(page_title="CPM Criteria ID Extractor", page_icon="📋", layout="wide")

# Session state
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
st.info("⚠️ **Note:** This is running on Streamlit Community Cloud. Files saved to the CPM Library are temporary and will be deleted when the app restarts. For permanent storage, download your extracted IDs.")

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

                # "Open CPM" button removed for cloud deployment (os.startfile is Windows-only)
                if col2.button("Download CPM"):
                    try:
                        with open(chosen_file["path"], "rb") as f:
                            st.download_button(
                                label="⬇️ Click to Download",
                                data=f,
                                file_name=chosen_file["filename"],
                                mime="application/octet-stream",
                                key=f"download_{chosen_file['filename']}"
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
                        
                        # "Open" button removed for cloud deployment
                        if st.button("Download", key=f"download_{idx}"):
                            try:
                                with open(file_info["path"], "rb") as f:
                                    st.download_button(
                                        label="⬇️ Click to Download",
                                        data=f,
                                        file_name=file_info["filename"],
                                        mime="application/octet-stream",
                                        key=f"dl_{idx}"
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