import streamlit as st
import anthropic
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import docx
import io
import re
from datetime import datetime

st.set_page_config(
    page_title="K2P - בדיקת מטלות",
    page_icon="📚",
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    #MainMenu, footer, header {visibility: hidden;}
    .block-container {
        padding-top: 3rem;
        padding-bottom: 3rem;
        max-width: 700px;
    }
    
    .main {
        background-color: #ffffff;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Helvetica', sans-serif;
    }
    
    /* לוגו */
    .logo-top {
        text-align: right;
        margin-bottom: 3rem;
    }
    
    /* כותרת */
    h1, h2, h3 {
        text-align: center;
        font-weight: 500;
        color: #1a1a1a;
    }
    
    h1 {
        font-size: 1.8rem;
        margin-bottom: 0.3rem;
    }
    
    h2 {
        font-size: 1rem;
        color: #666;
        font-weight: 400;
        margin-bottom: 3rem;
    }
    
    /* ריבוע העלאה */
    [data-testid="stFileUploader"] {
        border: 2px solid #e0e0e0;
        border-radius: 8px;
        background: #fafafa;
        padding: 4rem 2rem;
        transition: border-color 0.2s;
        max-width: 450px;
        margin: 0 auto;
    }
    
    [data-testid="stFileUploader"]:hover {
        border-color: #0080C8;
    }
    
    [data-testid="stFileUploader"] section {
        padding: 0;
    }
    
    [data-testid="stFileUploader"] label {
        font-size: 0.95rem !important;
        font-weight: 500 !important;
        color: #666 !important;
    }
    
    /* כפתור */
    .stButton>button {
        background-color: #0080C8;
        color: white;
        border: none;
        border-radius: 6px;
        padding: 0.6rem 2rem;
        font-size: 0.95rem;
        font-weight: 500;
        width: 200px;
        margin: 2rem auto;
        display: block;
        transition: background-color 0.2s;
    }
    
    .stButton>button:hover {
        background-color: #006ba1;
    }
    
    /* הגדרות */
    .streamlit-expanderHeader {
        background: transparent;
        border: 1px solid #e0e0e0;
        border-radius: 6px;
        font-weight: 400;
        font-size: 0.9rem;
        color: #666;
    }
    
    /* הודעות */
    .stSuccess, .stError, .stInfo {
        border-radius: 6px;
        padding: 0.8rem;
        font-size: 0.9rem;
        border: none;
    }
    
    .stSuccess {
        background-color: #f0f9ff;
        color: #0369a1;
    }
    
    .stError {
        background-color: #fef2f2;
        color: #991b1b;
    }
    
    /* מטריקות */
    [data-testid="stMetricValue"] {
        font-size: 1.8rem;
        font-weight: 600;
        color: #1a1a1a;
    }
    
    [data-testid="stMetricLabel"] {
        font-size: 0.8rem;
        color: #888;
        font-weight: 400;
    }
    
    /* תוצאות */
    .results-box {
        background: #fafafa;
        border-radius: 8px;
        padding: 2rem;
        margin-top: 3rem;
    }
    
    /* טבלה */
    table {
        font-size: 0.85rem;
        margin-top: 1.5rem;
    }
    
    table th {
        background-color: #fafafa;
        color: #666;
        font-weight: 500;
        padding: 10px;
        border-bottom: 1px solid #e0e0e0;
        text-transform: none;
    }
    
    table td {
        padding: 10px;
        border-bottom: 1px solid #f5f5f5;
    }
    
    /* טקסט */
    .stTextInput>div>div>input {
        border: 1px solid #e0e0e0;
        border-radius: 6px;
        font-size: 0.9rem;
    }
    
    .stTextInput>div>div>input:focus {
        border-color: #0080C8;
        box-shadow: none;
    }
    
    /* Progress */
    .stProgress > div > div > div > div {
        background-color: #0080C8;
    }
</style>
""", unsafe_allow_html=True)

# לוגו
st.markdown('<div class="logo-top">', unsafe_allow_html=True)
try:
    st.image("https://5el36i5klq.ufs.sh/f/Z3t1XHIXUkD6xQHZrGWFhpxfDNksJS2BnKoAX3W6gZbLziVm", width=350)
except:
    st.markdown("<div style='text-align:right;color:#0080C8;font-weight:600;'>K2P</div>", unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# כותרת
st.markdown("<h1 style='background: linear-gradient(135deg, #0080C8 0%, #00BCD4 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;'>מערכת בדיקת מטלות אקדמאיות</h1>", unsafe_allow_html=True)
st.markdown("## קורס התנהגות ארגונית")

# API
if 'api_key' not in st.session_state:
    st.session_state.api_key = ""

# *** שינוי 1: הוספת counter למחיקת קבצים ***
if 'uploader_key' not in st.session_state:
    st.session_state.uploader_key = 0

with st.expander("הגדרות", expanded=False):
    api_key = st.text_input(
        "Claude API Key",
        type="password",
        value=st.session_state.api_key,
        placeholder="הזן מפתח",
        key="api_input"
    )
    if api_key:
        st.session_state.api_key = api_key
        st.success("נשמר")

def read_docx(file):
    try:
        doc = docx.Document(file)
        return '\n'.join([p.text for p in doc.paragraphs if p.text.strip()])
    except Exception as e:
        return f"שגיאה: {str(e)}"

def extract_work_number(filename):
    name = filename.replace('.docx', '').replace('.doc', '')
    match = re.search(r'WorkCode[_-]?(\d+)', name, re.IGNORECASE)
    if match:
        return match.group(1)
    match = re.search(r'\b(\d{8,9})\b', name)
    if match:
        return match.group(1)
    match = re.search(r'\b(\d{4,})\b', name)
    if match:
        return match.group(1)
    match = re.search(r'(\d+)', name)
    if match:
        return match.group(1)
    return ""

def grade_assignment(content, filename, api_key):
    try:
        client = anthropic.Anthropic(api_key=api_key)
        
        prompt = f"""אתה בודק מטלות בקורס התנהגות ארגונית. בדוק לפי המחוון:

**מחוון (100 נק'):**

שאלה 1 - תרבות (40):
- א (15): תרבות כללית = המדינה. אם חסר לגמרי → 15-
- ב (15): תרבות ארגונית. אם חסר פירוט → 5-
- ג (10): יחסי גומלין. אם חסר → 10-

שאלה 2 - מבנה (20): 3 תיאוריות
שאלה 3 - תהליך (20): 2 תיאוריות
שאלה 4 - תוכן (20): 2 תיאוריות

**הפחתת נקודות:**
- "ניתן להרחיב" או "חסר פירוט קל" → 2-3 נקודות
- חסר דבר משמעותי → 5-15 נקודות

**חשוב מאוד - כתיבת הערות:**
1. כתוב **רק** מה שחסר או חלש
2. אל תכתוב "עבודה טובה", "מצוין", "כל הכבוד" - שום דבר חיובי!
3. אל תכתוב "הסטודנט", "כתב", "לא הבין"
4. כל הערה בשורה נפרדת
5. **חובה**: כתוב את ההפחתה בסוגריים: (-X)
6. אם אין מה לכתוב - השאר ריק לגמרי
7. תהיה נדיב - רוב הציונים 85-95

**פורמט:**
"שאלה 1: חסרה תרבות כללית - תרבות המדינה (-15)
שאלה 3: ניתן להרחיב על מוטיבציה (-2)"

JSON:
{{
  "workNumber": "מספר",
  "grade": 0-100,
  "comments": "הערות או ריק"
}}

קובץ: {filename}
תוכן: {content[:12000]}"""

        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=4000,
            messages=[{"role": "user", "content": prompt}]
        )
        
        response_text = message.content[0].text
        json_match = re.search(r'\{[\s\S]*\}', response_text)
        
        if json_match:
            import json
            result = json.loads(json_match.group(0))
            return result
        
        return {
            "workNumber": extract_work_number(filename),
            "grade": 0,
            "comments": "לא הצלחתי לפענח תשובה"
        }
        
    except Exception as e:
        return {
            "workNumber": extract_work_number(filename),
            "grade": 0,
            "comments": f"שגיאה: {str(e)}"
        }

def create_styled_excel(results):
    wb = Workbook()
    ws = wb.active
    ws.title = "תוצאות"
    
    headers = ['שם קובץ', 'מספר', 'ציון', 'הערות']
    ws.append(headers)
    
    header_fill = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
    header_font = Font(bold=True, size=11, name="Arial")
    header_alignment = Alignment(horizontal="center", vertical="center")
    
    for col in range(1, 5):
        cell = ws.cell(1, col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
    
    def get_grade_color(grade):
        if grade >= 90: return "E8F5E9"
        if grade >= 85: return "E3F2FD"
        if grade >= 80: return "FFF9C4"
        return "FFEBEE"
    
    for idx, result in enumerate(results):
        row_num = idx + 2
        
        ws.append([
            result['filename'],
            result['workNumber'],
            result['grade'],
            result['comments']
        ])
        
        for col in range(1, 5):
            cell = ws.cell(row_num, col)
            
            if col == 3:
                cell.fill = PatternFill(start_color=get_grade_color(result['grade']), 
                                       end_color=get_grade_color(result['grade']), 
                                       fill_type="solid")
                cell.font = Font(bold=True, size=14, name="Arial")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.font = Font(size=10, name="Arial")
                if col == 2:
                    cell.font = Font(bold=True, size=11, name="Arial")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
    
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 100
    
    ws.row_dimensions[1].height = 25
    for idx, result in enumerate(results):
        row_num = idx + 2
        lines = len(result['comments'].split('\n')) if result['comments'] else 1
        ws.row_dimensions[row_num].height = max(50, lines * 18 + 10)
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# העלאה - *** שינוי 2: הוספת key שמשתנה ***
uploaded_files = st.file_uploader(
    "גרור קבצים או לחץ לבחירה",
    type=['docx'],
    accept_multiple_files=True,
    help="קבצי Word בלבד",
    key=f"uploader_{st.session_state.uploader_key}"
)

if uploaded_files:
    st.success(f"{len(uploaded_files)} קבצים הועלו")
    
    if st.button("התחל בדיקה"):
        if not st.session_state.api_key:
            st.error("הזן API Key בהגדרות")
        else:
            results = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for idx, file in enumerate(uploaded_files):
                status_text.text(f"בודק {idx + 1}/{len(uploaded_files)}")
                progress_bar.progress((idx + 1) / len(uploaded_files))
                
                content = read_docx(file)
                result = grade_assignment(content, file.name, st.session_state.api_key)
                results.append({
                    'filename': file.name,
                    'workNumber': result.get('workNumber', ''),
                    'grade': result.get('grade', 0),
                    'comments': result.get('comments', '')
                })
            
            st.session_state.results = results
            progress_bar.empty()
            status_text.empty()
            st.success("הושלם")
            st.rerun()

# תוצאות
if 'results' in st.session_state and st.session_state.results:
    st.markdown('<div class="results-box">', unsafe_allow_html=True)
    st.markdown("### תוצאות")
    
    grades = [r['grade'] for r in st.session_state.results]
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("ממוצע", f"{sum(grades)/len(grades):.1f}")
    with col2:
        st.metric("מקסימום", f"{max(grades)}")
    with col3:
        st.metric("מינימום", f"{min(grades)}")
    with col4:
        st.metric("סה״כ", f"{len(grades)}")
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # טבלה
    table_html = "<table style='width:100%;margin-top:1.5rem;'>"
    table_html += "<thead><tr>"
    table_html += "<th style='text-align:right;'>קובץ</th>"
    table_html += "<th style='text-align:center;'>מספר</th>"
    table_html += "<th style='text-align:center;'>ציון</th>"
    table_html += "<th style='text-align:right;'>הערות</th>"
    table_html += "</tr></thead><tbody>"
    
    for r in st.session_state.results:
        if r['grade'] >= 90:
            color = "#e8f5e9"
        elif r['grade'] >= 85:
            color = "#e3f2fd"
        elif r['grade'] >= 80:
            color = "#fff9c4"
        else:
            color = "#ffebee"
            
        table_html += "<tr>"
        table_html += f"<td style='text-align:right;'>{r['filename']}</td>"
        table_html += f"<td style='text-align:center;font-weight:600;'>{r['workNumber']}</td>"
        table_html += f"<td style='text-align:center;background:{color};font-weight:600;font-size:1.05rem;'>{r['grade']}</td>"
        table_html += f"<td style='text-align:right;white-space:pre-line;font-size:0.85rem;color:#555;'>{r['comments']}</td>"
        table_html += "</tr>"
    
    table_html += "</tbody></table>"
    st.markdown(table_html, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        excel_file = create_styled_excel(st.session_state.results)
        st.download_button(
            label="הורד Excel",
            data=excel_file,
            file_name=f"תוצאות_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col2:
        if st.button("נקה"):
            del st.session_state.results
            st.session_state.uploader_key += 1  # *** שינוי 3: העלאת counter ***
            st.rerun()

st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown("<div style='text-align:center;color:#ccc;font-size:0.8rem;'>K2P • גרסה 2.0</div>", unsafe_allow_html=True)
