import streamlit as st
import anthropic
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import docx
import io
import re
from datetime import datetime

# הגדרות עמוד
st.set_page_config(
    page_title="K2P - מערכת בדיקת מטלות",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# CSS נקי ומינימליסטי
st.markdown("""
<style>
    /* ניקוי */
    #MainMenu, footer, header {visibility: hidden;}
    .block-container {padding-top: 1.5rem; max-width: 1200px;}
    
    /* רקע נקי */
    .main {
        background-color: #f8f9fa;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }
    
    /* לוגו בפינה */
    .logo-corner {
        position: fixed;
        top: 1rem;
        right: 2rem;
        z-index: 999;
    }
    
    /* כותרת פשוטה */
    .header-simple {
        text-align: center;
        padding: 1.5rem 2rem;
        background: white;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        margin-bottom: 2rem;
        border-left: 4px solid #0080C8;
    }
    
    .header-simple h1 {
        color: #1a1a1a;
        font-size: 1.8rem;
        font-weight: 600;
        margin: 0;
    }
    
    .header-simple h2 {
        color: #666;
        font-size: 1.1rem;
        font-weight: 400;
        margin: 0.3rem 0 0 0;
    }
    
    /* ריבוע העלאה באמצע */
    .upload-square-container {
        max-width: 500px;
        margin: 3rem auto;
    }
    
    [data-testid="stFileUploader"] {
        border: 2px dashed #0080C8;
        border-radius: 12px;
        background: white;
        padding: 2.5rem;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        transition: all 0.3s ease;
        min-height: 250px;
        display: flex;
        align-items: center;
        justify-content: center;
    }
    
    [data-testid="stFileUploader"]:hover {
        border-color: #7FBA00;
        box-shadow: 0 6px 16px rgba(0,128,200,0.15);
        transform: translateY(-2px);
    }
    
    [data-testid="stFileUploader"] section {
        padding: 1.5rem;
        text-align: center;
        width: 100%;
    }
    
    [data-testid="stFileUploader"] label {
        font-size: 1.1rem !important;
        font-weight: 600 !important;
        color: #0080C8 !important;
    }
    
    /* כפתור */
    .stButton>button {
        background-color: #0080C8;
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.7rem 2rem;
        font-size: 1rem;
        font-weight: 600;
        transition: all 0.2s ease;
        box-shadow: 0 2px 8px rgba(0,128,200,0.2);
    }
    
    .stButton>button:hover {
        background-color: #006ba1;
        box-shadow: 0 4px 12px rgba(0,128,200,0.3);
        transform: translateY(-1px);
    }
    
    /* מרכוז כפתור */
    .button-center {
        text-align: center;
        margin: 2rem auto;
    }
    
    /* הגדרות */
    .streamlit-expanderHeader {
        background: white;
        border-radius: 8px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.05);
        font-weight: 500;
        padding: 0.8rem;
        border-left: 3px solid #0080C8;
    }
    
    /* הודעות */
    .stSuccess {
        background-color: #f0f9ff;
        border-left: 4px solid #0080C8;
        border-radius: 8px;
        padding: 1rem;
        color: #0369a1;
    }
    
    .stInfo {
        background-color: #eff6ff;
        border-left: 4px solid #3b82f6;
        border-radius: 8px;
        padding: 1rem;
        color: #1e40af;
    }
    
    .stError {
        background-color: #fef2f2;
        border-left: 4px solid #ef4444;
        border-radius: 8px;
        padding: 1rem;
        color: #991b1b;
    }
    
    /* מטריקות */
    [data-testid="stMetricValue"] {
        font-size: 2rem;
        font-weight: 700;
        color: #0080C8;
    }
    
    [data-testid="stMetricLabel"] {
        font-size: 0.85rem;
        color: #666;
        font-weight: 600;
    }
    
    /* קלפי תוצאות */
    .results-container {
        background: white;
        border-radius: 12px;
        padding: 2rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        margin-top: 2rem;
    }
    
    /* טבלה */
    table {
        background: white;
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        border: 1px solid #e5e7eb;
    }
    
    table thead {
        background-color: #f8f9fa;
    }
    
    table th {
        color: #1a1a1a !important;
        font-weight: 600;
        padding: 12px !important;
        border-bottom: 2px solid #e5e7eb;
    }
    
    table td {
        padding: 10px !important;
        border-bottom: 1px solid #f3f4f6;
    }
    
    /* תיבת טקסט */
    .stTextInput>div>div>input {
        border-radius: 8px;
        border: 1px solid #e0e0e0;
        padding: 0.6rem;
    }
    
    .stTextInput>div>div>input:focus {
        border-color: #0080C8;
        box-shadow: 0 0 0 2px rgba(0,128,200,0.1);
    }
    
    /* Progress bar */
    .stProgress > div > div > div > div {
        background-color: #0080C8;
    }
    
    /* מרווחים */
    .space-small {margin: 1rem 0;}
    .space-medium {margin: 2rem 0;}
    .space-large {margin: 3rem 0;}
</style>
""", unsafe_allow_html=True)

# לוגו בפינה
st.markdown('<div class="logo-corner">', unsafe_allow_html=True)
try:
    st.image("k2p_logo.png", width=150)
except:
    pass
st.markdown('</div>', unsafe_allow_html=True)

# כותרת פשוטה
st.markdown("""
<div class="header-simple">
    <h1>מערכת בדיקת מטלות אקדמאיות</h1>
    <h2>קורס התנהגות ארגונית</h2>
</div>
""", unsafe_allow_html=True)

# API Key
if 'api_key' not in st.session_state:
    st.session_state.api_key = ""

# הגדרות
with st.expander("⚙️ הגדרות", expanded=False):
    api_key = st.text_input(
        "Claude API Key",
        type="password",
        value=st.session_state.api_key,
        placeholder="הזן API Key",
        key="api_input"
    )
    if api_key:
        st.session_state.api_key = api_key
        st.success("API Key נשמר")

# פונקציות
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
2. אל תכתוב "עבודה טובה", "מצוין", "כל הכבוד", "נעשה יפה" - שום דבר חיובי!
3. אל תכתוב "הסטודנט", "כתב", "לא הבין"
4. כל הערה בשורה נפרדת
5. **חובה**: כתוב את ההפחתה בסוגריים בסוף כל הערה
6. אם אין מה לכתוב - השאר ריק (אל תכתוב כלום!)
7. תהיה נדיב - רוב הציונים 85-95

**פורמט נכון:**
"שאלה 1: חסרה תרבות כללית - תרבות המדינה (-15)
שאלה 3: ניתן להרחיב על מוטיבציה (-2)"

**פורמט לא נכון (אסור!):**
"עבודה טובה מאוד! רק..."
"הסטודנט כתב יפה אבל..."
"מצוין! חסר רק..."

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
    ws.title = "תוצאות בדיקה"
    
    headers = ['שם קובץ', 'מספר', 'ציון', 'הערות']
    ws.append(headers)
    
    header_fill = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
    header_font = Font(bold=True, size=12, name="Arial", color="1A1A1A")
    header_alignment = Alignment(horizontal="center", vertical="center")
    
    for col in range(1, 5):
        cell = ws.cell(1, col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = Border(
            left=Side(style='thin', color="E5E7EB"),
            right=Side(style='thin', color="E5E7EB"),
            top=Side(style='thin', color="E5E7EB"),
            bottom=Side(style='medium', color="E5E7EB")
        )
    
    row_colors = ["FFFFFF", "F9FAFB"]
    
    def get_grade_color(grade):
        if grade >= 90: return "D1FAE5"
        if grade >= 85: return "DBEAFE"
        if grade >= 80: return "FEF3C7"
        if grade >= 70: return "FED7AA"
        return "FEE2E2"
    
    for idx, result in enumerate(results):
        row_num = idx + 2
        bg_color = row_colors[idx % 2]
        
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
                cell.font = Font(bold=True, size=16, name="Arial")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
                cell.font = Font(size=11, name="Arial")
                
                if col == 2:
                    cell.font = Font(bold=True, size=12, name="Arial")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
            
            cell.border = Border(
                left=Side(style='thin', color="E5E7EB"),
                right=Side(style='thin', color="E5E7EB"),
                top=Side(style='thin', color="E5E7EB"),
                bottom=Side(style='thin', color="E5E7EB")
            )
    
    ws.column_dimensions['A'].width = 45
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 100
    
    ws.row_dimensions[1].height = 30
    for idx, result in enumerate(results):
        row_num = idx + 2
        lines = len(result['comments'].split('\n')) if result['comments'] else 1
        ws.row_dimensions[row_num].height = max(60, lines * 20 + 10)
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ריבוע העלאה באמצע
st.markdown('<div class="upload-square-container">', unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "📤 גרור קבצים לכאן או לחץ לבחירה",
    type=['docx'],
    accept_multiple_files=True,
    help="תומך ב-Word (.docx) בלבד"
)

st.markdown('</div>', unsafe_allow_html=True)

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} קבצים הועלו")
    
    st.markdown('<div class="button-center">', unsafe_allow_html=True)
    if st.button("🚀 התחל בדיקה", type="primary"):
        if not st.session_state.api_key:
            st.error("❌ נא להזין Claude API Key בהגדרות")
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
            st.success("✅ הבדיקה הושלמה!")
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# תוצאות
if 'results' in st.session_state and st.session_state.results:
    st.markdown('<div class="results-container">', unsafe_allow_html=True)
    st.markdown("### 📊 תוצאות הבדיקה")
    
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
    
    st.markdown('<div class="space-medium"></div>', unsafe_allow_html=True)
    
    # טבלה
    table_html = "<table style='width:100%;'>"
    table_html += "<thead><tr>"
    table_html += "<th style='text-align: right;'>קובץ</th>"
    table_html += "<th style='text-align: center;'>מספר</th>"
    table_html += "<th style='text-align: center;'>ציון</th>"
    table_html += "<th style='text-align: right;'>הערות</th>"
    table_html += "</tr></thead><tbody>"
    
    for idx, r in enumerate(st.session_state.results):
        
        if r['grade'] >= 90:
            grade_color = "#d1fae5"
        elif r['grade'] >= 85:
            grade_color = "#dbeafe"
        elif r['grade'] >= 80:
            grade_color = "#fef3c7"
        else:
            grade_color = "#fee2e2"
            
        table_html += "<tr>"
        table_html += f"<td style='text-align: right;'>{r['filename']}</td>"
        table_html += f"<td style='text-align: center; font-weight: 600;'>{r['workNumber']}</td>"
        table_html += f"<td style='text-align: center; background-color: {grade_color}; font-weight: 700; font-size: 1.1rem;'>{r['grade']}</td>"
        table_html += f"<td style='text-align: right; white-space: pre-line; font-size: 0.9rem; color: #555;'>{r['comments']}</td>"
        table_html += "</tr>"
    
    table_html += "</tbody></table>"
    st.markdown(table_html, unsafe_allow_html=True)
    
    st.markdown('<div class="space-medium"></div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([2, 1, 2])
    
    with col1:
        excel_file = create_styled_excel(st.session_state.results)
        st.download_button(
            label="📥 הורד Excel",
            data=excel_file,
            file_name=f"תוצאות_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with col3:
        if st.button("🗑️ נקה", use_container_width=True):
            del st.session_state.results
            st.rerun()

st.markdown('<div class="space-large"></div>', unsafe_allow_html=True)
st.markdown("<div style='text-align:center;color:#999;font-size:0.85rem;'>K2P - Knowledge to People • גרסה 2.0</div>", unsafe_allow_html=True)
