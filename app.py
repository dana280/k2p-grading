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

# CSS מקצועי מינימליסטי
st.markdown("""
<style>
    /* ניקוי כללי */
    #MainMenu, footer, header {visibility: hidden;}
    .block-container {padding-top: 2rem; padding-bottom: 2rem;}
    
    /* רקע לבן נקי */
    .main {
        background-color: #ffffff;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }
    
    /* הסתרת padding מיותר */
    .stApp {
        background-color: #ffffff;
    }
    
    /* כותרות מינימליסטיות */
    h1, h2, h3 {
        font-weight: 600;
        letter-spacing: -0.5px;
        margin: 0;
        padding: 0;
    }
    
    h1 {
        font-size: 1.8rem;
        color: #1a1a1a;
    }
    
    h2 {
        font-size: 1.3rem;
        color: #666;
        font-weight: 400;
    }
    
    h3 {
        font-size: 1.1rem;
        color: #1a1a1a;
        margin-top: 2rem;
    }
    
    /* אזור העלאה נקי */
    [data-testid="stFileUploader"] {
        border: 2px solid #e5e7eb;
        border-radius: 8px;
        background-color: #fafafa;
        padding: 2.5rem;
        transition: all 0.2s ease;
    }
    
    [data-testid="stFileUploader"]:hover {
        border-color: #0080C8;
        background-color: #f8f9fa;
    }
    
    [data-testid="stFileUploader"] section {
        padding: 1.5rem;
    }
    
    [data-testid="stFileUploader"] label {
        font-size: 1rem !important;
        font-weight: 500 !important;
        color: #374151 !important;
    }
    
    /* כפתורים נקיים */
    .stButton>button {
        background-color: #0080C8;
        color: white;
        border: none;
        border-radius: 6px;
        padding: 0.65rem 1.5rem;
        font-size: 0.95rem;
        font-weight: 500;
        transition: all 0.2s ease;
        box-shadow: none;
    }
    
    .stButton>button:hover {
        background-color: #006ba1;
        box-shadow: 0 1px 3px rgba(0,0,0,0.12);
    }
    
    /* הגדרות */
    .streamlit-expanderHeader {
        background-color: transparent;
        border-radius: 6px;
        font-weight: 500;
        color: #374151;
        font-size: 0.95rem;
    }
    
    .streamlit-expanderHeader:hover {
        background-color: #f9fafb;
    }
    
    /* הודעות */
    .stSuccess, .stInfo, .stWarning, .stError {
        border-radius: 6px;
        border: 1px solid;
        padding: 0.75rem 1rem;
        font-size: 0.9rem;
    }
    
    .stSuccess {
        background-color: #f0fdf4;
        border-color: #bbf7d0;
        color: #166534;
    }
    
    .stInfo {
        background-color: #eff6ff;
        border-color: #bfdbfe;
        color: #1e40af;
    }
    
    .stError {
        background-color: #fef2f2;
        border-color: #fecaca;
        color: #991b1b;
    }
    
    /* מטריקות */
    [data-testid="stMetricValue"] {
        font-size: 1.5rem;
        font-weight: 600;
        color: #1a1a1a;
    }
    
    [data-testid="stMetricLabel"] {
        font-size: 0.85rem;
        color: #6b7280;
        font-weight: 500;
    }
    
    /* קו מפריד */
    hr {
        border: none;
        border-top: 1px solid #e5e7eb;
        margin: 2rem 0;
    }
    
    /* טבלה */
    table {
        font-size: 0.9rem;
        border-radius: 8px;
        overflow: hidden;
        border: 1px solid #e5e7eb;
    }
    
    table th {
        background-color: #f9fafb;
        color: #374151;
        font-weight: 600;
        font-size: 0.85rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    table td {
        border-color: #f3f4f6;
    }
    
    /* תיבת טקסט */
    .stTextInput>div>div>input {
        border-radius: 6px;
        border: 1px solid #e5e7eb;
        font-size: 0.9rem;
    }
    
    .stTextInput>div>div>input:focus {
        border-color: #0080C8;
        box-shadow: 0 0 0 3px rgba(0,128,200,0.1);
    }
    
    /* Progress bar */
    .stProgress > div > div > div > div {
        background-color: #0080C8;
    }
</style>
""", unsafe_allow_html=True)

# Header עם לוגו
col1, col2, col3 = st.columns([1, 6, 1])
with col3:
    try:
        st.image("k2p_logo.png", width=120)
    except:
        pass

st.markdown("<br>", unsafe_allow_html=True)

# כותרת
st.markdown("# מערכת בדיקת מטלות אקדמאיות")
st.markdown("## קורס התנהגות ארגונית")

st.markdown("<br>", unsafe_allow_html=True)

# API Key
if 'api_key' not in st.session_state:
    st.session_state.api_key = ""

# הגדרות
with st.expander("הגדרות", expanded=False):
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

st.markdown("<br>", unsafe_allow_html=True)

# פונקציות
def read_docx(file):
    try:
        doc = docx.Document(file)
        return '\n'.join([p.text for p in doc.paragraphs if p.text.strip()])
    except Exception as e:
        return f"שגיאה: {str(e)}"

def extract_work_number(filename):
    """חילוץ מספר מטלה בלבד מהשם הקובץ"""
    # הסרת סיומת
    name = filename.replace('.docx', '').replace('.doc', '')
    
    # חיפוש WorkCode_123 או WorkCode-123
    match = re.search(r'WorkCode[_-]?(\d+)', name, re.IGNORECASE)
    if match:
        return match.group(1)
    
    # חיפוש מספר של 8-9 ספרות (מספר תעודת זהות)
    match = re.search(r'\b(\d{8,9})\b', name)
    if match:
        return match.group(1)
    
    # חיפוש כל מספר של 4+ ספרות
    match = re.search(r'\b(\d{4,})\b', name)
    if match:
        return match.group(1)
    
    # חיפוש כל מספר
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
    
    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    header_font = Font(bold=True, size=12, name="Arial")
    header_alignment = Alignment(horizontal="center", vertical="center")
    
    for col in range(1, 5):
        cell = ws.cell(1, col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
    
    row_colors = ["E6F2FF", "E8F5E9", "FFF9E6", "F3E5F5", "FFE6F0", "E1F5FE"]
    
    def get_grade_color(grade):
        if grade >= 90: return "C8E6C9"
        if grade >= 85: return "BBDEFB"
        if grade >= 80: return "FFF59D"
        if grade >= 70: return "FFCC80"
        return "FFCDD2"
    
    for idx, result in enumerate(results):
        row_num = idx + 2
        bg_color = row_colors[idx % len(row_colors)]
        
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
                left=Side(style='thin', color="CCCCCC"),
                right=Side(style='thin', color="CCCCCC"),
                top=Side(style='thin', color="CCCCCC"),
                bottom=Side(style='thin', color="CCCCCC")
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

# העלאת קבצים
st.markdown("### העלאת מטלות")

uploaded_files = st.file_uploader(
    "גרור קבצים לכאן או לחץ לבחירה",
    type=['docx'],
    accept_multiple_files=True,
    help="תומך ב-Word (.docx)"
)

if uploaded_files:
    st.success(f"{len(uploaded_files)} קבצים הועלו")
    
    if st.button("התחל בדיקה", type="primary"):
        if not st.session_state.api_key:
            st.error("נא להזין Claude API Key בהגדרות")
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
            st.success("הבדיקה הושלמה")
            st.rerun()

# תוצאות
if 'results' in st.session_state and st.session_state.results:
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("---")
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
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # טבלה מינימליסטית
    table_html = "<table style='width:100%; border-collapse: collapse;'>"
    table_html += "<thead><tr>"
    table_html += "<th style='padding: 12px; border-bottom: 2px solid #e5e7eb; text-align: right; font-weight: 600;'>קובץ</th>"
    table_html += "<th style='padding: 12px; border-bottom: 2px solid #e5e7eb; text-align: center; font-weight: 600;'>מספר</th>"
    table_html += "<th style='padding: 12px; border-bottom: 2px solid #e5e7eb; text-align: center; font-weight: 600;'>ציון</th>"
    table_html += "<th style='padding: 12px; border-bottom: 2px solid #e5e7eb; text-align: right; font-weight: 600;'>הערות</th>"
    table_html += "</tr></thead><tbody>"
    
    for idx, r in enumerate(st.session_state.results):
        bg = "#fafafa" if idx % 2 == 0 else "#ffffff"
        
        if r['grade'] >= 90:
            grade_color = "#dcfce7"
        elif r['grade'] >= 85:
            grade_color = "#dbeafe"
        elif r['grade'] >= 80:
            grade_color = "#fef3c7"
        else:
            grade_color = "#fee2e2"
            
        table_html += f"<tr style='background-color: {bg};'>"
        table_html += f"<td style='padding: 12px; border-bottom: 1px solid #f3f4f6; text-align: right; font-size: 0.9rem;'>{r['filename']}</td>"
        table_html += f"<td style='padding: 12px; border-bottom: 1px solid #f3f4f6; text-align: center; font-weight: 600; font-size: 0.9rem;'>{r['workNumber']}</td>"
        table_html += f"<td style='padding: 12px; border-bottom: 1px solid #f3f4f6; text-align: center; background-color: {grade_color}; font-weight: 700; font-size: 1.1rem;'>{r['grade']}</td>"
        table_html += f"<td style='padding: 12px; border-bottom: 1px solid #f3f4f6; text-align: right; white-space: pre-line; font-size: 0.85rem; color: #4b5563;'>{r['comments']}</td>"
        table_html += "</tr>"
    
    table_html += "</tbody></table>"
    st.markdown(table_html, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    col1, col2 = st.columns([4, 1])
    
    with col1:
        excel_file = create_styled_excel(st.session_state.results)
        st.download_button(
            label="הורד Excel",
            data=excel_file,
            file_name=f"תוצאות_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    with col2:
        if st.button("נקה", use_container_width=True):
            del st.session_state.results
            st.rerun()

st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown("<div style='text-align:center;color:#9ca3af;font-size:0.85rem;'>K2P • גרסה 2.0</div>", unsafe_allow_html=True)
