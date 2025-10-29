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
    page_title="מערכת בדיקת מטלות - K2P",
    page_icon="📚",
    layout="wide"
)

# CSS מודרני ונקי
st.markdown("""
<style>
    /* רקע נקי */
    .main {
        background-color: #ffffff;
        padding-top: 1rem;
    }
    
    /* הסתרת תפריט המבורגר */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* לוגו */
    .logo-container {
        text-align: right;
        padding: 0;
        margin-bottom: -20px;
    }
    
    /* כותרות מודרניות */
    h1 {
        text-align: center;
        background: linear-gradient(90deg, #0080C8 0%, #7FBA00 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 2.2em;
        font-weight: 800;
        margin: 0.5rem 0;
        padding: 0;
    }
    
    h2 {
        text-align: center;
        color: #0080C8;
        font-size: 1.5em;
        font-weight: 600;
        margin: 0.3rem 0 1rem 0;
    }
    
    h3 {
        color: #0080C8;
        font-weight: 700;
        font-size: 1.3em;
        margin-top: 2rem;
    }
    
    /* אזור העלאה מודרני */
    [data-testid="stFileUploader"] {
        background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%);
        border: 3px dashed #0080C8;
        border-radius: 16px;
        padding: 3rem;
        margin: 2rem 0;
    }
    
    [data-testid="stFileUploader"]:hover {
        border-color: #7FBA00;
        background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%);
        box-shadow: 0 8px 24px rgba(0,128,200,0.15);
        transform: translateY(-2px);
        transition: all 0.3s ease;
    }
    
    [data-testid="stFileUploader"] section {
        padding: 2rem;
    }
    
    [data-testid="stFileUploader"] label {
        font-size: 1.4rem !important;
        font-weight: 700 !important;
        color: #0080C8 !important;
    }
    
    /* כפתורים */
    .stButton>button {
        background: linear-gradient(90deg, #0080C8 0%, #7FBA00 100%);
        color: white;
        font-size: 1.1rem;
        font-weight: 700;
        padding: 0.8rem 2rem;
        border-radius: 12px;
        border: none;
        width: 100%;
        box-shadow: 0 4px 12px rgba(0,128,200,0.3);
        transition: all 0.3s ease;
    }
    
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(0,128,200,0.4);
    }
    
    /* אזור הגדרות */
    .streamlit-expanderHeader {
        background-color: #f8fafc;
        border-radius: 10px;
        font-weight: 600;
    }
    
    /* הודעות */
    .stSuccess, .stInfo, .stWarning, .stError {
        border-radius: 10px;
        padding: 1rem;
        margin: 1rem 0;
    }
    
    /* מטריקות */
    [data-testid="stMetricValue"] {
        font-size: 2rem;
        font-weight: 800;
        color: #0080C8;
    }
    
    /* טבלה */
    table {
        border-radius: 10px;
        overflow: hidden;
    }
    
    /* חלוקת עמודות */
    [data-testid="column"] {
        padding: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# לוגו בראש העמוד
col1, col2, col3 = st.columns([1, 6, 1])
with col3:
    try:
        st.image("k2p_logo.png", width=180)
    except:
        st.markdown("### K2P")

# כותרות
st.markdown("# 📚 מערכת בדיקת מטלות אקדמאיות")
st.markdown("## 🎓 קורס התנהגות ארגונית")

# API Key
if 'api_key' not in st.session_state:
    st.session_state.api_key = ""

# הגדרות מתקפלות
with st.expander("⚙️ הגדרות", expanded=False):
    api_key = st.text_input(
        "Claude API Key",
        type="password",
        value=st.session_state.api_key,
        help="הזן את ה-API Key שלך מ-Anthropic",
        key="api_input"
    )
    if api_key:
        st.session_state.api_key = api_key
        st.success("✅ API Key נשמר")
    
    st.divider()
    st.info("**גרסה:** 2.0  \n**מפתח:** K2P - Knowledge to People")

st.divider()

# פונקציות עזר
def read_docx(file):
    try:
        doc = docx.Document(file)
        return '\n'.join([p.text for p in doc.paragraphs if p.text.strip()])
    except Exception as e:
        return f"שגיאה: {str(e)}"

def extract_work_number(filename):
    match = re.search(r'WorkCode[_-]?(\d+)', filename)
    if match:
        return match.group(1)
    match = re.search(r'(\d{8})', filename)
    if match:
        return match.group(1)
    return ""

def grade_assignment(content, filename, api_key):
    try:
        client = anthropic.Anthropic(api_key=api_key)
        
        prompt = f"""אתה בודק מטלות בקורס התנהגות ארגונית. בדוק לפי המחוון:

**מחוון (100 נק'):**

שאלה 1 - תרבות (40):
- א (15): תרבות כללית = המדינה. אם חסר → 15-
- ב (15): תרבות ארגונית. אם חסר פירוט → 5-
- ג (10): יחסי גומלין. אם חסר → 10-

שאלה 2 - מבנה (20): 3 תיאוריות
שאלה 3 - תהליך (20): 2 תיאוריות
שאלה 4 - תוכן (20): 2 תיאוריות

"ניתן להרחיב" → 5-
"יישום דל" → 5-

**חשוב:**
1. אל תחמיר! רוב הציונים 80-90
2. אסור לכתוב: "לא הבין", "כתב", "הסטודנט"
3. רק מה שחסר
4. כל הערה בשורה נפרדת

**דוגמה:**
"שאלה 1: חסרה התייחסות לתרבות הכללית - תרבות המדינה (15-)
שאלה 3: ניתן להרחיב על מניע העובדים (5-)"

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
    
    # עיצוף כותרות
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
    
    # צבעי שורות
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
            
            if col == 3:  # עמודת ציון
                cell.fill = PatternFill(start_color=get_grade_color(result['grade']), 
                                       end_color=get_grade_color(result['grade']), 
                                       fill_type="solid")
                cell.font = Font(bold=True, size=16, name="Arial")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
                cell.font = Font(size=11, name="Arial")
                
                if col == 2:  # מספר מטלה
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
    
    # רוחב עמודות
    ws.column_dimensions['A'].width = 45
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 100
    
    # גובה שורות - כל הערה בשורה נפרדת!
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
st.markdown("### 📤 העלאת מטלות")

uploaded_files = st.file_uploader(
    "גרור קבצים לכאן או לחץ לבחירה",
    type=['docx'],
    accept_multiple_files=True,
    help="תומך ב-Word (.docx) | עד 50 קבצים"
)

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} קבצים הועלו בהצלחה!")
    
    if st.button("🚀 התחל בדיקה"):
        if not st.session_state.api_key:
            st.error("❌ נא להזין Claude API Key בהגדרות")
        else:
            results = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for idx, file in enumerate(uploaded_files):
                status_text.text(f"בודק מטלה {idx + 1} מתוך {len(uploaded_files)}...")
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
            st.success("✅ הבדיקה הושלמה בהצלחה!")
            st.rerun()

# הצגת תוצאות
if 'results' in st.session_state and st.session_state.results:
    st.divider()
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
        st.metric("מטלות", f"{len(grades)}")
    
    st.markdown("#### 📋 פירוט מטלות")
    
    # טבלה מעוצבת
    table_html = "<table style='width:100%; border-collapse: collapse; margin: 1rem 0;'>"
    table_html += "<thead><tr style='background: linear-gradient(90deg, #0080C8 0%, #7FBA00 100%);'>"
    table_html += "<th style='padding: 12px; border: 1px solid #ddd; color: white; text-align: right;'>שם קובץ</th>"
    table_html += "<th style='padding: 12px; border: 1px solid #ddd; color: white; text-align: center;'>מספר</th>"
    table_html += "<th style='padding: 12px; border: 1px solid #ddd; color: white; text-align: center;'>ציון</th>"
    table_html += "<th style='padding: 12px; border: 1px solid #ddd; color: white; text-align: right;'>הערות</th>"
    table_html += "</tr></thead><tbody>"
    
    for r in st.session_state.results:
        if r['grade'] >= 90:
            grade_color = "#C8E6C9"
        elif r['grade'] >= 80:
            grade_color = "#BBDEFB"
        elif r['grade'] >= 70:
            grade_color = "#FFF59D"
        else:
            grade_color = "#FFCDD2"
            
        table_html += "<tr>"
        table_html += f"<td style='padding: 12px; border: 1px solid #ddd; text-align: right;'>{r['filename']}</td>"
        table_html += f"<td style='padding: 12px; border: 1px solid #ddd; text-align: center; font-weight: bold;'>{r['workNumber']}</td>"
        table_html += f"<td style='padding: 12px; border: 1px solid #ddd; text-align: center; background-color: {grade_color}; font-weight: bold; font-size: 18px;'>{r['grade']}</td>"
        table_html += f"<td style='padding: 12px; border: 1px solid #ddd; text-align: right; white-space: pre-line;'>{r['comments']}</td>"
        table_html += "</tr>"
    
    table_html += "</tbody></table>"
    st.markdown(table_html, unsafe_allow_html=True)
    
    # כפתורי הורדה וניקוי
    col1, col2 = st.columns([3, 1])
    
    with col1:
        excel_file = create_styled_excel(st.session_state.results)
        st.download_button(
            label="📥 הורד קובץ Excel מעוצב",
            data=excel_file,
            file_name=f"דוח_מטלות_K2P_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with col2:
        if st.button("🗑️ נקה תוצאות", use_container_width=True):
            del st.session_state.results
            st.rerun()
    
    st.info("🎨 **קובץ Excel כולל:** ✅ שורות צבעוניות | ✅ כל הערה בשורה נפרדת | ✅ עיצוב מקצועי")

st.divider()
st.markdown("<div style='text-align:center;color:#888;padding:2rem 0;'>K2P - Powered by Claude AI | גרסה 2.0</div>", unsafe_allow_html=True)
