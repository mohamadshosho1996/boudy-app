import os
import datetime
import pandas as pd
import streamlit as st

# ==================== إعدادات الصفحة والتصميم ====================
st.set_page_config(
    page_title="برنامج بودى للمشورة الأسرية",
    page_icon="🌸",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main .block-container { padding-top: 1.0rem; padding-bottom: 1rem; max-width: 100%; }
    .main { background-color: #FFF5F8; }
    .stButton>button { background-color: #EC4899; color: white; border-radius: 8px; font-weight: bold; border: none; width: 100%; }
    .stButton>button:hover { background-color: #BE185D; }
    h1, h2, h3 { color: #701A75; }
    footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# ==================== الثوابت وإعدادات البيانات ====================
EXCEL_FILE = "template.xlsx"

DEFAULT_USERS = {
    "admin": {"pass": "admin123", "role": "admin", "name": "د. شيماء 🌸"},
    "user1": {"pass": "1234", "role": "user", "name": "د. علا 🎀"},
    "user2": {"pass": "1234", "role": "user", "name": "د. عبير 🎀"},
    "user3": {"pass": "1234", "role": "user", "name": "د. ايه 🎀"}
}

DROPDOWN_OPTIONS = {
    "مستوى التعليم": ["امى", "يجيد القراءة", "مؤهل متوسط", "فوق متوسط", "مؤهل عالى"],
    "مستوى التعليم للام": ["امى", "يجيد القراءة", "مؤهل متوسط", "فوق متوسط", "مؤهل عالى"],
    "مستوى التعليم للاب": ["امى", "يجيد القراءة", "مؤهل متوسط", "فوق متوسط", "مؤهل عالى"],
    "الوظيفة": ["يعمل", "لا تعمل"],
    "الوظيفة للام": ["يعمل", "لا تعمل"],
    "نوع الولادة": ["طبيعى", "قيصرى"],
    "مكان الولادة": ["المستشفى", "المنزل"],
    "سبب دخول الحضانة": ["انخفاض وزن الطفل.", "احتياج الطفل لأدوية محددة بهذا الوقت.", "صعوبة شديدة في التنفس", "إصابة الطفل بالصفراء.", "أخرى"],
    "رضاعة طبيعية مطلقة": ["3 شهور", "4 شهور", "6 شهور"],
    "موقف إستخدام وسيلة تنظيم أسرة": ["توجد", "لا يوجد", "مرغوب", "غير مرغوب"],
    "النمو والتطور الحركي": ["طبيعي", "متأخر", "يحتاج متابعة", "تم التوعية"],
}

CHILD_COLUMNS = [
    "تاريخ التسجيل", "اسم المستخدم", "تاريخ اول زيارة", "رقم الحالة", "اسم الام", "الرقم القومى للام", 
    "رقم الموبايل للام", "تاريخ ميلاد الام", "مستوى التعليم للام", "عدد الاطفال لدى الام", "المدة بين اخر حملين", 
    "الوظيفة للام", "الرقم القومى للاب", "رقم الموبايل للاب", "اسم الاب", "مستوى التعليم للاب", "اسم الطفل", 
    "تاريخ الميلاد للطفل", "العمر الحالى للطفل (شهور)", "العمر الرحمى للطفل (أسابيع)",
    "مكان المتابعة", "وحدة", "مستشفى", "أخرى",
    "مصدر الاحاله", "مستشفى الولادة", "عيادة خاصة", "عيادة التطعيمات", "نصيحة",
    "نوع الولادة", "مكان الولادة", "وزن الطفل عند الولادة", "طول الطفل عند الولادة", "مقاس راس الطفل عند الولادة", 
    "دخول الحضانة", "سبب دخول الحضانة", "مدة البقاء فى الحضانة", "ملامسة الجلد فى الساعة الذهبية الأولى", 
    "الرضاعة الطبيعية فى الساعة الذهبية الأولى", "موعد الزيارة", "تاريخ الزيارة", "رضاعة طبيعية مطلقة", 
    "رضاعة طبيعية مع سوائل وأعشاب", "رضاعة طبيعية مع صناعي", "رضاعة لبن صناعي", "الوزن (كجم)", "الطول (سم)", 
    "محيط الرأس (سم)", "النمو والتطور الحركي", "التطور الإدراكي والمعرفي", "التطور اللغوي", "رسائل التربية الإيجابية", 
    "الأنشطة التحفيزية", "إعطاء الجرعة اليومية من الحديد", "أهمية إستخدام وسيلة تنظيم أسرة وأهمية المباعدة", 
    "موقف إستخدام وسيلة تنظيم أسرة", "الحمل الجديد", "الخدمات الغير ملباه", "ملاحظات/ توصيات"
]

# ==================== دوال المساعدة ====================
def parse_national_id(nat_id):
    if nat_id and len(nat_id) == 14 and nat_id.isdigit():
        century = 2000 if int(nat_id[0]) == 3 else 1900
        birth_date = datetime.date(century + int(nat_id[1:3]), int(nat_id[3:5]), int(nat_id[5:7]))
        return str(birth_date)
    return ""

def get_existing_data(nat_id, sheet_name, id_column):
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, dtype=str)
            match = df[df[id_column] == nat_id]
            if not match.empty: return match.iloc[-1].to_dict()
        except: pass
    return {}

# ==================== تسجيل الدخول ====================
if "logged_in" not in st.session_state: st.session_state.logged_in = False

if not st.session_state.logged_in:
    username = st.selectbox("اختر الطبيبة", list(DEFAULT_USERS.keys()))
    password = st.text_input("كلمة المرور", type="password")
    if st.button("دخول"):
        if DEFAULT_USERS[username]["pass"] == password:
            st.session_state.update({"logged_in": True, "user": username, "name": DEFAULT_USERS[username]["name"], "role": DEFAULT_USERS[username]["role"]})
            st.rerun()
    st.stop()

# ==================== واجهة الأطفال ====================
st.markdown("<h2>👶 سجل المشورة الأسرية للأطفال</h2>", unsafe_allow_html=True)
nat_id_mom = st.text_input("الرقم القومى للام", max_chars=14)

form_data = {col: "" for col in CHILD_COLUMNS}

# الحقول الجديدة بالقوائم المنسدلة
col_nav1, col_nav2 = st.columns(2)
with col_nav1:
    place_choice = st.selectbox("مكان المتابعة", ["وحدة", "مستشفى", "أخرى"])
with col_nav2:
    source_choice = st.selectbox("مصدر الاحاله", ["مستشفى الولادة", "عيادة خاصة", "عيادة التطعيمات", "نصيحة"])

# باقي الحقول هنا (تم تبسيط العرض)
name_mom = st.text_input("اسم الأم")

if st.button("💾 حفظ البيانات"):
    form_data.update({
        "اسم الام": name_mom,
        "الرقم القومى للام": nat_id_mom,
        "مكان المتابعة": place_choice,
        "وحدة": "تم" if place_choice == "وحدة" else "",
        "مستشفى": "تم" if place_choice == "مستشفى" else "",
        "أخرى": "تم" if place_choice == "أخرى" else "",
        "مصدر الاحاله": source_choice,
        "مستشفى الولادة": "تم" if source_choice == "مستشفى الولادة" else "",
        "عيادة خاصة": "تم" if source_choice == "عيادة خاصة" else "",
        "عيادة التطعيمات": "تم" if source_choice == "عيادة التطعيمات" else "",
        "نصيحة": "تم" if source_choice == "نصيحة" else ""
    })
    
    # كود الحفظ في الإكسيل
    new_df = pd.DataFrame([form_data])
    if os.path.exists(EXCEL_FILE):
        with pd.ExcelWriter(EXCEL_FILE, mode='a', if_sheet_exists='overlay', engine="openpyxl") as writer:
            new_df.to_excel(writer, sheet_name="سجل المشورة للاطفال", index=False, header=False, startrow=writer.sheets["سجل المشورة للاطفال"].max_row)
    else:
        new_df.to_excel(EXCEL_FILE, sheet_name="سجل المشورة للاطفال", index=False)
    st.success("تم الحفظ بنجاح!")
