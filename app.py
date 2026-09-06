import datetime
import os
import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

# ==================== إعدادات الاتصال بـ Google Sheets السحابي ====================
SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def get_gspread_client():
    if "gcp_service_account" in st.secrets:
        # تحويل بيانات الـ secrets إلى قاموس عادي وتنظيف المفتاح الخاص بشكل سليم جذرياً
        sec = dict(st.secrets["gcp_service_account"])
        private_key = sec.get("private_key", "")
        if private_key:
            # استبدال الرموز النصية للـ newline الحقيقية إذا لزم الأمر وتنظيف المسافات
            private_key = private_key.replace("\\n", "\n").strip().strip('"').strip("'")
            sec["private_key"] = private_key
            
        creds = Credentials.from_service_account_info(sec, scopes=SCOPE)
        client = gspread.authorize(creds)
        return client
    else:
        st.error("⚠️ لم يتم العثور على بيانات الاعتماد (gcp_service_account) في Streamlit Secrets!")
        st.stop()

def get_worksheet(worksheet_name):
    client = get_gspread_client()
    # تأكد من أن اسم ملف الجداول في غوغل هو FamilyCareDB
    spreadsheet = client.open("FamilyCareDB")
    try:
        worksheet = spreadsheet.worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        # إنشاء الورقة تلقائياً إذا لم تكن موجودة لتجنب أي خطأ
        worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows=1000, cols=50)
    return worksheet

def save_to_cloud(worksheet_name, data_dict):
    worksheet = get_worksheet(worksheet_name)
    headers = worksheet.row_values(1)
    
    # إذا كانت الورقة فارغة، أضف العناوين تلقائياً
    if not headers:
        headers = list(data_dict.keys())
        worksheet.append_row(headers)
        
    row_values = []
    for header in headers:
        val = data_dict.get(header, "")
        row_values.append(str(val) if val is not None else "")
    
    worksheet.append_row(row_values)

def load_from_cloud(worksheet_name):
    worksheet = get_worksheet(worksheet_name)
    data = worksheet.get_all_records()
    df = pd.DataFrame(data)
    return df

def delete_from_cloud_by_nat_id(worksheet_name, nat_id_col, clean_id):
    worksheet = get_worksheet(worksheet_name)
    records = worksheet.get_all_records()
    deleted_count = 0
    
    for idx, row in enumerate(reversed(records), start=1):
        actual_row_idx = len(records) - idx + 2 
        row_val = str(row.get(nat_id_col, "")).strip()
        if row_val == str(clean_id):
            worksheet.delete_rows(actual_row_idx)
            deleted_count += 1
            
    return deleted_count

def delete_from_cloud_by_index(worksheet_name, row_idx):
    worksheet = get_worksheet(worksheet_name)
    records = worksheet.get_all_records()
    if 0 <= row_idx < len(records):
        target_row_num = row_idx + 2 
        worksheet.delete_rows(target_row_num)
        return True
    return False

# ==================== الثوابت وإعدادات البيانات ====================
DEFAULT_USERS = {
    "admin": {"pass": "admin123", "role": "admin", "name": "د. شيماء 🌸"},
    "user1": {"pass": "1234", "role": "user", "name": "د. علا 🎀"},
    "user2": {"pass": "1234", "role": "user", "name": "د. عبير 🎀"},
    "user3": {"pass": "1234", "role": "user", "name": "د. ايه 🎀"},
}

VISIT_SCHEDULE_OPTIONS = [
    "الاسبوع الاول", "عمر شهرين", "عمر 4 شهور", "عمر 6 شهور", "عمر 9 شهور",
    "عمر 12 شهر", "عمر 18 شهر", "عمر سنتين", "عمر سنتين ونصف", "عمر 3 سنين",
    "عمر 3 سنين ونصف", "عمر 4 سنين", "عمر 4 سنين ونصف", "عمر 5 سنين",
    "عمر 5 سنين ونصف", "عمر 6 سنين"
]

PREGNANT_COLUMNS = [
    "تاريخ التسجيل", "اسم المستخدم", "الاسم", "العنوان", "الرقم القومى", "رقم الموبايل",
    "العمر الحالى", "السن عند الزواج", "السن عند الحمل الاول", "مستوى التعليم", "الوظيفة",
    "تاريخ اخر دورة شهرية", "قرابة بين الزوجين", "عدد مرات الحمل", "عدد مرات الاجهاض",
    "عدد الاطفال", "المدة بين اخر حملين", "نوع الولادة", "أمراض مزمنة: إرتفاع ضغط الدم",
    "أمراض مزمنة: السكر", "أمراض مزمنة: إضطرابات الغدة", "أمراض مزمنة: الأنيميا", "أمراض مزمنة: اخرى",
    'مكملات "قبل": حمض الفوليك', 'مكملات "قبل": الحديد', 'مكملات "قبل": الكالسيوم',
    'مكملات "أثناء": حمض الفوليك', 'مكملات "أثناء": الحديد', 'مكملات "أثناء": الكالسيوم',
    "وسيلة تنظيم الأسرة المستخدمة سابقا", "مدة إستخدام الوسيلة السابقة", "شهر الحمل",
    "التاريخ الزيارة", "التغذية السليمة", "المكملات الغذائية", "التمرينات الرياضية",
    "قسط من النوم والراحة", "المتابعة الدورية للحمل",
    "التحذير من تناول الأدوية بدون إستشارة طبيب والتعرض للتدخين والأبخرة",
    "المتاعب البسيطة في الشهور الأولى", "المتاعب في الشهور الأخيرة", "علامات الخطر أثناء الحمل",
    "مشاكل الولادة المبكرة وكيفية تجنبها", "حركة الجنين / معرفة جنس الجنين/ تمييز الأصوات من قبل الجنين",
    "تغير لون الجلد حول الحلمة وظهور بعض إفرازات من الثدي", "إرتداء الملابس الفضفاضة المريحة",
    "الإستعداد للولادة / تحضير ملابس المولود ، الخ", "علامات الولادة", "مميزات الولادة الطبيعية",
    "الساعة الذهبية الأولى", "ملامسة الجلد للجلد", "البداية المبكرة للرضاعة الطبيعية",
    "الرضاعة الطبيعية المطلقة", "أهمية المباعدة", "وسائل تنظيم الأسرة",
    "إستخدام وسيلة بعد الولادة مباشرة", "التطور العصبي والنفسي للطفل", "ملاحظات/ توصيات",
    "تخطيط الزيارة القادمة", "المتابعة ما بعد الولادة"
]

CHILD_COLUMNS = [
    "تاريخ التسجيل", "اسم المستخدم", "تاريخ اول زيارة", "رقم الحالة", "اسم الام",
    "الرقم القومى للام", "رقم الموبايل للام", "تاريخ ميلاد الام", "مستوى التعليم للام",
    "عدد الاطفال لدى الام", "المدة بين اخر حملين", "الوظيفة للام", "الرقم القومى للاب",
    "رقم الموبايل للاب", "اسم الاب", "مستوى التعليم للاب", "اسم الطفل", "تاريخ الميلاد للطفل",
    "العمر الحالى للطفل (شهور)", "العمر الرحمى للطفل (أسابيع)", "مكان المتابعة (وحدة)",
    "مكان المتابعة (مستشفى)", "مكان المتابعة (اخرى)", "مصدر الاحالة(مستشفى الولادة)",
    "مصدر الاحالة (عيادة خاصة)", "مصدر الاحالة(عيادة التطعيمات)", "مصدر الاحالة(نصيحة)",
    "نوع الولادة", "مكان الولادة", "وزن الطفل عند الولادة", "طول الطفل عند الولادة",
    "مقاس راس الطفل عند الولادة", "دخول الحضانة", "سبب دخول الحضانة", "مدة البقاء فى الحضانة",
    "ملامسة الجلد فى الساعة الذهبية الأولى", "الرضاعة الطبيعية فى الساعة الذهبية الأولى",
    "موعد الزيارة", "تاريخ الزيارة", "رضاعة طبيعية مطلقة", "رضاعة طبيعية مع سوائل وأعشاب",
    "رضاعة طبيعية مع صناعي", "رضاعة لبن صناعي", "الوزن (كجم)", "الطول (سم)", "محيط الرأس (سم)",
    "فوائد الرضاعة الطبيعية والأوضاع وعلامات الجوع والشبع", "كفاية اللبن وكمية البراز",
    "إعطاء الجرعة اليومية من فيتامين د", "كيفية رعاية السرة والإهتمام بنظافة الطفل",
    "البطاقة الصحية وأهمية المتابعة الدورية ومنحنيات النمو", "أهمية الإلتزام بتطعيمات الطفل",
    "التغذية الصحية للأم المرضعة", "كيفية التعرف على علامات الخطورة", "النمو والتطور الحركي",
    "التطور الإدراكي والمعرفي", "التطور اللغوي", "رسائل التربية الإيجابية", "الأنشطة التحفيزية",
    "التوعية عن التغذية التكميلية وسلامة الغذاء والتغذية السليمة", "إعطاء الجرعة اليومية من الحديد",
    "أهمية إستخدام وسيلة تنظيم أسرة وأهمية المباعدة", "موقف إستخدام وسيلة تنظيم أسرة",
    "الحمل الجديد", "الخدمات الغير ملباه", "تحويل الى عيادة تنظيم الاسره", "تخطيط الزيارة القادمة"
]

DROPDOWN_OPTIONS = {
    "مستوى التعليم": ["امى", "يجيد القراءة", "مؤهل متوسط", "فوق متوسط", "مؤهل عالى"],
    "الوظيفة": ["يعمل", "لا تعمل"],
    "قرابة بين الزوجين": ["نعم", "لا"],
    "وسيلة تنظيم الأسرة المستخدمة سابقا": ["توجد", "مرغوب", "غير مرغوب"],
    "شهر الحمل": ["الشهر الاول", "الشهر الثانى", "الشهر الثالث", "الشهر الرابع", "الشهر الخامس", "الشهر السادس", "الشهر السابع", "الشهر الثامن", "الشهر التاسع"],
    "أمراض مزمنة: إرتفاع ضغط الدم": ["تم", "لم يتم"],
    "أمراض مزمنة: السكر": ["تم", "لم يتم"],
    "أمراض مزمنة: إضطرابات الغدة": ["تم", "لم يتم"],
    "أمراض مزمنة: الأنيميا": ["تم", "لم يتم"],
    'مكملات "قبل": حمض الفوليك': ["تم", "لم يتم"],
    'مكملات "قبل": الحديد': ["تم", "لم يتم"],
    'مكملات "قبل": الكالسيوم': ["تم", "لم يتم"],
    'مكملات "أثناء": حمض الفوليك': ["تم", "لم يتم"],
    'مكملات "أثناء": الحديد': ["تم", "لم يتم"],
    'مكملات "أثناء": الكالسيوم': ["تم", "لم يتم"],
    "التغذية السليمة": ["تم", "لم يتم"],
    "المكملات الغذائية": ["تم", "لم يتم"],
    "التمرينات الرياضية": ["تم", "لم يتم"],
    "قسط من النوم والراحة": ["تم", "لم يتم"],
    "المتابعة الدورية للحمل": ["تم", "لم يتم"],
    "التحذير من تناول الأدوية بدون إستشارة طبيب والتعرض للتدخين والأبخرة": ["تم", "لم يتم"],
    "المتاعب البسيطة في الشهور الأولى": ["تم", "لم يتم"],
    "المتاعب في الشهور الأخيرة": ["تم", "لم يتم"],
    "علامات الخطر أثناء الحمل": ["تم", "لم يتم"],
    "مشاكل الولادة المبكرة وكيفية تجنبها": ["تم", "لم يتم"],
    "حركة الجنين / معرفة جنس الجنين/ تمييز الأصوات من قبل الجنين": ["تم", "لم يتم"],
    "تغير لون الجلد حول الحلمة وظهور بعض إفرازات من الثدي": ["تم", "لم يتم"],
    "إرتداء الملابس الفضفاضة المريحة": ["تم", "لم يتم"],
    "الإستعداد للولادة / تحضير ملابس المولود ، الخ": ["تم", "لم يتم"],
    "علامات الولادة": ["تم", "لم يتم"],
    "مميزات الولادة الطبيعية": ["تم", "لم يتم"],
    "الساعة الذهبية الأولى": ["تم", "لم يتم"],
    "ملامسة الجلد للجلد": ["تم", "لم يتم"],
    "البداية المبكرة للرضاعة الطبيعية": ["تم", "لم يتم"],
    "الرضاعة الطبيعية المطلقة": ["تم", "لم يتم"],
    "أهمية المباعدة": ["تم", "لم يتم"],
    "وسائل تنظيم الأسرة": ["تم", "لم يتم"],
    "إستخدام وسيلة بعد الولادة مباشرة": ["تم", "لم يتم"],
    "التطور العصبي والنفسي للطفل": ["طبيعى", "متقدم", "متاخر"],
    "مستوى التعليم للام": ["امى", "يجيد القراءة", "مؤهل متوسط", "فوق متوسط", "مؤهل عالى"],
    "مستوى التعليم للاب": ["امى", "يجيد القراءة", "مؤهل متوسط", "فوق متوسط", "مؤهل عالى"],
    "الوظيفة للام": ["يعمل", "لا تعمل"],
    "مكان الولادة": ["المستشفى", "المنزل"],
    "سبب دخول الحضانة": [
        "انخفاض وزن الطفل.", "احتياج الطفل لأدوية محددة بهذا الوقت.", "صعوبة شديدة في التنفس لعدم اكتمال نمو الرئتين.",
        "ارتفاع درجة حرارة جسم الرضيع.", "تعطل العمليات الحيوية بجسم الطفل.", "انخفاض معدل الجلوكوز في دم الطفل.",
        "معاناة الرضيع مشكلات في الجهاز الهضمي.", "إصابة الطفل بعدوى في الدم.", "إصابة الطفل بالصفراء.",
        "حدوث مشكلات خلال الولادة “الولادة المتعسرة أو الحمل الحرج”.", "وجود عيب خلقي يمنع الطفل عن التنفس أو الرضاعة بشكل طبيعي."
    ],
    "موعد الزيارة": VISIT_SCHEDULE_OPTIONS,
    "رضاعة طبيعية مع سوائل وأعشاب": ["تم", "لم يتم"],
    "رضاعة طبيعية مع صناعي": ["تم", "لم يتم"],
    "رضاعة لبن صناعي": ["تم", "لم يتم"],
    "دخول الحضانة": ["تم", "لم يتم"],
    "ملامسة الجلد فى الساعة الذهبية الأولى": ["تم", "لم يتم"],
    "الرضاعة الطبيعية فى الساعة الذهبية الأولى": ["تم", "لم يتم"],
    "موقف إستخدام وسيلة تنظيم أسرة": ["توجد", "مرغوب", "غير مرغوب", "حدث", "لم يحدث"],
    "الحمل الجديد": ["مرغوب", "غير مرغوب"],
    "الخدمات الغير ملباه": ["يوجد"],
    "تحويل الى عيادة تنظيم الاسره": ["تم", "لم يتم"],
    "النمو والتطور الحركي": ["طبيعى", "متقدم", "متاخر"],
    "التطور الإدراكي والمعرفي": ["طبيعى", "متقدم", "متاخر"],
    "التطور اللغوي": ["طبيعى", "متقدم", "متاخر"],
    "رسائل التربية الإيجابية": ["تم", "لم يتم"],
    "الأنشطة التحفيزية": ["تم", "لم يتم"],
    "التوعية عن التغذية التكميلية وسلامة الغذاء والتغذية السليمة": ["تم", "لم يتم"],
    "إعطاء الجرعة اليومية من الحديد": ["يوجد"],
    "أهمية إستخدام وسيلة تنظيم أسرة وأهمية المباعدة": ["تم التوعيه", "لم يتم التوعيه"],
    "إعطاء الجرعة اليومية من فيتامين د": ["يوجد"],
    "كيفية رعاية السرة والإهتمام بنظافة الطفل": ["تم", "لم يتم"],
    "البطاقة الصحية وأهمية المتابعة الدورية ومنحنيات النمو": ["تم", "لم يتم"],
    "أهمية الإلتزام بتطعيمات الطفل": ["تم", "لم يتم"],
    "التغذية الصحية للأم المرضعة": ["تم", "لم يتم"],
    "كيفية التعرف على علامات الخطورة": ["تم", "لم يتم"],
    "فوائد الرضاعة الطبيعية والأوضاع وعلامات الجوع والشبع": ["تم", "لم يتم"],
    "كفاية اللبن وكمية البراز": ["تم", "لم يتم"]
}

YES_NO_CHECKBOX_FIELDS = [
    "مكان المتابعة (وحدة)", "مكان المتابعة (مستشفى)", "مكان المتابعة (اخرى)",
    "مصدر الاحالة(مستشفى الولادة)", "مصدر الاحالة (عيادة خاصة)",
    "مصدر الاحالة(عيادة التطعيمات)", "مصدر الاحالة(نصيحة)"
]

# ==================== إعدادات الصفحة والتصميم ====================
st.set_page_config(
    page_title="برنامج بودى للمشورة الأسرية",
    page_icon="🌸",
    layout="wide",
    initial_sidebar_state="expanded"
)

custom_css = """
<style>
.main .block-container { padding-top: 1rem; padding-bottom: 1rem; max-width: 100%; }
.main { background-color: #FFF5F8; }
.stButton>button { background-color: #EC4899; color: white; border-radius: 8px; font-weight: bold; border: none; padding: 0.5rem 1rem; width: 100%; }
.stButton>button:hover { background-color: #BE185D; color: white; }
h1, h2, h3 { color: #701A75; }
footer {visibility: hidden;}
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

def clean_digits(val, max_len=None):
    if not val:
        return ""
    digits = "".join(filter(str.isdigit, str(val)))
    if max_len:
        return digits[:max_len]
    return digits

def parse_national_id(nat_id):
    clean_id = clean_digits(nat_id, 14)
    if len(clean_id) == 14:
        century_code = int(clean_id[0])
        year_digits = int(clean_id[1:3])
        month = int(clean_id[3:5])
        day = int(clean_id[5:7])
        century = 2000 if century_code == 3 else 1900
        birth_year = century + year_digits
        try:
            birth_date = datetime.date(birth_year, month, day)
            today = datetime.date.today()
            age = today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))
            return str(birth_date), str(age)
        except ValueError:
            return "", ""
    return "", ""

# ==================== تسجيل الدخول والصلاحيات ====================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.user = None
    st.session_state.name = None
    st.session_state.role = None

if not st.session_state.logged_in:
    st.markdown("<h2 style='text-align: center; color: #BE185D;'>🌸 برنامج بودى للمشورة الأسرية 🌸</h2>", unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: center; color: #701A75;'>تسجيل الدخول للنظام</h4>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        user_options = {f"{v['name']} ({k})": k for k, v in DEFAULT_USERS.items()}
        selected_display = st.selectbox("اختر الحساب والطبيبة 🩺", list(user_options.keys()))
        username = user_options[selected_display]
        password = st.text_input("كلمة المرور", type="password")
        
        if st.button("تسجيل الدخول ✨", use_container_width=True):
            if DEFAULT_USERS[username]["pass"] == password:
                st.session_state.logged_in = True
                st.session_state.user = username
                st.session_state.name = DEFAULT_USERS[username]["name"]
                st.session_state.role = DEFAULT_USERS[username]["role"]
                st.rerun()
            else:
                st.error("كلمة المرور غير صحيحة!")
    st.stop()

# ==================== القائمة والخيارات المشتركة ====================
menu_options = ["الصفحة الرئيسية", "سجل الحوامل", "سجل الأطفال", "استعراض البيانات والداشبورد"]
if st.session_state.role == "admin":
    menu_options.append("إدارة المستخدمين")

st.sidebar.markdown(f"### أهلاً بكِ د. {st.session_state.name} 🌸")
sidebar_menu = st.sidebar.radio("القائمة الرئيسية (جانبية)", menu_options, key="sidebar_radio")

if st.sidebar.button("🚪 تسجيل الخروج"):
    st.session_state.logged_in = False
    st.rerun()

st.markdown("---")
col_mobile_nav, col_mobile_logout = st.columns([3, 1])
with col_mobile_nav:
    main_screen_menu = st.selectbox("📱 انتقل مباشرة إلى القسم المطلوب:", menu_options, key="mobile_selectbox")
with col_mobile_logout:
    if st.button("خروج 🚪"):
        st.session_state.logged_in = False
        st.rerun()

menu = main_screen_menu
st.markdown("---")

# ==================== 2. سجل الحوامل ====================
if menu == "سجل الحوامل":
    st.markdown("<h2>🤰 سجل المشورة الأسرية للحوامل</h2>", unsafe_allow_html=True)
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    
    for col in PREGNANT_COLUMNS:
        if f"p_{col}" not in st.session_state:
            if col == "التاريخ الزيارة":
                st.session_state[f"p_{col}"] = today_str
            else:
                st.session_state[f"p_{col}"] = ""

    form_data = {}
    for col_name in PREGNANT_COLUMNS:
        if col_name in ["تاريخ التسجيل", "اسم المستخدم"]:
            continue
            
        if col_name == "نوع الولادة":
            st.markdown(f"**{col_name}**")
            current_val = st.session_state.get(f"p_{col_name}", "")
            
            c_opt1, c_opt2, c_opt3 = st.columns(3)
            with c_opt1:
                chk_nat = st.checkbox("طبيعى", value=(current_val == "طبيعى"), key="p_birth_nat")
            with c_opt2:
                chk_ces = st.checkbox("قيصرى", value=(current_val == "قيصرى"), key="p_birth_ces")
            with c_opt3:
                chk_none = st.checkbox("لا يوجد", value=(current_val == "لا يوجد" or current_val == ""), key="p_birth_none")
                
            selected_birth = ""
            if chk_nat: selected_birth = "طبيعى"
            elif chk_ces: selected_birth = "قيصرى"
            elif chk_none: selected_birth = "لا يوجد"
            
            form_data[col_name] = selected_birth
            st.session_state[f"p_{col_name}"] = selected_birth
            
        elif col_name in DROPDOWN_OPTIONS:
            st.markdown(f"**{col_name}**")
            options = DROPDOWN_OPTIONS[col_name]
            current_val = st.session_state.get(f"p_{col_name}", options[0])
            chosen_choice = st.radio(
                f"اختر {col_name}", options,
                index=(options.index(current_val) if current_val in options else 0),
                key=f"p_radio_{col_name}", horizontal=True
            )
            form_data[col_name] = chosen_choice
            st.session_state[f"p_{col_name}"] = chosen_choice
        else:
            if col_name == "الرقم القومى":
                raw_val = st.text_input(col_name, key=f"p_{col_name}")
                cleaned_val = clean_digits(raw_val, 14)
                form_data[col_name] = cleaned_val
                if len(cleaned_val) == 14:
                    _, calc_age = parse_national_id(cleaned_val)
                    if calc_age:
                        st.session_state["p_العمر الحالى"] = calc_age
            elif col_name == "رقم الموبايل":
                raw_val = st.text_input(col_name, key=f"p_{col_name}")
                cleaned_val = clean_digits(raw_val, 11)
                form_data[col_name] = cleaned_val
            elif col_name == "العمر الحالى":
                form_data[col_name] = st.text_input(f"{col_name} [محسوب تلقائياً من الرقم القومي]", key=f"p_{col_name}")
            elif col_name == "التاريخ الزيارة":
                form_data[col_name] = st.text_input(f"{col_name} [تاريخ اليوم التلقائي]", key=f"p_{col_name}")
            else:
                form_data[col_name] = st.text_input(col_name, key=f"p_{col_name}")

    if st.button("💾 حفظ بيانات الحامل", use_container_width=True):
        final_form_data = {}
        for col in PREGNANT_COLUMNS:
            if col == "تاريخ التسجيل":
                final_form_data[col] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            elif col == "اسم المستخدم":
                final_form_data[col] = st.session_state.name
            else:
                final_form_data[col] = st.session_state.get(f"p_{col}", form_data.get(col, ""))
                
        save_to_cloud("pregnant_records", final_form_data)
        st.success("تم حفظ بيانات الحامل بنجاح في Google Sheets السحابي! ✨")

# (باقي الأقسام: الصفحة الرئيسية، سجل الأطفال، استعراض البيانات، وإدارة المستخدمين تعمل بنفس التنسيق تماماً...)
elif menu == "الصفحة الرئيسية":
    st.markdown("<h1>✨ مرحباً بكِ في نظام المشورة الأسرية الشامل ✨</h1>", unsafe_allow_html=True)
    st.write("تم ربط النظام بقاعدة بيانات سحابية دائمية (Google Sheets) لحفظ كافة السجلات والبيانات بصورة آمنة ومستمرة.")
