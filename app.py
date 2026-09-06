import streamlit as st
import pandas as pd
import sqlite3
import datetime

# ==================== إعدادات الصفحة والتصميم ====================
st.set_page_config(
    page_title="برنامج بودى للمشورة الأسرية",
    page_icon="🌸",
    layout="wide",
    initial_sidebar_state="expanded"
)

custom_css = """
<style>
.main .block-container {
    padding-top: 1rem;
    padding-bottom: 1rem;
    max-width: 100%;
}
.main {
    background-color: #FFF5F8;
}
.stButton>button {
    background-color: #EC4899;
    color: white;
    border-radius: 8px;
    font-weight: bold;
    border: none;
    padding: 0.5rem 1rem;
    width: 100%;
}
.stButton>button:hover {
    background-color: #BE185D;
    color: white;
}
h1, h2, h3 {
    color: #701A75;
}
footer {visibility: hidden;}
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ==================== إعداد قاعدة البيانات SQLite الثابتة ====================
DB_FILE = "family_counseling.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS pregnant_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            "تاريخ التسجيل" TEXT,
            "اسم المستخدم" TEXT,
            "الرقم القومى" TEXT,
            data_json TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS child_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            "تاريخ التسجيل" TEXT,
            "اسم المستخدم" TEXT,
            "الرقم القومى للام" TEXT,
            data_json TEXT
        )
    ''')
    conn.commit()
    conn.close()

init_db()

def save_to_db(table_name, df):
    conn = sqlite3.connect(DB_FILE)
    df.to_sql(table_name, conn, if_exists='replace', index=False)
    conn.close()

def load_from_db(table_name):
    conn = sqlite3.connect(DB_FILE)
    try:
        df = pd.read_sql(f"SELECT * FROM {table_name}", conn, dtype=str)
        if "index" in df.columns:
            df = df.drop(columns=["index"])
        conn.close()
        return df
    except Exception:
        conn.close()
        return pd.DataFrame()

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
        "انخفاض وزن الطفل.", "احتياج الطفل لأدوية محددة بهذا الوقت.",
        "صعوبة شديدة في التنفس لعدم اكتمال نمو الرئتين.", "ارتفاع درجة حرارة جسم الرضيع.",
        "تعطل العمليات الحيوية بجسم الطفل.", "انخفاض معدل الجلوكوز في دم الطفل.",
        "معاناة الرضيع مشكلات في الجهاز الهضمي.", "إصابة الطفل بعدوى في الدم.",
        "إصابة الطفل بالصفراء.", "حدوث مشكلات خلال الولادة “الولادة المتعسرة أو الحمل الحرج”.",
        "وجود عيب خلقي يمنع الطفل عن التنفس أو الرضاعة بشكل طبيعي."
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
    "كفاية اللبن وكمية البراز": ["تم", "لم يتم"],
}

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
    "الرضاعة الطبيعية المطلقة", "أهمية المباعدة", "وسائل تنظيم الأسرة", "إستخدام وسيلة بعد الولادة مباشرة",
    "التطور العصبي والنفسي للطفل", "ملاحظات/ توصيات", "تخطيط الزيارة القادمة", "المتابعة ما بعد الولادة"
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

YES_NO_CHECKBOX_FIELDS = [
    "مكان المتابعة (وحدة)", "مكان المتابعة (مستشفى)", "مكان المتابعة (اخرى)",
    "مصدر الاحالة(مستشفى الولادة)", "مصدر الاحالة (عيادة خاصة)",
    "مصدر الاحالة(عيادة التطعيمات)", "مصدر الاحالة(نصيحة)"
]

def clean_digits(val, max_len=None):
    if not val: return ""
    digits = "".join(filter(str.isdigit, str(val)))
    return digits[:max_len] if max_len else digits

def parse_national_id(nat_id):
    clean_id = clean_digits(nat_id, 14)
    if len(clean_id) == 14:
        century_code = int(clean_id[0])
        year_digits = int(clean_id[1:3])
        month = int(clean_id[3:5])
        day = int(clean_id[5:7])
        century = 2000 if century_code == 3 else 1900
        try:
            birth_date = datetime.date(century + year_digits, month, day)
            today = datetime.date.today()
            age = today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))
            return str(birth_date), str(age)
        except ValueError:
            return "", ""
    return "", ""

def get_existing_data(nat_id, table_name, id_column):
    clean_id = clean_digits(nat_id, 14)
    df = load_from_db(table_name)
    if not df.empty and id_column in df.columns and len(clean_id) == 14:
        match = df[df[id_column].astype(str).str.strip() == clean_id]
        if not match.empty:
            return match.iloc[-1].to_dict()
    return {}

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

# ==================== 1. الصفحة الرئيسية ====================
if menu == "الصفحة الرئيسية":
    st.markdown("<h1>✨ مرحباً بكِ في نظام المشورة الأسرية الشامل ✨</h1>", unsafe_allow_html=True)
    st.write("النظام يعمل بقاعدة بيانات SQLite مستقرة وتم ضبط التعبئة التلقائية لتاريخ ميلاد الأم من الرقم القومي أوتوماتيكياً.")

# ==================== 2. سجل الحوامل ====================
elif menu == "سجل الحوامل":
    st.markdown("<h2>🤰 سجل المشورة الأسرية للحوامل</h2>", unsafe_allow_html=True)
    today_str = datetime.date.today().strftime("%Y-%m-%d")

    for col in PREGNANT_COLUMNS:
        if f"p_{col}" not in st.session_state:
            st.session_state[f"p_{col}"] = today_str if col == "التاريخ الزيارة" else ""

    form_data = {}
    for i, col_name in enumerate(PREGNANT_COLUMNS):
        if col_name in ["تاريخ التسجيل", "اسم المستخدم"]:
            continue
        
        unique_key = f"p_input_field_{i}_{col_name}"
        
        if col_name == "نوع الولادة":
            st.markdown(f"**{col_name}**")
            current_val = st.session_state.get(f"p_{col_name}", "")
            c1, c2, c3 = st.columns(3)
            with c1: chk_nat = st.checkbox("طبيعى", value=(current_val == "طبيعى"), key=f"{unique_key}_nat")
            with c2: chk_ces = st.checkbox("قيصرى", value=(current_val == "قيصرى"), key=f"{unique_key}_ces")
            with c3: chk_none = st.checkbox("لا يوجد", value=(current_val in ["لا يوجد", ""]), key=f"{unique_key}_none")
            
            selected_birth = "طبيعى" if chk_nat else ("قيصرى" if chk_ces else "لا يوجد")
            form_data[col_name] = selected_birth
            st.session_state[f"p_{col_name}"] = selected_birth

        elif col_name in DROPDOWN_OPTIONS:
            st.markdown(f"**{col_name}**")
            options = DROPDOWN_OPTIONS[col_name]
            current_val = st.session_state.get(f"p_{col_name}", options[0])
            
            cols_checkboxes = st.columns(min(len(options), 4))
            selected_value = current_val
            
            for idx, opt in enumerate(options):
                col_target = cols_checkboxes[idx % len(cols_checkboxes)]
                with col_target:
                    is_checked = st.checkbox(opt, value=(current_val == opt), key=f"{unique_key}_chk_{idx}")
                    if is_checked:
                        selected_value = opt
            
            form_data[col_name] = selected_value
            st.session_state[f"p_{col_name}"] = selected_value
        else:
            if col_name == "الرقم القومى":
                raw_val = st.text_input(col_name, key=unique_key)
                cleaned_val = clean_digits(raw_val, 14)
                form_data[col_name] = cleaned_val
                if len(cleaned_val) == 14:
                    _, calc_age = parse_national_id(cleaned_val)
                    if calc_age: st.session_state["p_العمر الحالى"] = calc_age
            elif col_name == "رقم الموبايل":
                form_data[col_name] = clean_digits(st.text_input(col_name, key=unique_key), 11)
            else:
                form_data[col_name] = st.text_input(col_name, key=unique_key)

    if st.button("💾 حفظ بيانات الحامل", use_container_width=True):
        final_form_data = {
            "تاريخ التسجيل": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "اسم المستخدم": st.session_state.name
        }
        for col in PREGNANT_COLUMNS:
            if col not in ["تاريخ التسجيل", "اسم المستخدم"]:
                final_form_data[col] = st.session_state.get(f"p_{col}", form_data.get(col, ""))

        df_existing = load_from_db("pregnant_records")
        new_df = pd.DataFrame([final_form_data], dtype=str)
        for col in PREGNANT_COLUMNS:
            if col not in new_df.columns: new_df[col] = ""
        new_df = new_df[PREGNANT_COLUMNS]
        
        combined_df = pd.concat([df_existing, new_df], ignore_index=True) if not df_existing.empty else new_df
        save_to_db("pregnant_records", combined_df)
        st.success("تم حفظ بيانات الحامل بنجاح في قاعدة البيانات! ✨")

# ==================== 3. سجل الأطفال ====================
elif menu == "سجل الأطفال":
    st.markdown("<h2>👶 سجل المشورة الأسرية للأطفال</h2>", unsafe_allow_html=True)
    today_str = datetime.date.today().strftime("%Y-%m-%d")

    for col in CHILD_COLUMNS:
        if f"c_{col}" not in st.session_state:
            st.session_state[f"c_{col}"] = today_str if col in ["تاريخ الزيارة", "تاريخ اول زيارة"] else ""

    # حقل الرقم القومي للأم مع حساب تلقائي وفوري لتاريخ الميلاد
    raw_nat_id_mom = st.text_input("الرقم القومى للام", key="c_الرقم القومى للام_input")
    nat_id_mom_input = clean_digits(raw_nat_id_mom, 14)
    
    if nat_id_mom_input:
        st.session_state["c_الرقم القومى للام"] = nat_id_mom_input
        b_date_mom, _ = parse_national_id(nat_id_mom_input)
        if b_date_mom:
            st.session_state["c_تاريخ ميلاد الام"] = b_date_mom

    # زر استرجاع البيانات المسجلة للأسرة فقط (بدون زر حساب تاريخ الميلاد)
    if st.button("🔍 استرجاع بيانات الأسرة المسجلة", use_container_width=True):
        if len(nat_id_mom_input) == 14:
            found_data = get_existing_data(nat_id_mom_input, "child_records", "الرقم القومى للام")
            if not found_data:
                found_data = get_existing_data(nat_id_mom_input, "pregnant_records", "الرقم القومى")
            for c_name in CHILD_COLUMNS:
                if c_name not in ["تاريخ التسجيل", "اسم المستخدم", "الرقم القومى للام"]:
                    val = found_data.get(c_name, "")
                    if val: st.session_state[f"c_{c_name}"] = str(val)
            st.success("تم استرجاع البيانات بنجاح!")
            st.rerun()

    for i, col_name in enumerate(CHILD_COLUMNS):
        if col_name in ["تاريخ التسجيل", "اسم المستخدم", "الرقم القومى للام"]:
            continue
        
        unique_key = f"c_input_field_{i}_{col_name}"
        
        if col_name == "تاريخ ميلاد الام":
            auto_mom_bdate = st.session_state.get("c_تاريخ ميلاد الام", "")
            st.text_input(col_name, value=auto_mom_bdate, key=unique_key)
            st.session_state[f"c_{col_name}"] = auto_mom_bdate
            continue

        if col_name == "نوع الولادة":
            st.markdown(f"**{col_name}**")
            current_val = st.session_state.get(f"c_{col_name}", "")
            c1, c2, c3 = st.columns(3)
            with c1: chk_nat = st.checkbox("طبيعى", value=(current_val == "طبيعى"), key=f"{unique_key}_nat")
            with c2: chk_ces = st.checkbox("قيصرى", value=(current_val == "قيصرى"), key=f"{unique_key}_ces")
            with c3: chk_none = st.checkbox("لا يوجد", value=(current_val in ["لا يوجد", ""]), key=f"{unique_key}_none")
            st.session_state[f"c_{col_name}"] = "طبيعى" if chk_nat else ("قيصرى" if chk_ces else "لا يوجد")

        elif col_name == "رضاعة طبيعية مطلقة":
            st.markdown(f"**{col_name}**")
            c1, c2, c3 = st.columns(3)
            current_val = st.session_state.get(f"c_{col_name}", "")
            with c1: chk_3 = st.checkbox("3 شهور", value=(current_val == "3 شهور"), key=f"{unique_key}_3")
            with c2: chk_4 = st.checkbox("4 شهور", value=(current_val == "4 شهور"), key=f"{unique_key}_4")
            with c3: chk_6 = st.checkbox("6 شهور", value=(current_val == "6 شهور"), key=f"{unique_key}_6")
            st.session_state[f"c_{col_name}"] = "3 شهور" if chk_3 else ("4 شهور" if chk_4 else ("6 شهور" if chk_6 else ""))

        elif col_name in YES_NO_CHECKBOX_FIELDS:
            checked = st.checkbox(col_name, value=False, key=unique_key)
            st.session_state[f"c_{col_name}"] = "نعم" if checked else ""

        elif col_name in DROPDOWN_OPTIONS:
            options = DROPDOWN_OPTIONS[col_name]
            if col_name == "سبب دخول الحضانة" and st.session_state.get("c_دخول الحضانة", "لم يتم") == "لم يتم":
                st.session_state[f"c_{col_name}"] = ""
                continue
            
            st.markdown(f"**{col_name}**")
            current_val = st.session_state.get(f"c_{col_name}", options[0])
            
            cols_checkboxes = st.columns(min(len(options), 4))
            selected_value = current_val
            
            for idx, opt in enumerate(options):
                col_target = cols_checkboxes[idx % len(cols_checkboxes)]
                with col_target:
                    is_checked = st.checkbox(opt, value=(current_val == opt), key=f"{unique_key}_chk_{idx}")
                    if is_checked:
                        selected_value = opt
            
            st.session_state[f"c_{col_name}"] = selected_value
        else:
            if col_name == "الرقم القومى للاب":
                st.session_state[f"c_{col_name}"] = clean_digits(st.text_input(col_name, key=unique_key), 14)
            elif col_name in ["رقم الموبايل للام", "رقم الموبايل للاب"]:
                st.session_state[f"c_{col_name}"] = clean_digits(st.text_input(col_name, key=unique_key), 11)
            elif col_name == "تاريخ الميلاد للطفل":
                chosen_date = st.date_input(col_name, value=datetime.date.today(), key=unique_key)
                st.session_state[f"c_{col_name}"] = str(chosen_date)
            else:
                st.text_input(col_name, key=unique_key)

    if st.button("💾 حفظ بيانات الطفل", use_container_width=True):
        final_child_data = {
            "تاريخ التسجيل": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "اسم المستخدم": st.session_state.name
        }
        for col in CHILD_COLUMNS:
            if col not in ["تاريخ التسجيل", "اسم المستخدم"]:
                if col == "سبب دخول الحضانة" and st.session_state.get("c_دخول الحضانة", "لم يتم") == "لم يتم":
                    final_child_data[col] = ""
                else:
                    final_child_data[col] = st.session_state.get(f"c_{col}", "")

        df_existing = load_from_db("child_records")
        new_child_df = pd.DataFrame([final_child_data], dtype=str)
        for col in CHILD_COLUMNS:
            if col not in new_child_df.columns: new_child_df[col] = ""
        new_child_df = new_child_df[CHILD_COLUMNS]
        
        combined_df = pd.concat([df_existing, new_child_df], ignore_index=True) if not df_existing.empty else new_child_df
        save_to_db("child_records", combined_df)
        st.success("تم حفظ بيانات الطفل بنجاح في قاعدة البيانات! ✨")

# ==================== 4. استعراض البيانات والداشبورد ====================
elif menu == "استعراض البيانات والداشبورد":
    st.markdown("<h2>📊 لوحة المؤشرات واستعراض البيانات المبسطة</h2>", unsafe_allow_html=True)
    
    sheet_to_show = st.selectbox("اختر السجل للاستعراض:", ["المشورة الاسرية للحامل", "سجل المشورة للاطفال"])
    db_table_name = "pregnant_records" if sheet_to_show == "المشورة الاسرية للحامل" else "child_records"
    df_view = load_from_db(db_table_name)

    if not df_view.empty:
        st.dataframe(df_view, use_container_width=True)

        if st.session_state.role == "admin":
            st.markdown("### 🗑️ لوحة التحكم الإدارية (حذف السجلات)")
            row_idx_to_delete = st.number_input("أدخل رقم الصف المراد حذفه:", min_value=0, max_value=max(0, len(df_view)-1), step=1)
            if st.button("🗑️ حذف هذا الصف"):
                df_main = load_from_db(db_table_name)
                df_main = df_main.drop(df_main.index[row_idx_to_delete])
                save_to_db(db_table_name, df_main)
                st.success("تم الحذف بنجاح!")
                st.rerun()
    else:
        st.warning("لا توجد بيانات مسجلة حتى الآن.")

# ==================== 5. إدارة المستخدمين ====================
elif menu == "إدارة المستخدمين" and st.session_state.role == "admin":
    st.markdown("<h2>⚙️ إدارة المستخدمين والصلاحيات</h2>", unsafe_allow_html=True)
    for k, v in DEFAULT_USERS.items():
        st.write(f"- **{v['name']}** | اسم المستخدم: `{k}` | الصلاحية: `{v['role']}`")
