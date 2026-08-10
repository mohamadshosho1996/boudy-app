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

# تنسيق CSS مع تأثيرات تطاير القلوب الحمراء والقلب الكبير
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

/* تصميم القلب الأحمر الكبير النابض */
.big-heart-container {
    display: flex;
    justify-content: center;
    align-items: center;
    margin: 20px 0;
}
.big-heart {
    position: relative;
    width: 160px;
    height: 140px;
    background-color: #ff3366;
    transform: rotate(-45deg);
    box-shadow: 0 10px 25px rgba(255, 51, 102, 0.5);
    animation: heartbeat 1s infinite;
}
.big-heart::before,
.big-heart::after {
    content: "";
    position: absolute;
    width: 160px;
    height: 140px;
    background-color: #ff3366;
    border-radius: 50%;
}
.big-heart::before {
    top: -80px;
    left: 0;
}
.big-heart::after {
    left: 80px;
    top: 0;
}
.heart-text {
    position: absolute;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%) rotate(45deg);
    color: white;
    font-size: 22px;
    font-weight: bold;
    text-align: center;
    width: 140px;
    z-index: 10;
    text-shadow: 1px 1px 3px rgba(0,0,0,0.4);
}
@keyframes heartbeat {
    0% { transform: scale(1) rotate(-45deg); }
    15% { transform: scale(1.1) rotate(-45deg); }
    30% { transform: scale(1) rotate(-45deg); }
    45% { transform: scale(1.15) rotate(-45deg); }
    60% { transform: scale(1) rotate(-45deg); }
}

/* تأثير تطاير القلوب الكثيرة */
.floating-hearts-overlay {
    position: fixed;
    top: 0;
    left: 0;
    width: 100vw;
    height: 100vh;
    pointer-events: none;
    z-index: 99999;
    overflow: hidden;
}
.f-heart {
    position: absolute;
    color: #ff2255;
    font-size: 24px;
    animation: flyUp 3s linear forwards;
    opacity: 0;
}
@keyframes flyUp {
    0% {
        transform: translateY(100vh) scale(0.5) rotate(0deg);
        opacity: 1;
    }
    50% {
        opacity: 1;
    }
    100% {
        transform: translateY(-10vh) scale(1.5) rotate(360deg);
        opacity: 0;
    }
}
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ==================== الثوابت وإعدادات البيانات ====================
EXCEL_FILE = "template.xlsx"

DEFAULT_USERS = {
    "admin": {"pass": "admin123", "role": "admin", "name": "د. شيماء 🌸"},
    "user1": {"pass": "1234", "role": "user", "name": "د. علا 🎀"},
    "user2": {"pass": "1234", "role": "user", "name": "د. عبير 🎀"},
    "user3": {"pass": "1234", "role": "user", "name": "د. ايه 🎀"}
}

VISIT_SCHEDULE_OPTIONS = [
    "الاسبوع الاول", "عمر شهرين", "عمر 4 شهور", "عمر 6 شهور",
    "عمر 9 شهور", "عمر 12 شهر", "عمر 18 شهر", "عمر سنتين",
    "عمر سنتين ونصف", "عمر 3 سنين", "عمر 3 سنين ونصف", "عمر 4 سنين",
    "عمر 4 سنين ونصف", "عمر 5 سنين", "عمر 5 سنين ونصف", "عمر 6 سنين"
]

DROPDOWN_OPTIONS = {
    "مستوى التعليم": ["امى", "يجيد القراءة", "مؤهل متوسط", "فوق متوسط", "مؤهل عالى"],
    "الوظيفة": ["يعمل", "لا تعمل"],
    "نوع الولادة": ["طبيعى", "قيصرى"],
    "قرابة بين الزوجين": ["نعم", "لا"],
    "وسيلة تنظيم الاسرة المستخدمة سابقا": ["توجد", "لا يوجد", "مرغوب", "غير مرغوب"],
    "شهر الحمل": ["الشهر الاول", "الشهر الثانى", "الشهر الثالث", "الشهر الرابع", "الشهر الخامس", "الشهر السادس", "الشهر السابع", "الشهر الثامن", "الشهر التاسع"],
    "ارتفاع ضغط الدم": ["تم", "لم يتم"],
    "السكر": ["تم", "لم يتم"],
    "اضرابات الغدة": ["تم", "لم يتم"],
    "الانيميا": ["تم", "لم يتم"],
    "حمض الفوليك قبل الحمل": ["لا يوجد", "تم", "لم يتم"],
    "الحديد قبل الحمل": ["لا يوجد", "تم", "لم يتم"],
    "الكالسيوم قبل الحمل": ["لا يوجد", "تم", "لم يتم"],
    "حمض الفوليك اثناء الحمل": ["تم", "لم يتم"],
    "الحديد اثناء الحمل": ["تم", "لم يتم"],
    "الكالسيوم اثناء الحمل": ["تم", "لم يتم"],
    "التغذية السليمة": ["تم", "لم يتم"],
    "المكملات الغذائية": ["تم", "لم يتم"],
    "التمرينات الرياضية": ["تم", "لم يتم"],
    "قسط من النوم والراحة": ["تم", "لم يتم"],
    "المتابعة الدورية للحمل": ["تم", "لم يتم"],
    "التحذير من تناول الادوية": ["تم", "لم يتم"],
    "المتاعب البسيطة فى الشهور الاولى": ["تم", "لم يتم"],
    "المتاعب فى الشهور الاخيرة": ["تم", "لم يتم"],
    "علامات الخطر اثناء الحمل": ["تم", "لم يتم"],
    "مشاكل الولادة المبكرة": ["تم", "لم يتم"],
    "حركة الجنين ومعرفة الجنس": ["تم", "لم يتم"],
    "تغير لون الجلد حول الحلمة": ["تم", "لم يتم"],
    "ارتداء الملابس الفضفاضة": ["تم", "لم يتم"],
    "الاستعداد للولاده": ["تم", "لم يتم"],
    "علامات الولاده": ["تم", "لم يتم"],
    "مميزات الولادة الطبيعية": ["تم", "لم يتم"],
    "الساعة الذهبية الاولى": ["تم", "لم يتم"],
    "ملامسة الجلد للجلد": ["تم", "لم يتم"],
    "البداية المبكرة للرضاعه الطبيعية": ["تم", "لم يتم"],
    "الرضاعة الطبيعية المطلقة": ["تم", "لم يتم"],
    "اهمية المباعدة": ["تم", "لم يتم"],
    "وسائل تنظيم الاسره": ["تم", "لم يتم"],
    "استخدام وسيلة بعد الولادة مباشرة": ["تم", "لم يتم"],
    "التطور العصبى والنفسى للطفل": ["طبيعى", "متقدم", "متاخر"],
    "مستوى التعليم للام": ["امى", "يجيد القراءة", "مؤهل متوسط", "فوق متوسط", "مؤهل عالى"],
    "مستوى التعليم للاب": ["امى", "يجيد القراءة", "مؤهل متوسط", "فوق متوسط", "مؤهل عالى"],
    "وظيفة الام": ["يعمل", "لا تعمل"],
    "مكان الولادة": ["المستشفى", "المنزل"],
    "سبب دخول الحضانة": [
        "انخفاض وزن الطفل.", "احتياج الطفل لأدوية محددة بهذا الوقت.", "صعوبة شديدة في التنفس لعدم اكتمال نمو الرئتين.",
        "ارتفاع درجة حرارة جسم الرضيع.", "تعطل العمليات الحيوية بجسم الطفل.", "انخفاض معدل الجلوكوز في دم الطفل.",
        "معاناة الرضيع مشكلات في الجهاز الهضمي.", "إصابة الطفل بعدوى في الدم.", "إصابة الطفل بالصفراء.",
        "حدوث مشكلات خلال الولادة “الولادة المتعسرة أو الحمل الحرج”.", "وجود عيب خلقي يمنع الطفل عن التنفس أو الرضاعة بشكل طبيعي."
    ],
    "مكان المتابعة": ["وحدة", "مستشفى", "أخرى"],
    "مصدر الاحالة": ["مستشفى الولادة", "عيادة خاصة", "عيادة التطعيمات", "نصيحة"],
    "موعد الزيارة": VISIT_SCHEDULE_OPTIONS,
    "رضاعة طبيعية مطلقة": ["3 شهور", "4 شهور", "6 شهور"],
    "رضاعة طبيعية مع سوائل": ["تم", "لم يتم"],
    "رضاعة طبيعية مع صناعى": ["تم", "لم يتم"],
    "رضاعة لبن صناعى": ["تم", "لم يتم"],
    "دخول الحضانه": ["تم", "لم يتم"],
    "ملامسة الجلد فى الساعه الذهبية الاولى": ["تم", "لم يتم"],
    "موقف استخدام الوسيلة": ["توجد", "لا يوجد", "مرغوب", "غير مرغوب", "حدث", "لم يحدث"],
    "الحمل الجديد": ["مرغوب", "غير مرغوب"],
    "الخدمات الغير ملباه": ["لا يوجد", "يوجد"],
    "النمو والتطور الحركى": ["طبيعى", "متقدم", "متاخر"],
    "التطور الادراكى والمعرفى": ["طبيعى", "متقدم", "متاخر"],
    "التطور اللغوى": ["طبيعى", "متقدم", "متاخر"],
    "رسائل التربية الايجابية": ["تم", "لم يتم"],
    "الانشطة التحفيزية": ["تم", "لم يتم"],
    "التوعية عن التغذية التكميلية وسلامة الغذاء": ["تم", "لم يتم"],
    "اعطاء جرعة الحديد اليومية": ["يوجد", "لا يوجد"],
    "اهمية استخدام الوسيلة": ["تم التوعية", "غير مهتمة", "مرفوض"],
    "اعطاء الجرعة اليومية من فيتامين د": ["يوجد", "لا يوجد"],
    "كيفية رعاية الاسرة": ["تم", "لم يتم"],
    "البطاقة الصحية واهمية المتابعة": ["تم", "لم يتم"],
    "اهمية الالتزام بالتطعيمات": ["تم", "لم يتم"],
    "التغذيه الصحيحة للام": ["تم", "لم يتم"],
    "كيفية التعرف على علامات الخطر": ["تم", "لم يتم"],
    "فوائد الرضاعه الطبيعية": ["تم", "لم يتم"],
    "كفاية اللبن وكمية البراز": ["تم", "لم يتم"]
}

PREGNANT_COLUMNS = [
    "تاريخ التسجيل", "اسم المستخدم", "الاسم", "العنوان", "الرقم القومى", "رقم الموبايل", "العمر الحالى", "العمر عند الزواج",
    "السن عند الحمل الاول", "مستوى التعليم", "الوظيفة", "تاريخ اخر دورة شهرية", "قرابة بين الزوجين", "عدد مرات الحمل",
    "عدد مرات الاجهاض", "عدد الاطفال", "المدة بين اخر حملين", "نوع الولادة", "ارتفاع ضغط الدم", "السكر",
    "اضرابات الغدة", "الانيميا", "اخرى", "حمض الفوليك قبل الحمل", "الحديد قبل الحمل", "الكالسيوم قبل الحمل",
    "حمض الفوليك اثناء الحمل", "الحديد اثناء الحمل", "الكالسيوم اثناء الحمل", "وسيلة تنظيم الاسرة المستخدمة سابقا",
    "مدة استخدام الوسيلة السابقة", "شهر الحمل", "تاريخ الزيارة", "التغذية السليمة", "المكملات الغذائية",
    "التمرينات الرياضية", "قسط من النوم والراحة", "المتابعة الدورية للحمل", "التحذير من تناول الادوية",
    "المتاعب البسيطة فى الشهور الاولى", "المتاعب فى الشهور الاخيرة", "علامات الخطر اثناء الحمل",
    "مشاكل الولادة المبكرة", "حركة الجنين ومعرفة الجنس", "تغير لون الجلد حول الحلمة", "ارتداء الملابس الفضفاضة",
    "الاستعداد للولاده", "علامات الولاده", "مميزات الولادة الطبيعية", "الساعة الذهبية الاولى", "ملامسة الجلد للجلد",
    "البداية المبكرة للرضاعه الطبيعية", "الرضاعة الطبيعية المطلقة", "اهمية المباعدة", "وسائل تنظيم الاسره",
    "استخدام وسيلة بعد الولادة مباشرة", "التطور العصبى والنفسى للطفل", "ملاحظات وتوصيات", "تخطيط الزيارة القادمة",
    "المتابعة بعد الولادة"
]

CHILD_COLUMNS = [
    "تاريخ التسجيل", "اسم المستخدم", "اسم الام", "الرقم القومى للام", "رقم موبايل الام", "تاريخ ميلاد الام",
    "مستوى التعليم للام", "عدد الاطفال لدى الام", "المدة بين اخر حملين", "وظيفة الام", "الرقم القومى للاب",
    "رقم موبايل الاب", "اسم الاب", "مستوى التعليم للاب", "اسم الطفل", "تاريخ ميلاد الطفل", "العمر الحالى للطفل",
    "العمر الرحمى للطفل", "مكان المتابعة", "مصدر الاحالة", "نوع الولادة", "مكان الولادة", "وزن الطفل عند الولادة",
    "طول الطفل عند الولادة", "مقاس راس الطفل عند الولادة", "دخول الحضانه", "سبب دخول الحضانه", "مدة البقاء فى الحضانة",
    "ملامسة الجلد فى الساعه الذهبية الاولى", "موعد الزيارة", "تاريخ الزيارة", "رضاعة طبيعية مطلقة",
    "رضاعة طبيعية مع سوائل", "رضاعة طبيعية مع صناعى", "رضاعة لبن صناعى", "الوزن الحالى", "الطول الحالى",
    "محيط الراس الحالى", "فوائد الرضاعه الطبيعية", "كفاية اللبن وكمية البراز", "اعطاء الجرعة اليومية من فيتامين د",
    "كيفية رعاية الاسرة", "البطاقة الصحية واهمية المتابعة", "اهمية الالتزام بالتطعيمات", "التغذيه الصحيحة للام",
    "كيفية التعرف على علامات الخطر", "النمو والتطور الحركى", "التطور الادراكى والمعرفى", "التطور اللغوى",
    "رسائل التربية الايجابية", "الانشطة التحفيزية", "التوعية عن التغذية التكميلية وسلامة الغذاء", "اعطاء جرعة الحديد اليومية",
    "اهمية استخدام الوسيلة", "موقف استخدام الوسيلة", "الحمل الجديد", "الخدمات الغير ملباه", "ملاحظات وتوصيات", "تخطيط الزيارة القادمة"
]

def parse_national_id(nat_id):
    if nat_id and len(nat_id) == 14 and nat_id.isdigit():
        century_code = int(nat_id[0])
        year_digits = int(nat_id[1:3])
        month = int(nat_id[3:5])
        day = int(nat_id[5:7])
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

def get_existing_data(nat_id, sheet_name, id_column):
    if os.path.exists(EXCEL_FILE) and nat_id and len(nat_id) == 14:
        try:
            excel = pd.ExcelFile(EXCEL_FILE)
            for s in excel.sheet_names:
                df = pd.read_excel(excel, sheet_name=s, dtype=str)
                if id_column in df.columns:
                    match = df[df[id_column] == nat_id]
                    if not match.empty:
                        return match.iloc[-1].to_dict()
        except Exception:
            pass
    return {}

if not os.path.exists(EXCEL_FILE):
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        pd.DataFrame(columns=PREGNANT_COLUMNS).to_excel(writer, sheet_name="المشورة الاسرية للحامل", index=False)
        pd.DataFrame(columns=CHILD_COLUMNS).to_excel(writer, sheet_name="سجل المشورة للاطفال", index=False)

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

# ==================== دالة تفريغ الحقول لسجل الأطفال ====================
def clear_child_form():
    for col in CHILD_COLUMNS:
        if f"c_{col}" in st.session_state:
            st.session_state[f"c_{col}"] = ""
    st.session_state["c_تاريخ الزيارة"] = datetime.date.today().strftime("%Y-%m-%d")
    st.session_state["c_الخدمات الغير ملباه"] = "لا يوجد"

# ==================== 1. الصفحة الرئيسية ====================
if menu == "الصفحة الرئيسية":
    st.markdown("<h1>✨ مرحباً بكِ في نظام المشورة الأسرية الشامل ✨</h1>", unsafe_allow_html=True)
    st.write("النظام جاهز تماماً لتسجيل حالات الحوامل والأطفال وحفظها مباشرة في ملف الإكسل والداشبورد.")

# ==================== 2. سجل الحوامل ====================
elif menu == "سجل الحوامل":
    st.markdown("<h2>🤰 سجل المشورة الأسرية للحوامل</h2>", unsafe_allow_html=True)
    
    auto_fill_start = False
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    
    for col in PREGNANT_COLUMNS:
        if f"p_{col}" not in st.session_state:
            if col == "تاريخ الزيارة":
                st.session_state[f"p_{col}"] = today_str
            elif col in ["حمض الفوليك قبل الحمل", "الحديد قبل الحمل", "الكالسيوم قبل الحمل"]:
                st.session_state[f"p_{col}"] = "لا يوجد"
            elif col == "ملاحظات وتوصيات":
                st.session_state[f"p_{col}"] = "تم"
            elif col == "التغذية السليمة":
                auto_fill_start = True
                st.session_state[f"p_{col}"] = "تم"
            elif auto_fill_start:
                if col == "ملاحظات وتوصيات":
                    auto_fill_start = False
                st.session_state[f"p_{col}"] = "تم"
            else:
                st.session_state[f"p_{col}"] = ""

    form_data = {}
    is_in_auto_range = False
    
    for col_name in PREGNANT_COLUMNS:
        if col_name in ["تاريخ التسجيل", "اسم المستخدم"]:
            continue
            
        if col_name == "التغذية السليمة":
            is_in_auto_range = True
            
        if col_name in ["حمض الفوليك قبل الحمل", "الحديد قبل الحمل", "الكالسيوم قبل الحمل"]:
            options = DROPDOWN_OPTIONS[col_name]
            current_val = st.session_state.get(f"p_{col_name}", "لا يوجد")
            idx = options.index(current_val) if current_val in options else 0
            form_data[col_name] = st.selectbox(col_name, options, index=idx, key=f"p_{col_name}")
        elif is_in_auto_range and col_name in DROPDOWN_OPTIONS:
            options = DROPDOWN_OPTIONS[col_name]
            current_val = st.session_state.get(f"p_{col_name}", "تم")
            idx = options.index(current_val) if current_val in options else 0
            form_data[col_name] = st.selectbox(col_name, options, index=idx, key=f"p_{col_name}")
        elif is_in_auto_range:
            form_data[col_name] = st.text_input(col_name, value="تم", key=f"p_{col_name}")
        else:
            if col_name in DROPDOWN_OPTIONS:
                options = DROPDOWN_OPTIONS[col_name]
                current_val = st.session_state.get(f"p_{col_name}", options[0])
                idx = options.index(current_val) if current_val in options else 0
                form_data[col_name] = st.selectbox(col_name, options, index=idx, key=f"p_{col_name}")
            else:
                if col_name == "الرقم القومى":
                    form_data[col_name] = st.text_input(col_name, max_chars=14, key=f"p_{col_name}")
                    if form_data[col_name] and len(form_data[col_name]) == 14:
                        _, calc_age = parse_national_id(form_data[col_name])
                        if calc_age:
                            st.session_state["p_العمر الحالى"] = calc_age
                elif col_name == "العمر الحالى":
                    form_data[col_name] = st.text_input(f"{col_name} [محسوب تلقائياً من الرقم القومي]", key=f"p_{col_name}")
                elif col_name == "تاريخ الزيارة":
                    form_data[col_name] = st.text_input(f"{col_name} [تاريخ اليوم التلقائي]", key=f"p_{col_name}")
                else:
                    form_data[col_name] = st.text_input(col_name, key=f"p_{col_name}")
                    
        if col_name == "ملاحظات وتوصيات":
            is_in_auto_range = False
    
    if st.button("💾 حفظ بيانات الحامل", use_container_width=True):
        form_data["تاريخ التسجيل"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        form_data["اسم المستخدم"] = st.session_state.name
        
        new_df = pd.DataFrame([form_data], dtype=str)
        excel = pd.ExcelFile(EXCEL_FILE)
        all_dfs = {s: pd.read_excel(excel, sheet_name=s, dtype=str) for s in excel.sheet_names}
        
        for col in PREGNANT_COLUMNS:
            if col not in new_df.columns:
                new_df[col] = ""
        new_df = new_df[PREGNANT_COLUMNS]

        if "المشورة الاسرية للحامل" in all_dfs:
            all_dfs["المشورة الاسرية للحامل"] = pd.concat([all_dfs["المشورة الاسرية للحامل"], new_df], ignore_index=True)
        else:
            all_dfs["المشورة الاسرية للحامل"] = new_df

        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            for s, df in all_dfs.items():
                df.to_excel(writer, sheet_name=s, index=False)
        st.success("تم حفظ بيانات الحامل بنجاح! ✨")

# ==================== 3. سجل الأطفال ====================
elif menu == "سجل الأطفال":
    st.markdown("<h2>👶 سجل المشورة الأسرية للأطفال</h2>", unsafe_allow_html=True)
    
    for col in CHILD_COLUMNS:
        if f"c_{col}" not in st.session_state:
            if col == "تاريخ الزيارة" and not st.session_state.get("c_تاريخ الزيارة"):
                st.session_state["c_تاريخ الزيارة"] = datetime.date.today().strftime("%Y-%m-%d")
            elif col == "الخدمات الغير ملباه":
                st.session_state[f"c_{col}"] = "لا يوجد"
            else:
                st.session_state[f"c_{col}"] = ""

    nat_id_mom_input = st.text_input("الرقم القومى للام (اختياري)", max_chars=14, key="c_الرقم القومى للام")
    
    if nat_id_mom_input and len(nat_id_mom_input) == 14:
        b_date_mom, _ = parse_national_id(nat_id_mom_input)
        if b_date_mom and not st.session_state.get("c_تاريخ ميلاد الام"):
            st.session_state["c_تاريخ ميلاد الام"] = b_date_mom

    if nat_id_mom_input and len(nat_id_mom_input) == 14:
        if st.button("🔍 استرجاع بيانات الأسرة المسجلة مسبقاً"):
            found_data = get_existing_data(nat_id_mom_input, "سجل المشورة للاطفال", "الرقم القومى للام") or get_existing_data(nat_id_mom_input, "المشورة الاسرية للحامل", "الرقم القومى")
            for c_name in CHILD_COLUMNS:
                if c_name in ["تاريخ التسجيل", "اسم المستخدم", "الرقم القومى للام"]:
                    continue
                val = found_data.get(c_name, "")
                if val:
                    st.session_state[f"c_{c_name}"] = str(val)
            st.rerun()

    form_data = {}
    
    for col_name in CHILD_COLUMNS:
        if col_name in ["تاريخ التسجيل", "اسم المستخدم"]:
            continue
            
        if col_name == "الرقم القومى للام":
            form_data[col_name] = nat_id_mom_input
            continue

        if col_name in DROPDOWN_OPTIONS:
            options = DROPDOWN_OPTIONS[col_name]
            
            if col_name == "موعد الزيارة":
                auto_visit_choice = VISIT_SCHEDULE_OPTIONS[0]
                try:
                    age_str = st.session_state.get("c_العمر الحالى للطفل", "")
                    if age_str:
                        if "يوم" in age_str or "أسبوع" in age_str:
                            auto_visit_choice = "الاسبوع الاول"
                        else:
                            age_num = float(''.join(filter(lambda x: x.isdigit() or x=='.', age_str)) or 0)
                            if age_num <= 2:
                                auto_visit_choice = "عمر شهرين"
                            elif age_num <= 4:
                                auto_visit_choice = "عمر 4 شهور"
                            elif age_num <= 6:
                                auto_visit_choice = "عمر 6 شهور"
                            elif age_num <= 9:
                                auto_visit_choice = "عمر 9 شهور"
                            elif age_num <= 12:
                                auto_visit_choice = "عمر 12 شهر"
                            elif age_num <= 18:
                                auto_visit_choice = "عمر 18 شهر"
                            elif age_num <= 24:
                                auto_visit_choice = "عمر سنتين"
                            elif age_num <= 30:
                                auto_visit_choice = "عمر سنتين ونصف"
                            elif age_num <= 36:
                                auto_visit_choice = "عمر 3 سنين"
                            elif age_num <= 42:
                                auto_visit_choice = "عمر 3 سنين ونصف"
                            elif age_num <= 48:
                                auto_visit_choice = "عمر 4 سنين"
                            elif age_num <= 54:
                                auto_visit_choice = "عمر 4 سنين ونصف"
                            elif age_num <= 60:
                                auto_visit_choice = "عمر 5 سنين"
                            elif age_num <= 66:
                                auto_visit_choice = "عمر 5 سنين ونصف"
                            else:
                                auto_visit_choice = "عمر 6 سنين"
                except Exception:
                    pass
                
                st.session_state[f"c_{col_name}"] = auto_visit_choice

            if col_name == "الخدمات الغير ملباه" and not st.session_state.get(f"c_{col_name}"):
                st.session_state[f"c_{col_name}"] = "لا يوجد"

            current_val = st.session_state.get(f"c_{col_name}", options[0])
            idx = options.index(current_val) if current_val in options else 0
            
            growth_fields_list = [
                "النمو والتطور الحركى", "التطور الادراكى والمعرفى", "التطور اللغوى"
            ]
            
            if col_name in growth_fields_list:
                options_growth = ["طبيعى", "متقدم", "متاخر"]
                
                try:
                    w_curr_val = float(st.session_state.get("c_الوزن الحالى", 3.0) or 3.0)
                    l_curr_val = float(st.session_state.get("c_الطول الحالى", 50.0) or 50.0)
                    
                    age_raw = st.session_state.get("c_العمر الحالى للطفل", "0")
                    if "يوم" in age_raw or "أسبوع" in age_raw:
                        age_in_months = 1.0
                    else:
                        age_in_months = float(''.join(filter(lambda x: x.isdigit() or x=='.', age_raw)) or 1.0)
                    
                    if age_in_months <= 12:
                        std_weight = 3.3 + (age_in_months * 0.55)
                        std_length = 50.0 + (age_in_months * 2.5)
                    else:
                        std_weight = 10.0 + ((age_in_months - 12) * 0.2)
                        std_length = 75.0 + ((age_in_months - 12) * 0.5)
                    
                    if w_curr_val < (std_weight * 0.85) or l_curr_val < (std_length * 0.90):
                        auto_m = "متاخر"
                    elif w_curr_val > (std_weight * 1.15) or l_curr_val > (std_length * 1.10):
                        auto_m = "متقدم"
                    else:
                        auto_m = "طبيعى"
                except Exception:
                    auto_m = "طبيعى"

                current_val = st.session_state.get(f"c_{col_name}", auto_m)
                m_idx = options_growth.index(current_val) if current_val in options_growth else options_growth.index(auto_m)
                
                form_data[col_name] = st.selectbox(
                    f"{col_name} [مقارنة بالمعدل الطبيعي للعمر - قابل للتعديل الفوري]", 
                    options_growth, 
                    index=m_idx, 
                    key=f"c_{col_name}"
                )
            else:
                # الحقول الخاصة برسائل التربية الإيجابية وحتى التوعية عن التغذية التكميلية (تم / لم يتم) وباقي الحقول
                if col_name in ["رسائل التربية الايجابية", "الانشطة التحفيزية", "التوعية عن التغذية التكميلية وسلامة الغذاء"]:
                    options_status = ["تم", "لم يتم"]
                    current_status = st.session_state.get(f"c_{col_name}", "تم")
                    s_idx = options_status.index(current_status) if current_status in options_status else 0
                    form_data[col_name] = st.selectbox(col_name, options_status, index=s_idx, key=f"c_{col_name}")
                else:
                    form_data[col_name] = st.selectbox(col_name, options, index=idx, key=f"c_{col_name}")
                
        else:
            if col_name == "تاريخ ميلاد الام":
                form_data[col_name] = st.text_input(f"{col_name} [يتولد تلقائياً إذا أُدخل الرقم القومي للأم]", key=f"c_{col_name}")
                
            elif col_name == "تاريخ ميلاد الطفل":
                default_date_val = datetime.date.today()
                existing_b_date = st.session_state.get(f"c_{col_name}", "")
                if existing_b_date:
                    try:
                        default_date_val = datetime.datetime.strptime(existing_b_date.strip(), "%Y-%m-%d").date()
                    except Exception:
                        pass
                
                chosen_date = st.date_input(col_name, value=default_date_val, key=f"c_date_input_{col_name}")
                form_data[col_name] = str(chosen_date)
                st.session_state[f"c_{col_name}"] = form_data[col_name]
                
                try:
                    if form_data[col_name]:
                        today_date = datetime.date.today()
                        delta_days = (today_date - chosen_date).days
                        
                        if delta_days >= 0:
                            if delta_days < 7:
                                age_display = f"{delta_days} يوم"
                            elif delta_days < 30:
                                weeks_count = round(delta_days / 7)
                                age_display = f"{weeks_count} أسبوع"
                            else:
                                months_count = round(delta_days / 30.44, 1)
                                if months_count.is_integer():
                                    months_count = int(months_count)
                                age_display = f"{months_count} شهر"
                            st.session_state["c_العمر الحالى للطفل"] = age_display
                            
                            gestational_weeks_calc = max(24, min(42, 40 - max(0, round((280 - delta_days) / 7))))
                            st.session_state["c_العمر الرحمى للطفل"] = f"{gestational_weeks_calc} أسبوع"
                except Exception:
                    pass

            elif col_name == "العمر الحالى للطفل":
                form_data[col_name] = st.text_input(f"{col_name} [محسوب تلقائياً]", key=f"c_{col_name}")
                
            elif col_name == "العمر الرحمى للطفل":
                form_data[col_name] = st.text_input(f"{col_name} [محسوب بدقة بناءً على تاريخ الميلاد]", key=f"c_{col_name}")

            elif col_name == "وزن الطفل عند الولادة":
                form_data[col_name] = st.text_input(col_name, key=f"c_{col_name}")
                
            elif col_name == "طول الطفل عند الولادة":
                form_data[col_name] = st.text_input(col_name, key=f"c_{col_name}")
                try:
                    w_val = st.session_state.get("c_وزن الطفل عند الولادة", "3.0")
                    if w_val and form_data[col_name]:
                        st.session_state["c_مقاس راس الطفل عند الولادة"] = str(round((float(form_data[col_name]) / 2) + (float(w_val) * 0.5) + 10, 1))
                except ValueError:
                    pass

            elif col_name == "مقاس راس الطفل عند الولادة":
                form_data[col_name] = st.text_input(f"{col_name} [محسوب تلقائياً]", key=f"c_{col_name}")

            elif col_name == "تاريخ الزيارة":
                if not st.session_state.get(f"c_{col_name}"):
                    st.session_state[f"c_{col_name}"] = datetime.date.today().strftime("%Y-%m-%d")
                form_data[col_name] = st.text_input(f"{col_name} [تاريخ اليوم - قابل للتعديل]", key=f"c_{col_name}")

            elif col_name == "تخطيط الزيارة القادمة":
                calc_next_visit_date = ""
                try:
                    current_visit_sel = st.session_state.get("c_موعد الزيارة", VISIT_SCHEDULE_OPTIONS[0])
                    if current_visit_sel in VISIT_SCHEDULE_OPTIONS:
                        curr_idx = VISIT_SCHEDULE_OPTIONS.index(current_visit_sel)
                        next_idx = min(curr_idx + 1, len(VISIT_SCHEDULE_OPTIONS) - 1)
                        next_visit_label = VISIT_SCHEDULE_OPTIONS[next_idx]
                        
                        schedule_days_map = {
                            "الاسبوع الاول": 7, "عمر شهرين": 60, "عمر 4 شهور": 120, "عمر 6 شهور": 180,
                            "عمر 9 شهور": 270, "عمر 12 شهر": 365, "عمر 18 شهر": 545, "عمر سنتين": 730,
                            "عمر سنتين ونصف": 912, "عمر 3 سنين": 1095, "عمر 3 سنين ونصف": 1277, "عمر 4 سنين": 1460,
                            "عمر 4 سنين ونصف": 1642, "عمر 5 سنين": 1825, "عمر 5 سنين ونصف": 2007, "عمر 6 سنين": 2190
                        }
                        
                        b_date_str = st.session_state.get("c_تاريخ ميلاد الطفل", "")
                        if b_date_str:
                            b_date_obj = datetime.datetime.strptime(b_date_str.strip(), "%Y-%m-%d").date()
                            target_days = schedule_days_map.get(next_visit_label, 30)
                            calc_next_visit_date = str(b_date_obj + datetime.timedelta(days=target_days))
                        else:
                            curr_v_date = datetime.datetime.strptime(st.session_state.get("c_تاريخ الزيارة", datetime.date.today().strftime("%Y-%m-%d")), "%Y-%m-%d").date()
                            calc_next_visit_date = str(curr_v_date + datetime.timedelta(days=30))
                except Exception:
                    pass
                
                st.session_state[f"c_{col_name}"] = calc_next_visit_date
                form_data[col_name] = st.text_input(f"{col_name} [تاريخ الزيارة التالية بناءً على موعد الزيارة القادم]", key=f"c_{col_name}")

            elif col_name == "الطول الحالى":
                form_data[col_name] = st.text_input(col_name, key=f"c_{col_name}")
                try:
                    age_s = st.session_state.get("c_العمر الحالى للطفل", "0")
                    w_curr = st.session_state.get("c_الوزن الحالى", "3.0")
                    if age_s and w_curr and form_data[col_name]:
                        if "يوم" in age_s or "أسبوع" in age_s:
                            age_m = 0.5
                        else:
                            age_m = float(''.join(filter(lambda x: x.isdigit() or x=='.', age_s)) or 0)
                        
                        weight_kg = float(''.join(filter(lambda x: x.isdigit() or x=='.', w_curr)) or 0)
                        length_cm = float(''.join(filter(lambda x: x.isdigit() or x=='.', form_data[col_name])) or 0)
                        if age_m <= 3:
                            base_h = 35 + (age_m * 1.8)
                        elif age_m <= 12:
                            base_h = 40.4 + ((age_m - 3) * 0.75)
                        else:
                            base_h = 47 + ((age_m - 12) * 0.25)
                        st.session_state["c_محيط الراس الحالى"] = str(round(base_h + (weight_kg * 0.15) + (length_cm * 0.05) - 2.5, 1))
                except ValueError:
                    pass

            elif col_name == "محيط الراس الحالى":
                form_data[col_name] = st.text_input(f"{col_name} [محسوب بالمعايير العالمية]", key=f"c_{col_name}")

            else:
                form_data[col_name] = st.text_input(col_name, key=f"c_{col_name}")

    if st.button("💾 حفظ بيانات الطفل", use_container_width=True):
        form_data["تاريخ التسجيل"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        form_data["اسم المستخدم"] = st.session_state.name
        form_data["الرقم القومى للام"] = nat_id_mom_input
        
        new_df = pd.DataFrame([form_data], dtype=str)
        
        for col in CHILD_COLUMNS:
            if col not in new_df.columns:
                new_df[col] = ""
        new_df = new_df[CHILD_COLUMNS]

        excel = pd.ExcelFile(EXCEL_FILE)
        all_dfs = {s: pd.read_excel(excel, sheet_name=s, dtype=str) for s in excel.sheet_names}
        
        if "سجل المشورة للاطفال" in all_dfs:
            all_dfs["سجل المشورة للاطفال"] = pd.concat([all_dfs["سجل المشورة للاطفال"], new_df], ignore_index=True)
        else:
            all_dfs["سجل المشورة للاطفال"] = new_df

        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            for s, df in all_dfs.items():
                df.to_excel(writer, sheet_name=s, index=False)
        
        # تأثير تطاير القلوب الكثيرة والقلب الكبير باسم د. شيماء
        hearts_html = ""
        import random
        for _ in range(35):
            left_pos = random.randint(0, 95)
            anim_dur = random.uniform(1.5, 3.0)
            delay = random.uniform(0, 0.5)
            size = random.randint(18, 38)
            hearts_html += f'<div class="f-heart" style="left: {left_pos}vw; font-size: {size}px; animation-duration: {anim_dur}s; animation-delay: {delay}s;">❤️</div>'

        st.markdown(f"""
            <div class="floating-hearts-overlay">
                {hearts_html}
            </div>
            <div class="big-heart-container">
                <div class="big-heart">
                    <div class="heart-text">د. شيماء</div>
                </div>
            </div>
        """, unsafe_allow_html=True)
        
        st.success("تم حفظ بيانات الطفل بنجاح! جاري تفريغ الحقول لتسجيل طفل جديد... ✨")
        
        import time
        time.sleep(3)
        clear_child_form()
        st.rerun()

# ==================== 4. استعراض البيانات والداشبورد ====================
elif menu == "استعراض البيانات والداشبورد":
    st.markdown("<h2>📊 لوحة المؤشرات (الداشبورد) وتصفية الحالات</h2>", unsafe_allow_html=True)
    
    if os.path.exists(EXCEL_FILE):
        excel = pd.ExcelFile(EXCEL_FILE)
        
        df_preg = pd.read_excel(excel, sheet_name="المشورة الاسرية للحامل", dtype=str) if "المشورة الاسرية للحامل" in excel.sheet_names else pd.DataFrame(columns=PREGNANT_COLUMNS)
        df_child = pd.read_excel(excel, sheet_name="سجل المشورة للاطفال", dtype=str) if "سجل المشورة للاطفال" in excel.sheet_names else pd.DataFrame(columns=CHILD_COLUMNS)
        
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.metric(label="🤰 إجمالي سجلات الحوامل المسجلة", value=len(df_preg))
        with col_m2:
            st.metric(label="👶 إجمالي سجلات الأطفال المسجلة", value=len(df_child))
            
        st.markdown("---")
        tab1, tab2 = st.tabs(["🤰 جدول الحوامل", "👶 جدول الأطفال"])
        
        with tab1:
            st.subheader("بيانات الحوامل المسجلة")
            if not df_preg.empty:
                search_p = st.text_input("بحث بالاسم أو أي بيانات للحامل", key="search_preg")
                if search_p:
                    filtered_preg = df_preg[df_preg.astype(str).apply(lambda x: x.str.contains(search_p, case=False)).any(axis=1)]
                else:
                    filtered_preg = df_preg
                
                if st.session_state.role == "admin":
                    st.markdown("### ⚙️ لوحة تحكم الآدمن: حذف سجل للحامل برقم الصف")
                    row_indices_p = list(filtered_preg.index)
                    row_to_delete_p = st.selectbox("اختر رقم الصف (Index) المراد حذفه نهائياً", options=[-1] + row_indices_p, format_func=lambda x: "اختر رقم الصف..." if x == -1 else f"صف رقم: {x} (الاسم: {filtered_preg.loc[x, 'الاسم'] if 'الاسم' in filtered_preg.columns and pd.notna(filtered_preg.loc[x, 'الاسم']) else 'غير متوفر'})", key="del_preg_index_select")
                    
                    if st.button("🗑️ إزالة السجل المحدد برقم الصف من شيت الحوامل", key="btn_del_preg_idx"):
                        if row_to_delete_p != -1:
                            updated_df_preg = df_preg.drop(index=row_to_delete_p).reset_index(drop=True)
                            all_dfs = {s: pd.read_excel(excel, sheet_name=s, dtype=str) for s in excel.sheet_names}
                            all_dfs["المشورة الاسرية للحامل"] = updated_df_preg
                            with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
                                for s, df in all_dfs.items():
                                    df.to_excel(writer, sheet_name=s, index=False)
                            st.success("تم حذف السجل بنجاح! جاري التحديث...")
                            st.rerun()
                        else:
                            st.warning("الرجاء اختيار رقم صف صحيح للحذف.")
                    st.markdown("---")

                display_df_preg = filtered_preg.copy()
                display_df_preg.insert(0, "رقم الصف", display_df_preg.index)
                st.dataframe(display_df_preg, use_container_width=True)
            else:
                st.info("لا توجد بيانات مسجلة للحوامل حتى الآن.")
                
        with tab2:
            st.subheader("بيانات الأطفال المسجلة")
            if not df_child.empty:
                search_c = st.text_input("بحث باسم الطفل أو الأم أو أي بيانات", key="search_child")
                if search_c:
                    filtered_child = df_child[df_child.astype(str).apply(lambda x: x.str.contains(search_c, case=False)).any(axis=1)]
                else:
                    filtered_child = df_child
                
                if st.session_state.role == "admin":
                    st.markdown("### ⚙️ لوحة تحكم الآدمن: حذف سجل لطفل برقم الصف")
                    row_indices_c = list(filtered_child.index)
                    row_to_delete_c = st.selectbox("اختر رقم الصف (Index) المراد حذفه نهائياً", options=[-1] + row_indices_c, format_func=lambda x: "اختر رقم الصف..." if x == -1 else f"صف رقم: {x} (طفل: {filtered_child.loc[x, 'اسم الطفل'] if 'اسم الطفل' in filtered_child.columns and pd.notna(filtered_child.loc[x, 'اسم الطفل']) else 'غير متوفر'} - أم: {filtered_child.loc[x, 'اسم الام'] if 'اسم الام' in filtered_child.columns and pd.notna(filtered_child.loc[x, 'اسم الام']) else 'غير متوفر'})", key="del_child_index_select")
                    
                    if st.button("🗑️ إزالة السجل المحدد برقم الصف من شيت الأطفال", key="btn_del_child_idx"):
                        if row_to_delete_c != -1:
                            updated_df_child = df_child.drop(index=row_to_delete_c).reset_index(drop=True)
                            all_dfs = {s: pd.read_excel(excel, sheet_name=s, dtype=str) for s in excel.sheet_names}
                            all_dfs["سجل المشورة للاطفال"] = updated_df_child
                            with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
                                for s, df in all_dfs.items():
                                    df.to_excel(writer, sheet_name=s, index=False)
                            st.success("تم حذف السجل بنجاح! جاري التحديث...")
                            st.rerun()
                        else:
                            st.warning("الرجاء اختيار رقم صف صحيح للحذف.")
                    st.markdown("---")

                display_df_child = filtered_child.copy()
                display_df_child.insert(0, "رقم الصف", display_df_child.index)
                st.dataframe(display_df_child, use_container_width=True)
            else:
                st.info("لا توجد بيانات مسجلة للأطفال حتى الآن.")
    else:
        st.warning("لم يتم العثور على ملف البيانات.")

# ==================== 5. إدارة المستخدمين (للآدمن فقط) ====================
elif menu == "إدارة المستخدمين":
    if st.session_state.role == "admin":
        st.markdown("<h2>⚙️ لوحة إدارة المستخدمين والصلاحيات</h2>", unsafe_allow_html=True)
        st.write("يمكنك الاطلاع على حسابات الطبيبات المسجلات في النظام:")
        
        user_df = pd.DataFrame([
            {"اسم المستخدم": k, "الاسم بالكامل": v["name"], "الدور": v["role"]}
            for k, v in DEFAULT_USERS.items()
        ])
        st.dataframe(user_df, use_container_width=True)
    else:
        st.error("عذراً، هذه الصفحة مخصصة للمدير (Admin) فقط.")
