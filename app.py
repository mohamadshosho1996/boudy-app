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
    .main .block-container {
        padding-top: 1.0rem;
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
""", unsafe_allow_html=True)

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
    "مستوى التعليم للام": ["امى", "يجيد القراءة", "مؤهل متوسط", "فوق متوسط", "مؤهل عالى"],
    "مستوى التعليم للاب": ["امى", "يجيد القراءة", "مؤهل متوسط", "فوق متوسط", "مؤهل عالى"],
    "الوظيفة": ["يعمل", "لا تعمل"],
    "الوظيفة للام": ["يعمل", "لا تعمل"],
    "قرابة بين الزوجين": ["لا يوجد", "من الدرجة الأولي", "من الدرجة الثانية", "من الدرجة الخامسة"],
    "نوع الولادة": ["طبيعى", "قيصرى"],
    "مكان الولادة": ["المستشفى", "المنزل"],
    "وسيلة تنظيم الأسرة المستخدمة سابقا": ["لا يوجد", "اقراص", "حقن", "كبسولات", "لولب", "طرق طبيعية"],
    "وسائل تنظيم الأسرة": ["لا يوجد", "اقراص", "حقن", "كبسولات", "لولب", "طرق طبيعية"],
    "سبب دخول الحضانة": [
        "انخفاض وزن الطفل.", "احتياج الطفل لأدوية محددة بهذا الوقت.", "صعوبة شديدة في التنفس لعدم اكتمال نمو الرئتين.",
        "ارتفاع درجة حرارة جسم الرضيع.", "تعطل العمليات الحيوية بجسم الطفل.", "انخفاض معدل الجلوكوز في دم الطفل.",
        "معاناة الرضيع مشكلات في الجهاز الهضمي.", "إصابة الطفل بعدوى في الدم.", "إصابة الطفل بالصفراء.",
        "حدوث مشكلات خلال الولادة “الولادة المتعسرة أو الحمل الحرج”.", "وجود عيب خلقي يمنع الطفل عن التنفس أو الرضاعة بشكل طبيعي."
    ],
    "مكان المتابعة": ["وحدة", "مستشفى", "أخرى"],
    "مصدر الاحاله": ["مستشفى الولادة", "عيادة خاصة", "عيادة التطعيمات", "نصيحة"],
    "موعد الزيارة": VISIT_SCHEDULE_OPTIONS,
    "رضاعة طبيعية مطلقة": ["3 شهور", "4 شهور", "6 شهور"],
    "رضاعة طبيعية مع سوائل وأعشاب": ["تم", "لم يتم"],
    "رضاعة طبيعية مع صناعي": ["تم", "لم يتم"],
    "رضاعة لبن صناعي": ["تم", "لم يتم"],
    "موقف إستخدام وسيلة تنظيم أسرة": ["توجد", "لا يوجد", "مرغوب", "غير مرغوب", "حدث", "لم يحدث"],
    "الحمل الجديد": ["مرغوب", "غير مرغوب"],
    "دخول الحضانة": ["تم", "لم يتم"],
    "ملامسة الجلد فى الساعة الذهبية الأولى": ["تم", "لم يتم"],
    "الرضاعة الطبيعية فى الساعة الذهبية الأولى": ["تم", "لم يتم"],
    "أمراض مزمنة: إرتفاع ضغط الدم": ["يوجد", "لا يوجد"],
    "أمراض مزمنة: السكر": ["يوجد", "لا يوجد"],
    "أمراض مزمنة: إضطرابات الغدة": ["يوجد", "لا يوجد"],
    "أمراض مزمنة: الأنيميا": ["يوجد", "لا يوجد"],
    'مكملات "قبل": حمض الفوليك': ["يوجد", "لا يوجد"],
    'مكملات "قبل": الحديد': ["يوجد", "لا يوجد"],
    'مكملات "قبل": الكالسيوم': ["يوجد", "لا يوجد"],
    'مكملات "أثناء": حمض الفوليك': ["يوجد", "لا يوجد"],
    'مكملات "أثناء": الحديد': ["يوجد", "لا يوجد"],
    'مكملات "أثناء": الكالسيوم': ["يوجد", "لا يوجد"],
    "النمو والتطور الحركي": ["طبيعي", "متقدم", "متأخر", "يحتاج متابعة", "تم التوعية"],
    "التطور الإدراكي والمعرفي": ["طبيعي", "متأخر", "يحتاج متابعة", "تم التوعية"],
    "التطور اللغوي": ["طبيعي", "متأخر", "يحتاج متابعة", "تم التوعية"],
    "رسائل التربية الإيجابية": ["طبيعي", "متأخر", "يحتاج متابعة", "تم التوعية"],
    "الأنشطة التحفيزية": ["طبيعي", "متأخر", "يحتاج متابعة", "تم التوعية"],
    "التوعية عن التغذية التكميلية وسلامة الغذاء والتغذية السليمة": ["طبيعي", "متأخر", "يحتاج متابعة", "تم التوعية"],
    "إعطاء الجرعة اليومية من الحديد": ["يوجد", "لا يوجد"],
    "أهمية إستخدام وسيلة تنظيم أسرة وأهمية المباعدة": ["تم التوعية", "غير مهتمة", "مرفوض"],
    "الخدمات الغير ملباه": ["يوجد", "لا يوجد"]
}

PREGNANT_COLUMNS = [
    "تاريخ التسجيل", "اسم المستخدم", "الاسم", "العنوان", "الرقم القومى", "رقم الموبايل", "العمر الحالى", "السن عند الزواج", "السن عند الحمل الاول",
    "مستوى التعليم", "الوظيفة", "تاريخ اخر دورة شهرية", "قرابة بين الزوجين", "عدد مرات الحمل", "عدد مرات الاجهاض",
    "عدد الاطفال", "المدة بين اخر حملين", "نوع الولادة", "أمراض مزمنة: إرتفاع ضغط الدم", "أمراض مزمنة: السكر",
    "أمراض مزمنة: إضطرابات الغدة", "أمراض مزمنة: الأنيميا", "أمراض مزمنة: اخرى", 'مكملات "قبل": حمض الفوليك',
    'مكملات "قبل": الحديد', 'مكملات "قبل": الكالسيوم', 'مكملات "أثناء": حمض الفوليك', 'مكملات "أثناء": الحديد',
    'مكملات "أثناء": الكالسيوم', "وسيلة تنظيم الأسرة المستخدمة سابقا", "مدة إستخدام الوسيلة السابقة", "شهر الحمل",
    "التاريخ الزيارة", "التغذية السليمة", "المكملات الغذائية", "التمرينات الرياضية", "قسط من النوم والراحة",
    "المتابعة الدورية للحمل", "التحذير من تناول الأدوية بدون إستشارة طبيب والتعرض للتدخين والأبخرة",
    "المتاعب البسيطة في الشهور الأولى", "المتاعب في الشهور الأخيرة", "علامات الخطر أثناء الحمل",
    "مشاكل الولادة المبكرة وكيفية تجنبها", "حركة الجنين / معرفة جنس الجنين/ تمييز الأصوات من قبل الجنين",
    "تغير لون الجلد حول الحلمة وظهور بعض إفرازات من الثدي", "إرتداء الملابس الفضفاضة المريحة",
    "الإستعداد للولادة / تحضير ملابس المولود ، الخ", "علامات الولادة", "مميزات الولادة الطبيعية",
    "الساعة الذهبية الأولى", "ملامسة الجلد للجلد", "البداية المبكرة للرضاعة الطبيعية", "الرضاعة الطبيعية المطلقة",
    "أهمية المباعدة", "وسائل تنظيم الأسرة", "إستخدام وسيلة بعد الولادة مباشرة", "التطور العصبي والنفسي للطفل",
    "ملاحظات/ توصيات", "تخطيط الزيارة القادمة", "المتابعة ما بعد الولادة"
]

CHILD_COLUMNS = [
    "تاريخ التسجيل", "اسم المستخدم", "تاريخ اول زيارة", "رقم الحالة", "اسم الام", "الرقم القومى للام", "رقم الموبايل للام", "تاريخ ميلاد الام",
    "مستوى التعليم للام", "عدد الاطفال لدى الام", "المدة بين اخر حملين", "الوظيفة للام", "الرقم القومى للاب",
    "رقم الموبايل للاب", "اسم الاب", "مستوى التعليم للاب", "اسم الطفل", "تاريخ الميلاد للطفل", "العمر الحالى للطفل (شهور)",
    "العمر الرحمى للطفل (أسابيع)", "مكان المتابعة", "وحدة", "مستشفى", "أخرى", "مصدر الاحاله", "مستشفى الولادة", "عيادة خاصة", "عيادة التطعيمات", "نصيحة", "نوع الولادة", "مكان الولادة", "وزن الطفل عند الولادة",
    "طول الطفل عند الولادة", "مقاس راس الطفل عند الولادة", "دخول الحضانة", "سبب دخول الحضانة", "مدة البقاء فى الحضانة",
    "ملامسة الجلد فى الساعة الذهبية الأولى", "الرضاعة الطبيعية فى الساعة الذهبية الأولى", "موعد الزيارة",
    "تاريخ الزيارة", "رضاعة طبيعية مطلقة", "رضاعة طبيعية مع سوائل وأعشاب", "رضاعة طبيعية مع صناعي",
    "رضاعة لبن صناعي", "الوزن (كجم)", "الطول (سم)", "محيط الرأس (سم)", "فوائد الرضاعة الطبيعية والأوضاع وعلامات الجوع والشبع",
    "كفاية اللبن وكمية البراز", "إعطاء الجرعة اليومية من فيتامين د", "كيفية رعاية السرة والإهتمام بنظافة الطفل",
    "البطاقة الصحية وأهمية المتابعة الدورية ومنحنيات النمو", "أهمية الإلتزام بتطعيمات الطفل", "التغذية الصحية للأم المرضعة",
    "كيفية التعرف على علامات الخطورة", "النمو والتطور الحركي", "التطور الإدراكي والمعرفي",
    "التطور اللغوي", "رسائل التربية الإيجابية", "الأنشطة التحفيزية", "التوعية عن التغذية التكميلية وسلامة الغذاء والتغذية السليمة",
    "إعطاء الجرعة اليومية من الحديد", "أهمية إستخدام وسيلة تنظيم أسرة وأهمية المباعدة", "موقف إستخدام وسيلة تنظيم أسرة",
    "الحمل الجديد", "الخدمات الغير ملباه", "ملاحظات/ توصيات", "تخطيط الزيارة القادمة"
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

# ==================== 1. الصفحة الرئيسية ====================
if menu == "الصفحة الرئيسية":
    st.markdown("<h1>✨ مرحباً بكِ في نظام المشورة الأسرية الشامل ✨</h1>", unsafe_allow_html=True)
    st.write("النظام مستقر وسريع جداً. يمكنكِ الكتابة في الحقول أو استخدام زر الميكروفون الموجود في لوحة المفاتيح لديكِ مباشرة.")

# ==================== 2. سجل الحوامل ====================
elif menu == "سجل الحوامل":
    st.markdown("<h2>🤰 سجل المشورة الأسرية للحوامل</h2>", unsafe_allow_html=True)
    nat_id = st.text_input("الرقم القومى", max_chars=14, key="pregnant_nat_id_input")
    
    if nat_id and len(nat_id) == 14:
        b_date, c_age = parse_national_id(nat_id)
        st.session_state["p_العمر الحالى"] = c_age

    form_data = {}
    form_data["الرقم القومى"] = nat_id
    for col_name in PREGNANT_COLUMNS:
        if col_name in ["تاريخ التسجيل", "اسم المستخدم", "الرقم القومى"]:
            continue
        if f"p_{col_name}" not in st.session_state:
            st.session_state[f"p_{col_name}"] = ""
            
        if col_name in DROPDOWN_OPTIONS:
            options = DROPDOWN_OPTIONS[col_name]
            current_val = st.session_state.get(f"p_{col_name}", options[0])
            idx = options.index(current_val) if current_val in options else 0
            form_data[col_name] = st.selectbox(col_name, options, index=idx, key=f"p_{col_name}")
        else:
            form_data[col_name] = st.text_input(col_name, key=f"p_{col_name}")
    
    if st.button("💾 حفظ بيانات الحامل", use_container_width=True):
        if not nat_id or len(nat_id) != 14:
            st.error("برجاء إدخال رقم قومي صحيح مكون من 14 رقماً!")
        else:
            form_data["تاريخ التسجيل"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            form_data["اسم المستخدم"] = st.session_state.name
            new_df = pd.DataFrame([form_data], dtype=str)
            excel = pd.ExcelFile(EXCEL_FILE)
            all_dfs = {s: pd.read_excel(excel, sheet_name=s, dtype=str) for s in excel.sheet_names}
            all_dfs["المشورة الاسرية للحامل"] = pd.concat([all_dfs["المشورة الاسرية للحامل"], new_df], ignore_index=True) if "المشورة الاسرية للحامل" in all_dfs else new_df
            with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
                for s, df in all_dfs.items():
                    df.to_excel(writer, sheet_name=s, index=False)
            st.success("تم حفظ بيانات الحامل بنجاح! ✨")

# ==================== 3. سجل الأطفال ====================
elif menu == "سجل الأطفال":
    st.markdown("<h2>👶 سجل المشورة الأسرية للأطفال</h2>", unsafe_allow_html=True)
    
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        name_mom_input = st.text_input("اسم الام", key="child_name_mom_input")
    with col_m2:
        nat_id_mom = st.text_input("الرقم القومى للام", max_chars=14, key="child_nat_id_mom_input")
    
    mom_birth_date_str = ""
    if nat_id_mom and len(nat_id_mom) == 14:
        mom_birth_date_str, _ = parse_national_id(nat_id_mom)
        st.session_state["c_تاريخ ميلاد الام"] = mom_birth_date_str

    if nat_id_mom and len(nat_id_mom) == 14:
        if st.button("🔍 استرجاع كافة بيانات الأب والأم والأسرة المسجلة"):
            found_data = get_existing_data(nat_id_mom, "سجل المشورة للاطفال", "الرقم القومى للام") or get_existing_data(nat_id_mom, "المشورة الاسرية للحامل", "الرقم القومى")
            for col_name in CHILD_COLUMNS:
                if col_name in ["تاريخ التسجيل", "اسم المستخدم", "الرقم القومى للام"]:
                    continue
                val = found_data.get(col_name, "")
                if col_name == "اسم الام" and not val: val = name_mom_input
                if col_name == "تاريخ ميلاد الام" and not val: val = mom_birth_date_str
                st.session_state[f"c_{col_name}"] = str(val)
            st.rerun()

    form_data = {}
    form_data["الرقم القومى للام"] = nat_id_mom
    form_data["اسم الام"] = name_mom_input

    for col_name in CHILD_COLUMNS:
        if col_name in ["تاريخ التسجيل", "اسم المستخدم", "الرقم القومى للام", "اسم الام", "تاريخ ميلاد الام", "العمر الحالى للطفل (شهور)", "العمر الرحمى للطفل (أسابيع)", "وحدة", "مستشفى", "أخرى", "مستشفى الولادة", "عيادة خاصة", "عيادة التطعيمات", "نصيحة", "مقاس راس الطفل عند الولادة", "محيط الرأس (سم)", "النمو والتطور الحركي"]:
            continue
            
        if f"c_{col_name}" not in st.session_state:
            st.session_state[f"c_{col_name}"] = ""
            
        if col_name in DROPDOWN_OPTIONS:
            options = DROPDOWN_OPTIONS[col_name]
            current_val = st.session_state.get(f"c_{col_name}", options[0])
            idx = options.index(current_val) if current_val in options else 0
            form_data[col_name] = st.selectbox(col_name, options, index=idx, key=f"c_{col_name}")
        else:
            if col_name == "تاريخ الميلاد للطفل":
                form_data[col_name] = st.text_input(col_name, key=f"c_{col_name}")
                
                calc_child_months = ""
                calc_gestational_weeks = ""
                try:
                    if form_data[col_name]:
                        b_date_obj = datetime.datetime.strptime(form_data[col_name].strip(), "%Y-%m-%d").date()
                        today_date = datetime.date.today()
                        diff_days = (today_date - b_date_obj).days
                        if diff_days >= 0:
                            calc_child_months = str(round(diff_days / 30.44, 1))
                            calc_gestational_weeks = str(max(37, round(40 - (diff_days / 7))))
                except Exception:
                    pass
                
                st.session_state["c_العمر الحالى للطفل (شهور)"] = calc_child_months
                st.session_state["c_العمر الرحمى للطفل (أسابيع)"] = calc_gestational_weeks

            elif col_name == "طول الطفل عند الولادة":
                form_data[col_name] = st.text_input(col_name, key=f"c_{col_name}")
                
                birth_w_val = st.session_state.get("c_وزن الطفل عند الولادة", "3.0")
                calc_b_head = ""
                try:
                    if birth_w_val and form_data[col_name]:
                        calc_b_head = str(round((float(form_data[col_name]) / 2) + (float(birth_w_val) * 0.5) + 10, 1))
                except ValueError:
                    pass
                st.session_state["c_مقاس راس الطفل عند الولادة"] = calc_b_head
                form_data["مقاس راس الطفل عند الولادة"] = st.text_input("مقاس راس الطفل عند الولادة (سم) [محسوب تلقائياً]", key="c_مقاس راس الطفل عند الولادة")

            elif col_name == "الطول (سم)":
                form_data[col_name] = st.text_input(col_name, key=f"c_{col_name}")
                
                curr_age_v = st.session_state.get("c_العمر الحالى للطفل (شهور)", "0")
                curr_w_v = st.session_state.get("c_الوزن (كجم)", "3.0")
                curr_len_v = form_data[col_name]
                calc_c_head = ""
                try:
                    if curr_age_v and curr_w_v and curr_len_v:
                        age_m = float(curr_age_v)
                        weight_kg = float(curr_w_v)
                        length_cm = float(curr_len_v)
                        if age_m <= 3:
                            base_head = 35 + (age_m * 1.8)
                        elif age_m <= 12:
                            base_head = 40.4 + ((age_m - 3) * 0.75)
                        else:
                            base_head = 47 + ((age_m - 12) * 0.25)
                        
                        adjustment = (weight_kg * 0.15) + (length_cm * 0.05)
                        calc_c_head = str(round(base_head + adjustment - 2.5, 1))
                except ValueError:
                    pass
                
                st.session_state["c_محيط الرأس (سم)"] = calc_c_head
                form_data["محيط الرأس (سم)"] = st.text_input("محيط الرأس الحالي (سم) [محسوب بالمعايير العالمية]", key="c_محيط الرأس (سم)")

            elif col_name == "النمو والتطور الحركي":
                curr_age_v = st.session_state.get("c_العمر الحالى للطفل (شهور)", "0")
                curr_w_v = st.session_state.get("c_الوزن (كجم)", "3.0")
                auto_motor = "طبيعي"
                try:
                    if curr_age_v and curr_w_v:
                        age_f = float(curr_age_v)
                        cw_f = float(curr_w_v)
                        exp_min = 3 + (age_f * 0.5)
                        if cw_f > (exp_min * 1.25):
                            auto_motor = "متقدم"
                        elif cw_f < (exp_min * 0.75):
                            auto_motor = "متأخر"
                        else:
                            auto_motor = "طبيعي"
                except ValueError:
                    pass
                
                m_options = DROPDOWN_OPTIONS["النمو والتطور الحركي"]
                default_m_idx = m_options.index(auto_motor) if auto_motor in m_options else 0
                form_data[col_name] = st.selectbox("النمو والتطور الحركي [محسوب تلقائياً وقابل للتعديل]", m_options, index=default_m_idx, key="c_auto_motor")
            else:
                form_data[col_name] = st.text_input(col_name, key=f"c_{col_name}")

    form_data["تاريخ ميلاد الام"] = st.session_state.get("c_تاريخ ميلاد الام", "")
    form_data["العمر الحالى للطفل (شهور)"] = st.text_input("العمر الحالى للطفل (شهور) [محسوب تلقائياً]", key="c_العمر الحالى للطفل (شهور)")
    form_data["العمر الرحمى للطفل (أسابيع)"] = st.text_input("العمر الرحمى للطفل (أسابيع) [محسوب تلقائياً]", key="c_العمر الرحمى للطفل (أسابيع)")

    if st.button("💾 حفظ بيانات الطفل", use_container_width=True):
        if not nat_id_mom or len(nat_id_mom) != 14:
            st.error("برجاء إدخال الرقم القومي للأم صحيحاً مكوناً من 14 رقماً!")
        else:
            form_data["تاريخ التسجيل"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            form_data["اسم المستخدم"] = st.session_state.name
            
            place_val = form_data.get("مكان المتابعة", "")
            source_val = form_data.get("مصدر الاحاله", "")
            form_data["وحدة"] = "تم" if place_val == "وحدة" else ""
            form_data["مستشفى"] = "تم" if place_val == "مستشفى" else ""
            form_data["أخرى"] = "تم" if place_val == "أخرى" else ""
            form_data["مستشفى الولادة"] = "تم" if source_val == "مستشفى الولادة" else ""
            form_data["عيادة خاصة"] = "تم" if source_val == "عيادة خاصة" else ""
            form_data["عيادة التطعيمات"] = "تم" if source_val == "عيادة التطعيمات" else ""
            form_data["نصيحة"] = "تم" if source_val == "نصيحة" else ""
            
            new_df = pd.DataFrame([form_data], dtype=str)
            excel = pd.ExcelFile(EXCEL_FILE)
            all_dfs = {s: pd.read_excel(excel, sheet_name=s, dtype=str) for s in excel.sheet_names}
            all_dfs["سجل المشورة للاطفال"] = pd.concat([all_dfs["سجل المشورة للاطفال"], new_df], ignore_index=True) if "سجل المشورة للاطفال" in all_dfs else new_df
            with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
                for s, df in all_dfs.items():
                    df.to_excel(writer, sheet_name=s, index=False)
            st.success("تم حفظ بيانات الطفل والحسابات التلقائية في الإكسيل بنجاح! ✨")

# ==================== 4. استعراض البيانات والداشبورد ====================
elif menu == "استعراض البيانات والداشبورد":
    st.markdown("<h2>📊 استعراض السجلات والبيانات المحفوظة وتنزيل الإكسيل</h2>", unsafe_allow_html=True)
    if os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, "rb") as f:
            st.download_button(
                label="📥 تنزيل ملف الإكسيل (template.xlsx) المحدث",
                data=f,
                file_name="template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        st.markdown("---")
        excel = pd.ExcelFile(EXCEL_FILE)
        for s in excel.sheet_names:
            st.subheader(f"📁 شيت: {s}")
            df = pd.read_excel(excel, sheet_name=s, dtype=str)
            st.dataframe(df, use_container_width=True)
    else:
        st.info("لا توجد بيانات مسجلة حتى الآن.")

# ==================== 5. إدارة المستخدمين ====================
elif menu == "إدارة المستخدمين" and st.session_state.role == "admin":
    st.markdown("<h2>⚙️ إدارة حسابات الطبيبات وصلاحيات النظام</h2>", unsafe_allow_html=True)
    for k, v in DEFAULT_USERS.items():
        st.info(f"اسم المستخدم (ID): **{k}** | الاسم: **{v['name']}** | الصلاحية: **{v['role']}**")
