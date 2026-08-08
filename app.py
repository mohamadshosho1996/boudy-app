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
    "النمو والتطور الحركي": ["طبيعي", "متأخر", "يحتاج متابعة", "تم التوعية"],
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
    "العمر الرحمى للطفل (أسابيع)", "وحدة", "مستشفى", "أخرى", "مستشفى الولادة", "عيادة خاصة", "عيادة التطعيمات", "نصيحة", "نوع الولادة", "مكان الولادة", "وزن الطفل عند الولادة",
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

BASIC_PREGNANT_FIELDS = ["الاسم", "العنوان", "رقم الموبايل", "العمر الحالى"]
BASIC_CHILD_FIELDS = ["اسم الام", "رقم الموبايل للام", "تاريخ ميلاد الام", "رقم الموبايل للاب", "اسم الاب"]

# دالة استخراج تاريخ الميلاد والعمر من الرقم القومي المصري
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

# دالة البحث عن البيانات السابقة للرقم القومي في الإكسيل
def get_existing_data(nat_id, sheet_name, id_column):
    if os.path.exists(EXCEL_FILE) and nat_id and len(nat_id) == 14:
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, dtype=str)
            if id_column in df.columns:
                match = df[df[id_column] == nat_id]
                if not match.empty:
                    return match.iloc[-1].to_dict()
        except Exception:
            pass
    return {}

# التأكد التام من إنشاء ملف الإكسيل بالأعمدة الكاملة
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
    st.write("النظام جاهز تماماً لتسجيل الحالات، واستخراج التواريخ والعمر من الرقم القومي، واسترجاع البيانات المسجلة مسبقاً تلقائياً.")

# ==================== 2. سجل الحوامل ====================
elif menu == "سجل الحوامل":
    st.markdown("<h2>🤰 سجل المشورة الأسرية للحوامل</h2>", unsafe_allow_html=True)
    
    nat_id = st.text_input("الرقم القومى", max_chars=14, key="pregnant_nat_id_input")
    
    if nat_id and len(nat_id) == 14:
        if st.button("🔍 استرجاع البيانات المسجلة"):
            old_data = get_existing_data(nat_id, "المشورة الاسرية للحامل", "الرقم القومى")
            _, calc_age = parse_national_id(nat_id)
            
            for col_name in PREGNANT_COLUMNS:
                if col_name in ["تاريخ التسجيل", "اسم المستخدم", "الرقم القومى"]:
                    continue
                val = old_data.get(col_name, "") if col_name in BASIC_PREGNANT_FIELDS else ""
                if col_name == "العمر الحالى" and not val:
                    val = calc_age
                st.session_state[f"p_{col_name}"] = str(val)
            
            if old_data:
                st.success("✨ تم العثور على سجل سابق لهذه الحالة وتم استرجاع البيانات الأساسية للأم تلقائياً!")
            st.rerun()

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
            
            if "المشورة الاسرية للحامل" in all_dfs:
                all_dfs["المشورة الاسرية للحامل"] = pd.concat([all_dfs["المشورة الاسرية للحامل"], new_df], ignore_index=True)
            else:
                all_dfs["المشورة الاسرية للحامل"] = new_df
                
            with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
                for s, df in all_dfs.items():
                    df.to_excel(writer, sheet_name=s, index=False)
            st.success("تم حفظ بيانات الحامل كاملة بنجاح في الإكسيل! ✨")

# ==================== 3. سجل الأطفال ====================
elif menu == "سجل الأطفال":
    st.markdown("<h2>👶 سجل المشورة الأسرية للأطفال</h2>", unsafe_allow_html=True)
    
    nat_id_mom = st.text_input("الرقم القومى للام", max_chars=14, key="child_nat_id_mom_input")
    
    if nat_id_mom and len(nat_id_mom) == 14:
        if st.button("🔍 استرجاع بيانات الأم"):
            found_data = get_existing_data(nat_id_mom, "سجل المشورة للاطفال", "الرقم القومى للام")
            if not found_data:
                found_data = get_existing_data(nat_id_mom, "المشورة الاسرية للحامل", "الرقم القومى")
            
            b_date_mom, _ = parse_national_id(nat_id_mom)
            
            for col_name in CHILD_COLUMNS:
                if col_name in ["تاريخ التسجيل", "اسم المستخدم", "الرقم القومى للام"]:
                    continue
                val = found_data.get(col_name, "") if col_name in BASIC_CHILD_FIELDS else ""
                if col_name == "تاريخ ميلاد الام" and not val:
                    val = b_date_mom
                st.session_state[f"c_{col_name}"] = str(val)
                
            if found_data:
                st.success("✨ تم العثور على سجل سابق للأم وتم استرجاع بياناتها الأساسية تلقائياً!")
            st.rerun()

    form_data = {}
    form_data["الرقم القومى للام"] = nat_id_mom
    
    for col_name in CHILD_COLUMNS:
        if col_name in ["تاريخ التسجيل", "اسم المستخدم", "الرقم القومى للام"]:
            continue
            
        if f"c_{col_name}" not in st.session_state:
            st.session_state[f"c_{col_name}"] = ""
            
        if col_name in DROPDOWN_OPTIONS:
            options = DROPDOWN_OPTIONS[col_name]
            current_val = st.session_state.get(f"c_{col_name}", options[0])
            idx = options.index(current_val) if current_val in options else 0
            form_data[col_name] = st.selectbox(col_name, options, index=idx, key=f"c_{col_name}")
        else:
            form_data[col_name] = st.text_input(col_name, key=f"c_{col_name}")
    
    if st.button("💾 حفظ بيانات الطفل", use_container_width=True):
        if not nat_id_mom or len(nat_id_mom) != 14:
            st.error("برجاء إدخال الرقم القومي للأم صحيحاً مكوناً من 14 رقماً!")
        else:
            form_data["تاريخ التسجيل"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            form_data["اسم المستخدم"] = st.session_state.name
            
            # --- المنطق البرمجي لتوزيع خيارات مكان المتابعة ومصدر الإحالة تلقائياً ---
            place_val = form_data.get("مكان المتابعة", "")
            form_data["وحدة"] = "تم" if place_val == "وحدة" else ""
            form_data["مستشفى"] = "تم" if place_val == "مستشفى" else ""
            form_data["أخرى"] = "تم" if place_val == "أخرى" else ""

            source_val = form_data.get("مصدر الاحاله", "")
            form_data["مستشفى الولادة"] = "تم" if source_val == "مستشفى الولادة" else ""
            form_data["عيادة خاصة"] = "تم" if source_val == "عيادة خاصة" else ""
            form_data["عيادة التطعيمات"] = "تم" if source_val == "عيادة التطعيمات" else ""
            form_data["نصيحة"] = "تم" if source_val == "نصيحة" else ""
            # ----------------------------------------------------------------------
            
            new_df = pd.DataFrame([form_data], dtype=str)
            excel = pd.ExcelFile(EXCEL_FILE)
            all_dfs = {s: pd.read_excel(excel, sheet_name=s, dtype=str) for s in excel.sheet_names}
            
            if "سجل المشورة للاطفال" in all_dfs:
                all_dfs["سجل المشورة للاطفال"] = pd.concat([all_dfs["سجل المشورة للاطفال"], new_df], ignore_index=True)
            else:
                all_dfs["سجل المشورة للاطفال"] = new_df
                
            with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
                for s, df in all_dfs.items():
                    df.to_excel(writer, sheet_name=s, index=False)
            st.success("تم حفظ بيانات الطفل وتوزيع الاختيرات في أعمدة الإكسيل بنجاح! ✨")

# ==================== 4. استعراض البيانات والداشبورد ====================
elif menu == "استعراض البيانات والداشبورد":
    st.markdown("<h2>📊 استعراض السجلات والبيانات المحفوظة</h2>", unsafe_allow_html=True)
    if os.path.exists(EXCEL_FILE):
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
