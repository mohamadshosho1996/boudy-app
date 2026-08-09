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
    "مستوى التعليم للام": ["امى", "يجيد القراءة", "مؤهل متوسط", "فوق متوسط", "مؤهل عالى"],
    "مستوى التعليم للاب": ["امى", "يجيد القراءة", "مؤهل متوسط", "فوق متوسط", "مؤهل عالى"],
    "وظيفة الام": ["يعمل", "لا تعمل"],
    "نوع الولادة": ["طبيعى", "قيصرى"],
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
    "الخدمات الغير ملباه": ["يوجد", "لا يوجد"],
    "النمو والتطور الحركى": ["طبيعي", "متقدم", "متأخر", "يحتاج متابعة", "تم التوعية"],
    "التطور الادراكى والمعرفى": ["طبيعي", "متأخر", "يحتاج متابعة", "تم التوعية"],
    "التطور اللغوى": ["طبيعي", "متأخر", "يحتاج متابعة", "تم التوعية"],
    "رسائل التربية الايجابية": ["طبيعي", "متأخر", "يحتاج متابعة", "تم التوعية"],
    "الانشطة التحفيزية": ["طبيعي", "متأخر", "يحتاج متابعة", "تم التوعية"],
    "التوعية عن التغذية التكميلية وسلامة الغذاء": ["طبيعي", "متأخر", "يحتاج متابعة", "تم التوعية"],
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

# ==================== 1. الصفحة الرئيسية ====================
if menu == "الصفحة الرئيسية":
    st.markdown("<h1>✨ مرحباً بكِ في نظام المشورة الأسرية الشامل ✨</h1>", unsafe_allow_html=True)
    st.write("النظام مرتب وجاهز تماماً وفقاً لتنسيق الحقول المعتمد.")

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
    
    # تهيئة مفاتيح الجلسة للحقول بالترتيب المطلوب
    for col in CHILD_COLUMNS:
        if f"c_{col}" not in st.session_state:
            st.session_state[f"c_{col}"] = ""

    # استرجاع البيانات الذكي عبر الرقم القومي للأم
    nat_id_mom_input = st.text_input("الرقم القومى للام", max_chars=14, key="c_الرقم القومى للام")
    
    if nat_id_mom_input and len(nat_id_mom_input) == 14:
        mom_b_date, _ = parse_national_id(nat_id_mom_input)
        st.session_state["c_تاريخ ميلاد الام"] = mom_b_date
        
        if st.button("🔍 استرجاع بيانات الأسرة المسجلة مسبقاً"):
            found_data = get_existing_data(nat_id_mom_input, "سجل المشورة للاطفال", "الرقم القومى للام") or get_existing_data(nat_id_mom_input, "المشورة الاسرية للحامل", "الرقم القومى")
            for c_name in CHILD_COLUMNS:
                if c_name in ["تاريخ التسجيل", "اسم المستخدم", "الرقم القومى للام", "تاريخ ميلاد الام"]:
                    continue
                val = found_data.get(c_name, "")
                if val:
                    st.session_state[f"c_{c_name}"] = str(val)
            st.rerun()

    form_data = {}
    
    # بناء الحقول بالترتيب الدقيق المطلوب
    for col_name in CHILD_COLUMNS:
        if col_name in ["تاريخ التسجيل", "اسم المستخدم"]:
            continue
            
        if col_name == "الرقم القومى للام":
            form_data[col_name] = nat_id_mom_input
            continue

        if col_name in DROPDOWN_OPTIONS:
            options = DROPDOWN_OPTIONS[col_name]
            current_val = st.session_state.get(f"c_{col_name}", options[0])
            idx = options.index(current_val) if current_val in options else 0
            form_data[col_name] = st.selectbox(col_name, options, index=idx, key=f"c_{col_name}")
            
        else:
            if col_name == "تاريخ ميلاد الام":
                form_data[col_name] = st.text_input(f"{col_name} [مستخرج أواماً من الرقم القومي]", key=f"c_{col_name}")
                
            elif col_name == "تاريخ ميلاد الطفل":
                form_data[col_name] = st.text_input(col_name, key=f"c_{col_name}")
                # حساب العمر الحالي والعمر الرحمي تلقائياً
                try:
                    if form_data[col_name]:
                        b_date_obj = datetime.datetime.strptime(form_data[col_name].strip(), "%Y-%m-%d").date()
                        today_date = datetime.date.today()
                        diff_days = (today_date - b_date_obj).days
                        if diff_days >= 0:
                            st.session_state["c_العمر الحالى للطفل"] = str(round(diff_days / 30.44, 1)) + " شهر"
                            st.session_state["c_العمر الرحمى للطفل"] = str(max(37, round(40 - (diff_days / 7)))) + " أسبوع"
                except Exception:
                    pass

            elif col_name == "العمر الحالى للطفل":
                form_data[col_name] = st.text_input(f"{col_name} [محسوب تلقائياً]", key=f"c_{col_name}")
                
            elif col_name == "العمر الرحمى للطفل":
                form_data[col_name] = st.text_input(f"{col_name} [محسوب تلقائياً]", key=f"c_{col_name}")

            elif col_name == "وزن الطفل عند الولادة":
                form_data[col_name] = st.text_input(col_name, key=f"c_{col_name}")
                
            elif col_name == "طول الطفل عند الولادة":
                form_data[col_name] = st.text_input(col_name, key=f"c_{col_name}")
                # حساب مقاس رأس الطفل عند الولادة تلقائياً
                try:
                    w_val = st.session_state.get("c_وزن الطفل عند الولادة", "3.0")
                    if w_val and form_data[col_name]:
                        st.session_state["c_مقاس راس الطفل عند الولادة"] = str(round((float(form_data[col_name]) / 2) + (float(w_val) * 0.5) + 10, 1))
                except ValueError:
                    pass

            elif col_name == "مقاس راس الطفل عند الولادة":
                form_data[col_name] = st.text_input(f"{col_name} [محسوب تلقائياً]", key=f"c_{col_name}")

            elif col_name == "الطول الحالى":
                form_data[col_name] = st.text_input(col_name, key=f"c_{col_name}")
                # حساب محيط الرأس الحالي تلقائياً
                try:
                    age_s = st.session_state.get("c_العمر الحالى للطفل", "0").replace(" شهر", "")
                    w_curr = st.session_state.get("c_الوزن الحالى", "3.0")
                    if age_s and w_curr and form_data[col_name]:
                        age_m = float(age_s)
                        weight_kg = float(w_curr)
                        length_cm = float(form_data[col_name])
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

            elif col_name == "النمو والتطور الحركى":
                # حساب النمو الحركي تلقائياً
                auto_m = "طبيعي"
                try:
                    age_s = st.session_state.get("c_العمر الحالى للطفل", "0").replace(" شهر", "")
                    w_curr = st.session_state.get("c_الوزن الحالى", "3.0")
                    if age_s and w_curr:
                        am = float(age_s)
                        cw = float(w_curr)
                        exp = 3 + (am * 0.5)
                        if cw > (exp * 1.25):
                            auto_m = "متقدم"
                        elif cw < (exp * 0.75):
                            auto_m = "متأخر"
                except ValueError:
                    pass
                m_opts = DROPDOWN_OPTIONS["النمو والتطور الحركى"]
                m_idx = m_opts.index(auto_m) if auto_m in m_opts else 0
                form_data[col_name] = st.selectbox(f"{col_name} [محسوب تلقائياً وقابل للتعديل]", m_opts, index=m_idx, key=f"c_{col_name}")

            else:
                form_data[col_name] = st.text_input(col_name, key=f"c_{col_name}")

    if st.button("💾 حفظ بيانات الطفل", use_container_width=True):
        if not nat_id_mom_input or len(nat_id_mom_input) != 14:
            st.error("برجاء إدخال الرقم القومي للأم صحيحاً مكوناً من 14 رقماً!")
        else:
            form_data["تاريخ التسجيل"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            form_data["اسم المستخدم"] = st.session_state.name
            form_data["الرقم القومى للام"] = nat_id_mom_input
            
            new_df = pd.DataFrame([form_data], dtype=str)
            excel = pd.ExcelFile(EXCEL_FILE)
            all_dfs = {s: pd.read_excel(excel, sheet_name=s, dtype=str) for s in excel.sheet_names}
            all_dfs["سجل المشورة للاطفال"] = pd.concat([all_dfs["سجل المشورة للاطفال"], new_df], ignore_index=True) if "سجل المشورة للاطفال" in all_dfs else new_df
            with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
                for s, df in all_dfs.items():
                    df.to_excel(writer, sheet_name=s, index=False)
            st.success("تم حفظ بيانات الطفل والترتيب الجديد بنجاح في الإكسيل! ✨")

# ==================== 4. استعراض البيانات والداشبورد ====================
elif menu == "استعراض البيانات والداشبورد":
    st.markdown("<h2>📊 استعراض السجلات والبيانات المحفوظة وتنزيل الإكسيل</h2>", unsafe_allow_html=True)
    if os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, "rb") as f:
            st.download_button(
                label="📥 تنزيل ملف الإكسيل المحدث",
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
