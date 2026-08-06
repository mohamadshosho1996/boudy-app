import os
import json
import datetime
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# ==================== الثوابت وإعدادات البيانات ====================
EXCEL_FILE = "template.xlsx"
USERS_FILE = "users.json"

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

VISIT_SCHEDULE_MONTHS = {
    "الاسبوع الاول": 0.25, "عمر شهرين": 2, "عمر 4 شهور": 4, "عمر 6 شهور": 6,
    "عمر 9 شهور": 9, "عمر 12 شهر": 12, "عمر 18 شهر": 18, "عمر سنتين": 24,
    "عمر سنتين ونصف": 30, "عمر 3 سنين": 36, "عمر 3 سنين ونصف": 42, "عمر 4 سنين": 48,
    "عمر 4 سنين ونصف": 54, "عمر 5 سنين": 60, "عمر 5 سنين ونصف": 66, "عمر 6 سنين": 72
}

AUTO_FILL_DONE_FIELDS = [
    "فوائد الرضاعة الطبيعية والأوضاع وعلامات الجوع والشبع",
    "كفاية اللبن وكمية البراز",
    "إعطاء الجرعة اليومية من فيتامين د",
    "كيفية رعاية السرة والإهتمام بنظافة الطفل",
    "البطاقة الصحية وأهمية المتابعة الدورية ومنحنيات النمو",
    "أهمية الإلتزام بتطعيمات الطفل",
    "التغذية الصحية للأم المرضعة",
    "كيفية التعرف على علامات الخطورة"
]

PREGNANT_AUTO_FILL_DONE_FIELDS = [
    "التغذية السليمة", "المكملات الغذائية", "التمرينات الرياضية", "قسط من النوم والراحة",
    "المتابعة الدورية للحمل", "التحذير من تناول الأدوية بدون إستشارة طبيب والتعرض للتدخين والأبخرة",
    "المتاعب البسيطة في الشهور الأولى", "المتاعب في الشهور الأخيرة", "علامات الخطر أثناء الحمل",
    "مشاكل الولادة المبكرة وكيفية تجنبها", "حركة الجنين / معرفة جنس الجنين/ تمييز الأصوات من قبل الجنين",
    "تغير لون الجلد حول الحلمة وظهور بعض إفرازات من الثدي", "إرتداء الملابس الفضفاضة المريحة",
    "الإستعداد للولادة / تحضير ملابس المولود ، الخ", "علامات الولادة", "مميزات الولادة الطبيعية",
    "الساعة الذهبية الأولى", "ملامسة الجلد للجلد", "البداية المبكرة للرضاعة الطبيعية", "الرضاعة الطبيعية المطلقة",
    "أهمية المباعدة", "وسائل تنظيم الأسرة", "إستخدام وسيلة بعد الولادة مباشرة", "التطور العصبي والنفسي للطفل",
    "ملاحظات/ توصيات"
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
    "سبب دخول الحضانة": [
        "لا يوجد", "انخفاض وزن الطفل.", "احتياج الطفل لأدوية محددة بهذا الوقت.", "صعوبة شديدة في التنفس لعدم اكتمال نمو الرئتين.",
        "ارتفاع درجة حرارة جسم الرضيع.", "تعطل العمليات الحيوية بجسم الطفل.", "انخفاض معدل الجلوكوز في دم الطفل.",
        "معاناة الرضيع مشكلات في الجهاز الهضمي.", "إصابة الطفل بعدوى في الدم.", "إصابة الطفل بالصفراء.",
        "حدوث مشكلات خلال الولادة “الولادة المتعسرة أو الحمل الحرج”.", "وجود عيب خلقي يمنع الطفل عن التنفس أو الرضاعة بشكل طبيعي."
    ],
    "مكان المتابعة": ["وحدة", "مستشفى", "أخرى"],
    "مصدر الاحالة": ["مستشفى الولادة", "عيادة خاصة", "عيادة التطعيمات", "نصيحة"],
    "موعد الزيارة": VISIT_SCHEDULE_OPTIONS,
    "رضاعة طبيعية مطلقة": ["3 شهور", "4 شهور", "6 شهور"],
    "رضاعة طبيعية مع سوائل وأعشاب": ["تم", "لم يتم"],
    "رضاعة طبيعية مع صناعي": ["تم", "لم يتم"],
    "رضاعة لبن صناعي": ["تم", "لم يتم"],
    "موقف إستخدام وسيلة تنظيم أسرة": ["توجد", "لا يوجد", "وسؤال الحمل الجديد", "مرغوب", "غير مرغوب", "وسؤال الخدمات غير الملباة", "حدث", "لم يحدث"],
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

# ==================== المساعدات البرمجية ====================
def load_users():
    if not os.path.exists(USERS_FILE):
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_USERS, f, ensure_ascii=False, indent=4)
        return DEFAULT_USERS
    try:
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return DEFAULT_USERS

def extract_dob_from_national_id(nid):
    nid = str(nid).strip()
    if len(nid) == 14 and nid.isdigit():
        century_digit = nid[0]
        year_digits = nid[1:3]
        month_digits = nid[3:5]
        day_digits = nid[5:7]

        full_year = ('19' if century_digit == '2' else '20') + year_digits if century_digit in ['2', '3'] else ""
        if full_year:
            try:
                return datetime.date(int(full_year), int(month_digits), int(day_digits)).strftime("%Y-%m-%d")
            except ValueError:
                return ""
    return ""

def calculate_child_age(dob_str):
    if not dob_str:
        return "", ""
    try:
        dob = datetime.datetime.strptime(str(dob_str).strip(), "%Y-%m-%d").date()
        today = datetime.date.today()
        if dob > today:
            return "", ""
        
        diff_days = (today - dob).days
        if diff_days < 30:
            weeks = diff_days // 7
            c_age = f"{diff_days} يوم" if weeks == 0 else f"{weeks} أسابيع"
        else:
            months = max(1, (today.year - dob.year) * 12 + (today.month - dob.month) - (1 if today.day < dob.day else 0))
            c_age = f"{months} شهر"
        return c_age, "40 أسبوع"
    except Exception:
        return "", ""

def save_to_excel(sheet_name, row_data, columns):
    new_df = pd.DataFrame([row_data]).astype(str)
    if not os.path.exists(EXCEL_FILE):
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            new_df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        excel = pd.ExcelFile(EXCEL_FILE)
        all_dfs = {s: pd.read_excel(excel, sheet_name=s, dtype=str) for s in excel.sheet_names}
        if sheet_name in all_dfs:
            all_dfs[sheet_name] = pd.concat([all_dfs[sheet_name], new_df], ignore_index=True)
        else:
            all_dfs[sheet_name] = new_df

        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            for s_name, df_data in all_dfs.items():
                df_data.astype(str).to_excel(writer, sheet_name=s_name, index=False)

# ==================== إعدادات الواجهة والصفحة ====================
st.set_page_config(page_title="برنامج بودى للمشورة الأسرية", page_icon="🌸", layout="wide")

# إعداد جلسة الحفظ للمستخدم
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.user_info = {}

users = load_users()

# ==================== شاشة تسجيل الدخول ====================
if not st.session_state.logged_in:
    st.title("🌸 برنامج بودى للمشورة الأسرية 🌸")
    st.subheader("تسجيل الدخول")
    
    user_map = {f"{v['name']} ({k})": k for k, v in users.items()}
    selected_user_label = st.selectbox("اختر الحساب والطبيبة 🩺:", list(user_map.keys()))
    password_input = st.text_input("كلمة المرور:", type="password")

    if st.button("تسجيل الدخول ✨", use_container_width=True):
        username = user_map[selected_user_label]
        if users[username]["pass"] == password_input:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.session_state.user_info = users[username]
            st.rerun()
        else:
            st.error("كلمة المرور غير صحيحة، يُرجى التأكد وإعادة المحاولة!")

else:
    # ==================== الشريط الجانبي والتنقل ====================
    st.sidebar.title(f"أهلاً بكِ {st.session_state.user_info['name']}")
    page = st.sidebar.radio("القائمة الرئيسية", ["🤰 سجل المشورة للحوامل", "👶 سجل المشورة للأطفال", "📊 لوحة الإحصائيات (الداشبورد)"])
    
    if st.sidebar.button("🚪 تسجيل الخروج"):
        st.session_state.logged_in = False
        st.rerun()

    # خيار تحميل شيت الإكسيل مباشرة على الموبايل
    if os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, "rb") as f:
            st.sidebar.download_button("📥 تحميل ملف الإكسيل الحالي", f, file_name="template.xlsx", use_container_width=True)

    # ==================== 1. سجل الحوامل ====================
    if page == "🤰 سجل المشورة للحوامل":
        st.header("🤰 سجل المشورة الأسرية للحوامل")
        
        with st.form("pregnant_form"):
            form_data = {}
            cols = st.columns(2)
            
            for idx, col_name in enumerate(PREGNANT_COLUMNS):
                if col_name in ["تاريخ التسجيل", "اسم المستخدم"]:
                    continue
                
                target_col = cols[idx % 2]
                default_val = "تم" if col_name in PREGNANT_AUTO_FILL_DONE_FIELDS else ""
                
                if col_name in DROPDOWN_OPTIONS:
                    opts = DROPDOWN_OPTIONS[col_name]
                    form_data[col_name] = target_col.selectbox(col_name, opts)
                elif col_name == "تاريخ اخر دورة شهرية":
                    form_data[col_name] = str(target_col.date_input(col_name, value=datetime.date.today()))
                else:
                    form_data[col_name] = target_col.text_input(col_name, value=default_val)

            submit = st.form_submit_button("💾 حفظ وتصدير للإكسيل", use_container_width=True)
            if submit:
                form_data["تاريخ التسجيل"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                form_data["اسم المستخدم"] = st.session_state.username
                save_to_excel("المشورة الاسرية للحامل", form_data, PREGNANT_COLUMNS)
                st.success("تم حفظ البيانات بنجاح في ملف الإكسيل! ✨")

    # ==================== 2. سجل الأطفال ====================
    elif page == "👶 سجل المشورة للأطفال":
        st.header("👶 سجل المشورة الأسرية للأطفال")
        
        with st.form("child_form"):
            form_data = {}
            cols = st.columns(2)
            
            for idx, col_name in enumerate(CHILD_COLUMNS):
                if col_name in ["تاريخ التسجيل", "اسم المستخدم", "وحدة", "مستشفى", "أخرى", "مستشفى الولادة", "عيادة خاصة", "عيادة التطعيمات", "نصيحة"]:
                    continue
                
                target_col = cols[idx % 2]
                default_val = "تم" if col_name in AUTO_FILL_DONE_FIELDS else ""

                if col_name in DROPDOWN_OPTIONS:
                    form_data[col_name] = target_col.selectbox(col_name, DROPDOWN_OPTIONS[col_name])
                elif col_name == "تاريخ الميلاد للطفل":
                    dob_val = target_col.date_input(col_name, value=datetime.date.today())
                    form_data[col_name] = str(dob_val)
                    c_age, _ = calculate_child_age(str(dob_val))
                    form_data["العمر الحالى للطفل (شهور)"] = c_age
                else:
                    form_data[col_name] = target_col.text_input(col_name, value=default_val)

            submit = st.form_submit_button("💾 حفظ وتصدير للإكسيل", use_container_width=True)
            if submit:
                # معالجة حقول الاختيارات الشرطية
                place = form_data.get("مكان المتابعة", "")
                form_data["وحدة"] = "تم" if place == "وحدة" else ""
                form_data["مستشفى"] = "تم" if place == "مستشفى" else ""
                form_data["أخرى"] = "تم" if place in ["أخرى", "اخرى"] else ""

                ref = form_data.get("مصدر الاحالة", "")
                form_data["مستشفى الولادة"] = "تم" if ref == "مستشفى الولادة" else ""
                form_data["عيادة خاصة"] = "تم" if ref == "عيادة خاصة" else ""
                form_data["عيادة التطعيمات"] = "تم" if ref == "عيادة التطعيمات" else ""
                form_data["نصيحة"] = "تم" if ref == "نصيحة" else ""

                form_data["تاريخ التسجيل"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                form_data["اسم المستخدم"] = st.session_state.username
                save_to_excel("سجل المشورة للاطفال", form_data, CHILD_COLUMNS)
                st.success("تم حفظ بيانات الطفل بنجاح! ✨")

    # ==================== 3. الداشبورد والإحصائيات ====================
    elif page == "📊 لوحة الإحصائيات (الداشبورد)":
        st.header("📊 لوحة البيانات والإحصائيات")

        if not os.path.exists(EXCEL_FILE):
            st.info("لا يوجد بيانات مسجلة حالياً.")
        else:
            excel = pd.ExcelFile(EXCEL_FILE)
            user_counts = {u: {"حوامل": 0, "أطفال": 0} for u in users.keys()}

            if "المشورة الاسرية للحامل" in excel.sheet_names:
                df_p = pd.read_excel(excel, "المشورة الاسرية للحامل", dtype=str)
                if "اسم المستخدم" in df_p.columns:
                    for u in df_p["اسم المستخدم"]:
                        if u in user_counts: user_counts[u]["حوامل"] += 1

            if "سجل المشورة للاطفال" in excel.sheet_names:
                df_c = pd.read_excel(excel, "سجل المشورة للاطفال", dtype=str)
                if "اسم المستخدم" in df_c.columns:
                    for u in df_c["اسم المستخدم"]:
                        if u in user_counts: user_counts[u]["أطفال"] += 1

            # عرض التقرير في جدول
            summary_data = []
            for u, counts in user_counts.items():
                summary_data.append({
                    "الطبيبة": users[u]["name"],
                    "حالات الحوامل": counts["حوامل"],
                    "حالات الأطفال": counts["أطفال"],
                    "الإجمالي": counts["حوامل"] + counts["أطفال"]
                })
            
            st.dataframe(pd.DataFrame(summary_data), use_container_width=True)

            # رسم بياني توضيحي
            fig, ax = plt.subplots(figsize=(8, 4))
            names = [u["الطبيبة"] for u in summary_data]
            p_counts = [u["حالات الحوامل"] for u in summary_data]
            c_counts = [u["حالات الأطفال"] for u in summary_data]

            x = range(len(names))
            ax.bar([i - 0.2 for i in x], p_counts, 0.4, label='حوامل', color='#EC4899')
            ax.bar([i + 0.2 for i in x], c_counts, 0.4, label='أطفال', color='#A855F7')
            ax.set_xticks(x)
            ax.set_xticklabels(names)
            ax.legend()
            st.pyplot(fig)
