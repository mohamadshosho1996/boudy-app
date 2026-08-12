import datetime
import os
import pandas as pd
import streamlit as st

# ==================== إعدادات الصفحة والتصميم ====================
st.set_page_config(
    page_title="برنامج بودى للمشورة الأسرية",
    page_icon="🌸",
    layout="wide",
    initial_sidebar_state="expanded",
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
    "user3": {"pass": "1234", "role": "user", "name": "د. ايه 🎀"},
}

VISIT_SCHEDULE_OPTIONS = [
    "الاسبوع الاول",
    "عمر شهرين",
    "عمر 4 شهور",
    "عمر 6 شهور",
    "عمر 9 شهور",
    "عمر 12 شهر",
    "عمر 18 شهر",
    "عمر سنتين",
    "عمر سنتين ونصف",
    "عمر 3 سنين",
    "عمر 3 سنين ونصف",
    "عمر 4 سنين",
    "عمر 4 سنين ونصف",
    "عمر 5 سنين",
    "عمر 5 سنين ونصف",
    "عمر 6 سنين",
]

DROPDOWN_OPTIONS = {
    "مستوى التعليم": [
        "امى",
        "يجيد القراءة",
        "مؤهل متوسط",
        "فوق متوسط",
        "مؤهل عالى",
    ],
    "الوظيفة": ["يعمل", "لا تعمل"],
    "نوع الولادة": ["طبيعى", "قيصرى"],
    "قرابة بين الزوجين": ["نعم", "لا"],
    "وسيلة تنظيم الأسرة المستخدمة سابقا": [
        "توجد",
        "لا يوجد",
        "مرغوب",
        "غير مرغوب",
    ],
    "شهر الحمل": [
        "الشهر الاول",
        "الشهر الثانى",
        "الشهر الثالث",
        "الشهر الرابع",
        "الشهر الخامس",
        "الشهر السادس",
        "الشهر السابع",
        "الشهر الثامن",
        "الشهر التاسع",
    ],
    "أمراض مزمنة: إرتفاع ضغط الدم": ["تم", "لم يتم"],
    "أمراض مزمنة: السكر": ["تم", "لم يتم"],
    "أمراض مزمنة: إضطرابات الغدة": ["تم", "لم يتم"],
    "أمراض مزمنة: الأنيميا": ["تم", "لم يتم"],
    "مكملات \"قبل\": حمض الفوليك": ["لا يوجد", "تم", "لم يتم"],
    "مكملات \"قبل\": الحديد": ["لا يوجد", "تم", "لم يتم"],
    "مكملات \"قبل\": الكالسيوم": ["لا يوجد", "تم", "لم يتم"],
    "مكملات \"أثناء\": حمض الفوليك": ["تم", "لم يتم"],
    "مكملات \"أثناء\": الحديد": ["تم", "لم يتم"],
    "مكملات \"أثناء\": الكالسيوم": ["تم", "لم يتم"],
    "التغذية السليمة": ["تم", "لم يتم"],
    "المكملات الغذائية": ["تم", "لم يتم"],
    "التمرينات الرياضية": ["تم", "لم يتم"],
    "قسط من النوم والراحة": ["تم", "لم يتم"],
    "المتابعة الدورية للحمل": ["تم", "لم يتم"],
    "التحذير من تناول الأدوية بدون إستشارة طبيب والتعرض للتدخين والأبخرة": [
        "تم",
        "لم يتم",
    ],
    "المتاعب البسيطة في الشهور الأولى": ["تم", "لم يتم"],
    "المتاعب في الشهور الأخيرة": ["تم", "لم يتم"],
    "علامات الخطر أثناء الحمل": ["تم", "لم يتم"],
    "مشاكل الولادة المبكرة وكيفية تجنبها": ["تم", "لم يتم"],
    "حركة الجنين / معرفة جنس الجنين/ تمييز الأصوات من قبل الجنين": [
        "تم",
        "لم يتم",
    ],
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
    "مستوى التعليم للام": [
        "امى",
        "يجيد القراءة",
        "مؤهل متوسط",
        "فوق متوسط",
        "مؤهل عالى",
    ],
    "مستوى التعليم للاب": [
        "امى",
        "يجيد القراءة",
        "مؤهل متوسط",
        "فوق متوسط",
        "مؤهل عالى",
    ],
    "الوظيفة للام": ["يعمل", "لا تعمل"],
    "مكان الولادة": ["المستشفى", "المنزل"],
    "سبب دخول الحضانة": [
        "انخفاض وزن الطفل.",
        "احتياج الطفل لأدوية محددة بهذا الوقت.",
        "صعوبة شديدة في التنفس لعدم اكتمال نمو الرئتين.",
        "ارتفاع درجة حرارة جسم الرضيع.",
        "تعطل العمليات الحيوية بجسم الطفل.",
        "انخفاض معدل الجلوكوز في دم الطفل.",
        "معاناة الرضيع مشكلات في الجهاز الهضمي.",
        "إصابة الطفل بعدوى في الدم.",
        "إصابة الطفل بالصفراء.",
        "حدوث مشكلات خلال الولادة “الولادة المتعسرة أو الحمل الحرج”.",
        "وجود عيب خلقي يمنع الطفل عن التنفس أو الرضاعة بشكل طبيعي.",
    ],
    "موعد الزيارة": VISIT_SCHEDULE_OPTIONS,
    "رضاعة طبيعية مطلقة": ["تم", "لم يتم"],
    "رضاعة طبيعية مع سوائل وأعشاب": ["تم", "لم يتم"],
    "رضاعة طبيعية مع صناعي": ["تم", "لم يتم"],
    "رضاعة لبن صناعي": ["تم", "لم يتم"],
    "دخول الحضانة": ["تم", "لم يتم"],
    "ملامسة الجلد فى الساعة الذهبية الأولى": ["تم", "لم يتم"],
    "الرضاعة الطبيعية فى الساعة الذهبية الأولى": ["تم", "لم يتم"],
    "موقف إستخدام وسيلة تنظيم أسرة": [
        "توجد",
        "لا يوجد",
        "مرغوب",
        "غير مرغوب",
        "حدث",
        "لم يحدث",
    ],
    "الحمل الجديد": ["مرغوب", "غير مرغوب"],
    "الخدمات الغير ملباه": ["لا يوجد", "يوجد"],
    "تحويل الى عيادة تنظيم الاسره": ["تم", "لم يتم"],
    "النمو والتطور الحركي": ["طبيعى", "متقدم", "متاخر"],
    "التطور الإدراكي والمعرفي": ["طبيعى", "متقدم", "متاخر"],
    "التطور اللغوي": ["طبيعى", "متقدم", "متاخر"],
    "رسائل التربية الإيجابية": ["تم", "لم يتم"],
    "الأنشطة التحفيزية": ["تم", "لم يتم"],
    (
        "التوعية عن التغذية التكميلية وسلامة الغذاء والتغذية السليمة"
    ): ["تم", "لم يتم"],
    "إعطاء الجرعة اليومية من الحديد": ["يوجد", "لا يوجد"],
    "أهمية إستخدام وسيلة تنظيم أسرة وأهمية المباعدة": [
        "تم التوعيه",
        "لم يتم التوعيه",
    ],
    "إعطاء الجرعة اليومية من فيتامين د": ["يوجد", "لا يوجد"],
    "كيفية رعاية السرة والإهتمام بنظافة الطفل": ["تم", "لم يتم"],
    "البطاقة الصحية وأهمية المتابعة الدورية ومنحنيات النمو": [
        "تم",
        "لم يتم",
    ],
    "أهمية الإلتزام بتطعيمات الطفل": ["تم", "لم يتم"],
    "التغذية الصحية للأم المرضعة": ["تم", "لم يتم"],
    "كيفية التعرف على علامات الخطورة": ["تم", "لم يتم"],
    "فوائد الرضاعة الطبيعية والأوضاع وعلامات الجوع والشبع": [
        "تم",
        "لم يتم",
    ],
    "كفاية اللبن وكمية البراز": ["تم", "لم يتم"],
}

PREGNANT_COLUMNS = [
    "تاريخ التسجيل",
    "اسم المستخدم",
    "الاسم",
    "العنوان",
    "الرقم القومى",
    "رقم الموبايل",
    "العمر الحالى",
    "السن عند الزواج",
    "السن عند الحمل الاول",
    "مستوى التعليم",
    "الوظيفة",
    "تاريخ اخر دورة شهرية",
    "قرابة بين الزوجين",
    "عدد مرات الحمل",
    "عدد مرات الاجهاض",
    "عدد الاطفال",
    "المدة بين اخر حملين",
    "نوع الولادة",
    "أمراض مزمنة: إرتفاع ضغط الدم",
    "أمراض مزمنة: السكر",
    "أمراض مزمنة: إضطرابات الغدة",
    "أمراض مزمنة: الأنيميا",
    "أمراض مزمنة: اخرى",
    'مكملات "قبل": حمض الفوليك',
    'مكملات "قبل": الحديد',
    'مكملات "قبل": الكالسيوم',
    'مكملات "أثناء": حمض الفوليك',
    'مكملات "أثناء": الحديد',
    'مكملات "أثناء": الكالسيوم',
    "وسيلة تنظيم الأسرة المستخدمة سابقا",
    "مدة إستخدام الوسيلة السابقة",
    "شهر الحمل",
    "التاريخ الزيارة",
    "التغذية السليمة",
    "المكملات الغذائية",
    "التمرينات الرياضية",
    "قسط من النوم والراحة",
    "المتابعة الدورية للحمل",
    "التحذير من تناول الأدوية بدون إستشارة طبيب والتعرض للتدخين والأبخرة",
    "المتاعب البسيطة في الشهور الأولى",
    "المتاعب في الشهور الأخيرة",
    "علامات الخطر أثناء الحمل",
    "مشاكل الولادة المبكرة وكيفية تجنبها",
    "حركة الجنين / معرفة جنس الجنين/ تمييز الأصوات من قبل الجنين",
    "تغير لون الجلد حول الحلمة وظهور بعض إفرازات من الثدي",
    "إرتداء الملابس الفضفاضة المريحة",
    "الإستعداد للولادة / تحضير ملابس المولود ، الخ",
    "علامات الولادة",
    "مميزات الولادة الطبيعية",
    "الساعة الذهبية الأولى",
    "ملامسة الجلد للجلد",
    "البداية المبكرة للرضاعة الطبيعية",
    "الرضاعة الطبيعية المطلقة",
    "أهمية المباعدة",
    "وسائل تنظيم الأسرة",
    "إستخدام وسيلة بعد الولادة مباشرة",
    "التطور العصبي والنفسي للطفل",
    "ملاحظات/ توصيات",
    "تخطيط الزيارة القادمة",
    "المتابعة ما بعد الولادة",
]

CHILD_COLUMNS = [
    "تاريخ التسجيل",
    "اسم المستخدم",
    "تاريخ اول زيارة",
    "رقم الحالة",
    "اسم الام",
    "الرقم القومى للام",
    "رقم الموبايل للام",
    "تاريخ ميلاد الام",
    "مستوى التعليم للام",
    "عدد الاطفال لدى الام",
    "المدة بين اخر حملين",
    "الوظيفة للام",
    "الرقم القومى للاب",
    "رقم الموبايل للاب",
    "اسم الاب",
    "مستوى التعليم للاب",
    "اسم الطفل",
    "تاريخ الميلاد للطفل",
    "العمر الحالى للطفل (شهور)",
    "العمر الرحمى للطفل (أسابيع)",
    "مكان المتابعة (وحدة)",
    "مكان المتابعة (مستشفى)",
    "مكان المتابعة (اخرى)",
    "مصدر الاحالة(مستشفى الولادة)",
    "مصدر الاحالة (عيادة خاصة)",
    "مصدر الاحالة(عيادة التطعيمات)",
    "مصدر الاحالة(نصيحة)",
    "نوع الولادة",
    "مكان الولادة",
    "وزن الطفل عند الولادة",
    "طول الطفل عند الولادة",
    "مقاس راس الطفل عند الولادة",
    "دخول الحضانة",
    "سبب دخول الحضانة",
    "مدة البقاء فى الحضانة",
    "ملامسة الجلد فى الساعة الذهبية الأولى",
    "الرضاعة الطبيعية فى الساعة الذهبية الأولى",
    "موعد الزيارة",
    "تاريخ الزيارة",
    "رضاعة طبيعية مطلقة",
    "رضاعة طبيعية مع سوائل وأعشاب",
    "رضاعة طبيعية مع صناعي",
    "رضاعة لبن صناعي",
    "الوزن (كجم)",
    "الطول (سم)",
    "محيط الرأس (سم)",
    "فوائد الرضاعة الطبيعية والأوضاع وعلامات الجوع والشبع",
    "كفاية اللبن وكمية البراز",
    "إعطاء الجرعة اليومية من فيتامين د",
    "كيفية رعاية السرة والإهتمام بنظافة الطفل",
    "البطاقة الصحية وأهمية المتابعة الدورية ومنحنيات النمو",
    "أهمية الإلتزام بتطعيمات الطفل",
    "التغذية الصحية للأم المرضعة",
    "كيفية التعرف على علامات الخطورة",
    "النمو والتطور الحركي",
    "التطور الإدراكي والمعرفي",
    "التطور اللغوي",
    "رسائل التربية الإيجابية",
    "الأنشطة التحفيزية",
    "التوعية عن التغذية التكميلية وسلامة الغذاء والتغذية السليمة",
    "إعطاء الجرعة اليومية من الحديد",
    "أهمية إستخدام وسيلة تنظيم أسرة وأهمية المباعدة",
    "موقف إستخدام وسيلة تنظيم أسرة",
    "الحمل الجديد",
    "الخدمات الغير ملباه",
    "تحويل الى عيادة تنظيم الاسره",
    "تخطيط الزيارة القادمة",
]

YES_NO_CHECKBOX_FIELDS = [
    "مكان المتابعة (وحدة)",
    "مكان المتابعة (مستشفى)",
    "مكان المتابعة (اخرى)",
    "مصدر الاحالة(مستشفى الولادة)",
    "مصدر الاحالة (عيادة خاصة)",
    "مصدر الاحالة(عيادة التطعيمات)",
    "مصدر الاحالة(نصيحة)",
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
      age = (
          today.year
          - birth_date.year
          - ((today.month, today.day) < (birth_date.month, birth_date.day))
      )
      return str(birth_date), str(age)
    except ValueError:
      return "", ""
  return "", ""


def calculate_motor_development(
    age_str, weight_birth, length_birth, weight_current, length_current
):
  """تحليل وحساب التطور الحركي والنمو تلقائياً بناءً على المعدلات القياسية"""
  try:
    # استخلاص قيمة رقمية للعمر بالشهور
    if not age_str:
      return "طبيعى"
    if "يوم" in age_str or "أسبوع" in age_str:
      age_months = 0.5
    else:
      age_months = float(
          "".join(filter(lambda x: x.isdigit() or x == ".", age_str)) or 1
      )

    w_curr = float(weight_current) if weight_current else 3.5
    l_curr = float(length_current) if length_current else 50.0

    # التقديرات القياسية التقريبية المعيارية للوزن بالنسبة للعمر
    # عند الولادة ~3.2 كجم، يزيد الطفل تقريباً 600-800 جرام شهرياً فى أول 6 شهور
    if age_months <= 1:
      expected_weight = 3.3 + (age_months * 0.8)
    elif age_months <= 12:
      expected_weight = 3.0 + (age_months * 0.75)
    else:
      expected_weight = 10.0 + ((age_months - 12) * 0.2)

    # حساب النسبة المئوية للانحراف عن المعدل القياسي العالمي
    diff_ratio = w_curr / expected_weight

    if diff_ratio < 0.82:
      return "متاخر"
    elif diff_ratio > 1.25:
      return "متقدم"
    else:
      return "طبيعى"
  except Exception:
    return "طبيعى"


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
    pd.DataFrame(columns=PREGNANT_COLUMNS).to_excel(
        writer, sheet_name="المشورة الاسرية للحامل", index=False
    )
    pd.DataFrame(columns=CHILD_COLUMNS).to_excel(
        writer, sheet_name="سجل المشورة للاطفال", index=False
    )

# ==================== تسجيل الدخول والصلاحيات ====================
if "logged_in" not in st.session_state:
  st.session_state.logged_in = False
  st.session_state.user = None
  st.session_state.name = None
  st.session_state.role = None

if not st.session_state.logged_in:
  st.markdown(
      "<h2 style='text-align: center; color: #BE185D;'>🌸 برنامج بودى للمشورة"
      " الأسرية 🌸</h2>",
      unsafe_allow_html=True,
  )
  st.markdown(
      "<h4 style='text-align: center; color: #701A75;'>تسجيل الدخول للنظام</h4>",
      unsafe_allow_html=True,
  )

  col1, col2, col3 = st.columns([1, 2, 1])
  with col2:
    user_options = {
        f"{v['name']} ({k})": k for k, v in DEFAULT_USERS.items()
    }
    selected_display = st.selectbox(
        "اختر الحساب والطبيبة 🩺", list(user_options.keys())
    )
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
menu_options = [
    "الصفحة الرئيسية",
    "سجل الحوامل",
    "سجل الأطفال",
    "استعراض البيانات والداشبورد",
]
if st.session_state.role == "admin":
  menu_options.append("إدارة المستخدمين")

st.sidebar.markdown(f"### أهلاً بكِ د. {st.session_state.name} 🌸")
sidebar_menu = st.sidebar.radio(
    "القائمة الرئيسية (جانبية)", menu_options, key="sidebar_radio"
)

if st.sidebar.button("🚪 تسجيل الخروج"):
  st.session_state.logged_in = False
  st.rerun()

st.markdown("---")
col_mobile_nav, col_mobile_logout = st.columns([3, 1])
with col_mobile_nav:
  main_screen_menu = st.selectbox(
      "📱 انتقل مباشرة إلى القسم المطلوب:", menu_options, key="mobile_selectbox"
  )
with col_mobile_logout:
  if st.button("خروج 🚪"):
    st.session_state.logged_in = False
    st.rerun()

menu = main_screen_menu
st.markdown("---")

# ==================== 1. الصفحة الرئيسية ====================
if menu == "الصفحة الرئيسية":
  st.markdown(
      "<h1>✨ مرحباً بكِ في نظام المشورة الأسرية الشامل ✨</h1>",
      unsafe_allow_html=True,
  )
  st.write(
      "النظام جاهز تماماً مع خاصية الحساب التلقائي للنمو والتطور الحركي للطفل"
      " بناءً على التقديرات العالمية."
  )

# ==================== 2. سجل الحوامل ====================
elif menu == "سجل الحوامل":
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
    if col_name in DROPDOWN_OPTIONS:
      options = DROPDOWN_OPTIONS[col_name]
      st.markdown(f"**{col_name}**")
      current_val = st.session_state.get(f"p_{col_name}", options[0])
      chosen_choice = st.radio(
          f"اختر {col_name}",
          options,
          index=(
              options.index(current_val) if current_val in options else 0
          ),
          key=f"p_radio_{col_name}",
          horizontal=True,
      )
      form_data[col_name] = chosen_choice
      st.session_state[f"p_{col_name}"] = chosen_choice
    else:
      if col_name == "الرقم القومى":
        form_data[col_name] = st.text_input(
            col_name, max_chars=14, key=f"p_{col_name}"
        )
        if form_data[col_name] and len(form_data[col_name]) == 14:
          _, calc_age = parse_national_id(form_data[col_name])
          if calc_age:
            st.session_state["p_العمر الحالى"] = calc_age
      elif col_name == "العمر الحالى":
        form_data[col_name] = st.text_input(
            f"{col_name} [محسوب تلقائياً من الرقم القومي]",
            key=f"p_{col_name}",
        )
      elif col_name == "التاريخ الزيارة":
        form_data[col_name] = st.text_input(
            f"{col_name} [تاريخ اليوم التلقائي]", key=f"p_{col_name}"
        )
      else:
        form_data[col_name] = st.text_input(col_name, key=f"p_{col_name}")

  if st.button("💾 حفظ بيانات الحامل", use_container_width=True):
    final_form_data = {}
    for col in PREGNANT_COLUMNS:
      if col == "تاريخ التسجيل":
        final_form_data[col] = datetime.datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )
      elif col == "اسم المستخدم":
        final_form_data[col] = st.session_state.name
      else:
        final_form_data[col] = st.session_state.get(
            f"p_{col}", form_data.get(col, "")
        )

    new_df = pd.DataFrame([final_form_data], dtype=str)
    excel = pd.ExcelFile(EXCEL_FILE)
    all_dfs = {
        s: pd.read_excel(excel, sheet_name=s, dtype=str)
        for s in excel.sheet_names
    }

    for col in PREGNANT_COLUMNS:
      if col not in new_df.columns:
        new_df[col] = ""
    new_df = new_df[PREGNANT_COLUMNS]

    if "المشورة الاسرية للحامل" in all_dfs:
      all_dfs["المشورة الاسرية للحامل"] = pd.concat(
          [all_dfs["المشورة الاسرية للحامل"], new_df], ignore_index=True
      )
    else:
      all_dfs["المشورة الاسرية للحامل"] = new_df

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
      for s, df in all_dfs.items():
        df.to_excel(writer, sheet_name=s, index=False)
    st.success("تم حفظ بيانات الحامل بنجاح! ✨")

# ==================== 3. سجل الأطفال ====================
elif menu == "سجل الأطفال":
  st.markdown("<h2>👶 سجل المشورة الأسرية للأطفال</h2>", unsafe_allow_html=True)

  today_str = datetime.date.today().strftime("%Y-%m-%d")
  for col in CHILD_COLUMNS:
    if f"c_{col}" not in st.session_state:
      if col == "تاريخ الزيارة":
        st.session_state[f"c_{col}"] = today_str
      elif col == "تاريخ اول زيارة":
        st.session_state[f"c_{col}"] = today_str
      else:
        st.session_state[f"c_{col}"] = ""

  nat_id_mom_input = st.text_input(
      "الرقم القومى للام (اختياري)", max_chars=14, key="c_الرقم القومى للام_input"
  )
  if nat_id_mom_input:
    st.session_state["c_الرقم القومى للام"] = nat_id_mom_input

  if nat_id_mom_input and len(nat_id_mom_input) == 14:
    b_date_mom, _ = parse_national_id(nat_id_mom_input)
    if b_date_mom and not st.session_state.get("c_تاريخ ميلاد الام"):
      st.session_state["c_تاريخ ميلاد الام"] = b_date_mom

  if nat_id_mom_input and len(nat_id_mom_input) == 14:
    if st.button("🔍 استرجاع بيانات الأسرة المسجلة مسبقاً"):
      found_data = get_existing_data(
          nat_id_mom_input, "سجل المشورة للاطفال", "الرقم القومى للام"
      ) or get_existing_data(
          nat_id_mom_input, "المشورة الاسرية للحامل", "الرقم القومى"
      )
      for c_name in CHILD_COLUMNS:
        if c_name in ["تاريخ التسجيل", "اسم المستخدم", "الرقم القومى للام"]:
          continue
        val = found_data.get(c_name, "")
        if val:
          st.session_state[f"c_{c_name}"] = str(val)
      st.rerun()

  for col_name in CHILD_COLUMNS:
    if col_name in ["تاريخ التسجيل", "اسم المستخدم", "الرقم القومى للام"]:
      continue

    if col_name in YES_NO_CHECKBOX_FIELDS:
      checked = st.checkbox(col_name, value=False, key=f"c_chk_{col_name}")
      st.session_state[f"c_{col_name}"] = "نعم" if checked else ""

    elif col_name in DROPDOWN_OPTIONS:
      options = DROPDOWN_OPTIONS[col_name]

      # حساب التطور الحركي تلقائياً عند الوصول لحقل النمو والتطور الحركي
      if col_name == "النمو والتطور الحركي":
        auto_motor = calculate_motor_development(
            st.session_state.get("c_العمر الحالى للطفل (شهور)", ""),
            st.session_state.get("c_وزن الطفل عند الولادة", ""),
            st.session_state.get("c_طول الطفل عند الولادة", ""),
            st.session_state.get("c_الوزن (كجم)", ""),
            st.session_state.get("c_الطول (سم)", ""),
        )
        if not st.session_state.get(f"c_{col_name}"):
          st.session_state[f"c_{col_name}"] = auto_motor

      if col_name == "موعد الزيارة":
        auto_visit_choice = VISIT_SCHEDULE_OPTIONS[0]
        try:
          age_str = st.session_state.get("c_العمر الحالى للطفل (شهور)", "")
          if age_str:
            if "يوم" in age_str or "أسبوع" in age_str:
              auto_visit_choice = "الاسبوع الاول"
            else:
              age_num = float(
                  "".join(
                      filter(lambda x: x.isdigit() or x == ".", age_str)
                  )
                  or 0
              )
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

      st.markdown(f"**{col_name}**")
      current_val = st.session_state.get(f"c_{col_name}", options[0])
      chosen_choice = st.radio(
          f"اختر {col_name}",
          options,
          index=(
              options.index(current_val) if current_val in options else 0
          ),
          key=f"c_radio_{col_name}",
          horizontal=True,
      )
      st.session_state[f"c_{col_name}"] = chosen_choice

    else:
      if col_name == "تاريخ ميلاد الام":
        st.text_input(
            f"{col_name} [يتولد تلقائياً إذا أُدخل الرقم القومي للأم]",
            key=f"c_{col_name}",
        )

      elif col_name == "تاريخ الميلاد للطفل":
        default_date_val = datetime.date.today()
        existing_b_date = st.session_state.get(f"c_{col_name}", "")
        if existing_b_date:
          try:
            default_date_val = datetime.datetime.strptime(
                existing_b_date.strip(), "%Y-%m-%d"
            ).date()
          except Exception:
            pass

        chosen_date = st.date_input(
            col_name, value=default_date_val, key=f"c_date_input_{col_name}"
        )
        st.session_state[f"c_{col_name}"] = str(chosen_date)

        try:
          if st.session_state[f"c_{col_name}"]:
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
              st.session_state["c_العمر الحالى للطفل (شهور)"] = age_display

              gestational_weeks_calc = max(
                  24, min(42, 40 - max(0, round((280 - delta_days) / 7)))
              )
              st.session_state["c_العمر الرحمى للطفل (أسابيع)"] = (
                  f"{gestational_weeks_calc} أسبوع"
              )
        except Exception:
          pass

      elif col_name == "العمر الحالى للطفل (شهور)":
        st.text_input(f"{col_name} [محسوب تلقائياً]", key=f"c_{col_name}")

      elif col_name == "العمر الرحمى للطفل (أسابيع)":
        st.text_input(
            f"{col_name} [محسوب بدقة بناءً على تاريخ الميلاد]",
            key=f"c_{col_name}",
        )

      elif col_name == "وزن الطفل عند الولادة":
        st.text_input(col_name, key=f"c_{col_name}")

      elif col_name == "طول الطفل عند الولادة":
        st.text_input(col_name, key=f"c_{col_name}")
        try:
          w_val = st.session_state.get("c_وزن الطفل عند الولادة", "3.0")
          l_val = st.session_state.get("c_طول الطفل عند الولادة", "50.0")
          if w_val and l_val:
            st.session_state["c_مقاس راس الطفل عند الولادة"] = str(
                round((float(l_val) / 2) + (float(w_val) * 0.5) + 10, 1)
            )
        except Exception:
          pass
      else:
        st.text_input(col_name, key=f"c_{col_name}")

  if st.button("💾 حفظ بيانات الطفل", use_container_width=True):
    final_child_data = {}
    for col in CHILD_COLUMNS:
      if col == "تاريخ التسجيل":
        final_child_data[col] = datetime.datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )
      elif col == "اسم المستخدم":
        final_child_data[col] = st.session_state.name
      else:
        final_child_data[col] = st.session_state.get(f"c_{col}", "")

    new_child_df = pd.DataFrame([final_child_data], dtype=str)
    excel = pd.ExcelFile(EXCEL_FILE)
    all_dfs = {
        s: pd.read_excel(excel, sheet_name=s, dtype=str)
        for s in excel.sheet_names
    }

    for col in CHILD_COLUMNS:
      if col not in new_child_df.columns:
        new_child_df[col] = ""
    new_child_df = new_child_df[CHILD_COLUMNS]

    if "سجل المشورة للاطفال" in all_dfs:
      all_dfs["سجل المشورة للاطفال"] = pd.concat(
          [all_dfs["سجل المشورة للاطفال"], new_child_df], ignore_index=True
      )
    else:
      all_dfs["سجل المشورة للاطفال"] = new_child_df

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
      for s, df in all_dfs.items():
        df.to_excel(writer, sheet_name=s, index=False)
    st.success("تم حفظ بيانات الطفل بنجاح! ✨")

# ==================== 4. استعراض البيانات والداشبورد ====================
elif menu == "استعراض البيانات والداشبورد":
  st.markdown("<h2>📊 لوحة المؤشرات واستعراض البيانات</h2>", unsafe_allow_html=True)
  if os.path.exists(EXCEL_FILE):
    excel = pd.ExcelFile(EXCEL_FILE)
    sheet_to_show = st.selectbox("اختر السجل للاستعراض:", excel.sheet_names)
    df_view = pd.read_excel(excel, sheet_name=sheet_to_show, dtype=str)
    st.dataframe(df_view, use_container_width=True)
    st.info(f"إجمالي عدد الحالات المسجلة في هذا القسم: {len(df_view)}")
  else:
    st.warning("لا توجد بيانات مسجلة حتى الآن.")

# ==================== 5. إدارة المستخدمين ====================
elif menu == "إدارة المستخدمين" and st.session_state.role == "admin":
  st.markdown("<h2>⚙️ إدارة المستخدمين والصلاحيات</h2>", unsafe_allow_html=True)
  st.write("هنا يمكنك مراجعة حسابات الطبيبات والنظام المسجلة مسبقاً.")
  for k, v in DEFAULT_USERS.items():
    st.write(f"- **{v['name']}** | اسم المستخدم: `{k}` | الصلاحية: `{v['role']}`")
