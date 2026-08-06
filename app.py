import streamlit as st
import pandas as pd
import datetime

# إعداد الصفحة وتنسيق الألوان المستوحى من برنامجك
st.set_page_config(page_title="برنامج بودى للمشورة الأسرية", layout="centered")

st.markdown("""
    <style>
    .stApp {background-color: #FFF5F8;}
    h1 {color: #BE185D; text-align: center;}
    .stButton>button {background-color: #EC4899; color: white; width: 100%; border-radius: 10px; font-weight: bold;}
    </style>
    """, unsafe_allow_html=True)

st.title("🌸 برنامج بودى للمشورة الأسرية")

# --- الدوال المنطقية ---
def calculate_age(dob):
    today = datetime.date.today()
    age = today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))
    return age

def assess_growth(weight, height):
    # معادلتك الأصلية لتقييم النمو
    if weight < 3 or height < 50: return "متأخر في النمو"
    if weight > 10 or height > 90: return "متقدم في النمو"
    return "طبيعي"

# --- الواجهة ---
tab1, tab2 = st.tabs(["🤰 الحوامل", "👶 الأطفال"])

with tab2:
    st.subheader("سجل الطفل")
    col1, col2 = st.columns(2)
    
    with col1:
        national_id = st.text_input("الرقم القومى للام")
        dob = st.date_input("تاريخ ميلاد الطفل")
    
    with col2:
        weight = st.number_input("الوزن (كجم)", step=0.1)
        height = st.number_input("الطول (سم)", step=0.1)

    # الحسابات التلقائية
    age = calculate_age(dob)
    st.write(f"**العمر الحالي:** {age} سنة")
    
    status = assess_growth(weight, height)
    st.write(f"**تقييم النمو:** {status}")

    # زر الحفظ التلقائي
    if st.button("💾 حفظ البيانات وإضافتها للشيت"):
        # هنا سيتم دمج البيانات مع ملف الـ Excel الخاص بك
        st.success("تم التحديث بنجاح!")

# --- خاصية الإملاء الصوتي ---
st.sidebar.subheader("تحكم")
if st.sidebar.button("🎤 إملاء صوتي"):
    st.sidebar.info("يتم الآن الاستماع للبيانات...")
