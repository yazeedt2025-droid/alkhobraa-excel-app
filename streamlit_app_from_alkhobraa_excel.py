# تطبيق Streamlit لعرض ملفات Excel (بالعربية)
# اسم الملف: streamlit_app_from_Alkhobraa_excel.py
# شرح: يقرأ ملف Excel (.xls أو .xlsx) ويعرض أزرارًا لكل شيت.
#       عند الضغط على أي زر يظهر محتوى الشيت مباشرة مع خيارات للبحث، اختيار الأعمدة، والتحميل.
# كيفية التشغيل محليًا:
# 1) أنشئ بيئة افتراضية (اختياري)
#    python -m venv venv
#    source venv/bin/activate  # على لينكس/ماك
#    venv\Scripts\activate     # على ويندوز
# 2) ثبت المتطلبات:
#    pip install streamlit pandas xlrd openpyxl
# 3) شغّل التطبيق:
#    streamlit run streamlit_app_from_Alkhobraa_excel.py

import streamlit as st
import pandas as pd
import io
from typing import Dict

st.set_page_config(page_title="نظام عرض البيانات - الخُبراء", layout="wide")

# ---------- واجهة المستخدم (عربية) ----------
st.title("📊 عرض بيانات Excel — اضغط زرًا لعرض البيانات")
st.markdown("""
هذا تطبيق ويب بسيط يقرأ ملف Excel (.xls أو .xlsx) ويُظهِر شيتات الملف كأزرار.
- يمكنك رفع الملف من جهازك أو ترك التطبيق ليحاول فتح الملف الافتراضي (إن وُجد).
- بعد الضغط على اسم الشيت يظهر المحتوى فورًا مع أدوات فرز/بحث/تحميل.
""")

# ---------- رفع الملف أو استخدام الملف الافتراضي ----------
uploaded = st.sidebar.file_uploader("رفع ملف Excel (.xls أو .xlsx)", type=["xls", "xlsx"]) 
use_default_path = False
DEFAULT_PATH = "/mnt/data/Alkhobraa Arabic plan 2025 (1) (1).xls"

if uploaded is None:
    st.sidebar.write("لم تقم برفع ملف. سيحاول التطبيق فتح الملف الافتراضي (إن وُجد).")
    if st.sidebar.button("استخدم الملف الافتراضي (إن وُجد)"):
        use_default_path = True
else:
    st.sidebar.success(f"تم رفع الملف: {uploaded.name}")

# ---------- دالة لقراءة ملف Excel إلى قاموس DataFrame لكل شيت ----------
@st.cache_data
def read_excel_file(file_like) -> Dict[str, pd.DataFrame]:
    # file_like يمكن أن يكون مسارًا (str) أو كائن BytesIO / UploadedFile
    try:
        if isinstance(file_like, str):
            # قراءة من مسار
            xls = pd.read_excel(file_like, sheet_name=None)
        else:
            # UploadedFile from Streamlit
            bytes_data = file_like.read()
            xls = pd.read_excel(io.BytesIO(bytes_data), sheet_name=None)
        return xls
    except Exception as e:
        # محاولة ثانية بمحركات مختلفة (xlrd / openpyxl)
        try:
            if isinstance(file_like, str):
                xls = pd.read_excel(file_like, sheet_name=None, engine='xlrd')
            else:
                bytes_data = file_like.read()
                xls = pd.read_excel(io.BytesIO(bytes_data), sheet_name=None, engine='xlrd')
            return xls
        except Exception as e2:
            st.error("حدث خطأ عند قراءة ملف Excel. تأكد من أن الملف صالح وأن الحزم (xlrd/openpyxl) مثبتة.")
            raise

# ---------- تحميل البيانات ----------
sheets_dict = None
if uploaded is not None:
    try:
        sheets_dict = read_excel_file(uploaded)
    except Exception:
        sheets_dict = None

if use_default_path:
    try:
        if not DEFAULT_PATH:
            st.sidebar.warning("لم يتم تحديد مسار افتراضي.")
        else:
            sheets_dict = read_excel_file(DEFAULT_PATH)
            st.sidebar.success(f"تم فتح الملف من: {DEFAULT_PATH}")
    except Exception:
        sheets_dict = None

if sheets_dict is None:
    st.info("ارفع ملف Excel أو استخدم الملف الافتراضي من الشريط الجانبي لبدء العرض.")
    st.stop()

# ---------- إنشاء أزرار لكل شيت ----------
sheet_names = list(sheets_dict.keys())
st.subheader("الشيتات المتاحة:")
cols = st.columns(3)
for i, name in enumerate(sheet_names):
    with cols[i % 3]:
        if st.button(name, key=f"btn_{i}"):
            st.session_state['selected_sheet'] = name

# ---------- عرض الشيت المحدد ----------
selected = st.session_state.get('selected_sheet', None)
if selected is None:
    st.info("اضغط على اسم الشيت الذي تريده من الأزرار أعلاه لعرضه.")
else:
    df = sheets_dict[selected]
    st.markdown(f"### الشيت: **{selected}** — عدد الصفوف: {len(df)}")

    # خيارات سريعة
    c1, c2, c3 = st.columns([1,1,2])
    with c1:
        if st.button("عرض كامل"):
            st.dataframe(df)
    with c2:
        if st.button("إعادة تحميل/تحديث"):
            # فقط إعادة تعيين الحالة (ستعاد القراءة عند إعادة التشغيل)
            st.experimental_rerun()
    with c3:
        search_query = st.text_input("بحث (فلترة أي خلية تحتوي على):")

    # اختيار أعمدة
    cols_list = df.columns.tolist()
    chosen = st.multiselect("اختر أعمدة للعرض (افتراضي: الكل)", cols_list, default=cols_list)
    display_df = df[chosen].copy()

    # تطبيق البحث (بصيغة نصية على كل القيم)
    if search_query:
        mask = display_df.apply(lambda row: row.astype(str).str.contains(search_query, case=False, na=False).any(), axis=1)
        display_df = display_df[mask]
        st.write(f"نتيجة البحث: {len(display_df)} صفوف")

    st.dataframe(display_df)

    # ملخص إحصائي مختصر
    if st.checkbox("أظهر ملخص إحصائي مختصر (للأعمدة الرقمية)"):
        try:
            st.write(display_df.describe())
        except Exception as e:
            st.write("لا يوجد أعمدة رقمية لعرض الملخص.")

    # زر تحميل CSV
    csv_bytes = display_df.to_csv(index=False).encode('utf-8-sig')
    st.download_button("تحميل الجزء المعروض كملف CSV", csv_bytes, file_name=f"{selected}.csv")

# ---------- نهاية التطبيق ----------
st.markdown("---")
st.caption("تم تطوير هذا العرض بواسطة مساعد — يمكن تخصيصه لإظهار 'أزرار جاهزة' لعرض أقسام محددة إذا زودتني بأسماء الأعمدة أو الشيتات المطلوب عرضها فورًا.")
