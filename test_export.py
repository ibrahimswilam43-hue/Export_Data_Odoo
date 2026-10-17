
import io
import zipfile

import streamlit as st
from lxml import etree


# ======================= إعدادات الصفحة =======================
st.set_page_config(
    page_title="إلغاء Wrap Text",
    page_icon="🧹",
    layout="centered",
)

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

st.markdown(
    """
    <style>
    .block-container {
        max-width: 750px;
        padding-top: 3rem;
    }
    .stApp {
        direction: rtl;
    }
    h1, h2, h3, p, label {
        text-align: right;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ======================= إزالة Wrap Text =======================
def unwrap_xlsx(file_bytes):
    output = io.BytesIO()
    changed_formats = 0

    with zipfile.ZipFile(io.BytesIO(file_bytes), "r") as source:
        with zipfile.ZipFile(
            output,
            "w",
            compression=zipfile.ZIP_DEFLATED,
            compresslevel=1,
            allowZip64=True,
        ) as target:

            for item in source.infolist():
                data = source.read(item.filename)

                if item.filename == "xl/styles.xml":
                    root = etree.fromstring(data)

                    for alignment in root.iter(
                        f"{{{NS}}}alignment"
                    ):
                        if alignment.get("wrapText") in ("1", "true"):
                            del alignment.attrib["wrapText"]
                            changed_formats += 1

                    data = etree.tostring(
                        root,
                        encoding="UTF-8",
                        xml_declaration=True,
                    )

                target.writestr(item, data)

    return output.getvalue(), changed_formats


# ======================= واجهة التطبيق =======================
st.title("🧹 إلغاء Wrap Text من Excel")

st.write(
    "ارفع ملف Excel بصيغة XLSX، وسيتم إلغاء التفاف النص "
    "من جميع التنسيقات التي تستخدم Wrap Text، مع الحفاظ "
    "على البيانات وباقي التنسيقات قدر الإمكان."
)

uploaded_file = st.file_uploader(
    "📂 اختار ملف Excel",
    type=["xlsx"],
    help="يدعم ملفات Excel بصيغة XLSX.",
)

if uploaded_file is not None:
    st.info(f"الملف المحدد: {uploaded_file.name}")

    if st.button(
        "🚀 إلغاء Wrap Text",
        type="primary",
        use_container_width=True,
    ):
        try:
            with st.spinner("جاري تعديل الملف..."):
                file_bytes = uploaded_file.getvalue()
                fixed_file, changed_count = unwrap_xlsx(file_bytes)

                original_name = uploaded_file.name.rsplit(".", 1)[0]
                output_name = f"{original_name}_nowrap.xlsx"

                st.session_state["fixed_file"] = fixed_file
                st.session_state["fixed_name"] = output_name
                st.session_state["changed_count"] = changed_count

        except Exception as error:
            st.error(f"حصل خطأ أثناء تعديل الملف: {error}")

    if "fixed_file" in st.session_state:
        st.success("✅ تم تجهيز الملف بنجاح.")

        st.write(
            f"عدد تنسيقات Wrap Text التي تم تعديلها: "
            f"**{st.session_state['changed_count']:,}**"
        )

        st.download_button(
            label="⬇️ تنزيل ملف Excel بعد التعديل",
            data=st.session_state["fixed_file"],
            file_name=st.session_state["fixed_name"],
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary",
            use_container_width=True,
        )
