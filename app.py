import streamlit as st
import pandas as pd
import os

# 1. إعدادات الصفحة
st.set_page_config(page_title="إدارة جرد المكاتب", layout="wide")
st.title("📦 إدارة المعدات حسب رقم المكتب")

# اسم ملف الإكسيل الخاص بك
FILE_PATH = "INVENTAIRE EN ORDRE.xlsx"

# 2. دالة لتحميل البيانات (مع تخزين مؤقت لتسريع التطبيق)
@st.cache_data
def load_data(file_path):
    if os.path.exists(file_path):
        return pd.read_excel(file_path)
    else:
        st.error(f"لم يتم العثور على الملف: {file_path}. يرجى التأكد من وجوده في نفس المجلد.")
        return pd.DataFrame()

# تحميل البيانات
df = load_data(FILE_PATH)

if not df.empty:
    # 3. استخراج أرقام المكاتب وإنشاء قائمة منسدلة
    # نستخرج الأرقام، نزيل الفراغات، ونرتبها
    def custom_sort(x):
        return (isinstance(x, str), x)
    bureaux = sorted(df['N° BUREAU'].dropna().unique(), key=custom_sort)
    selected_bureau = st.selectbox("📌 اختر رقم المكتب (N° BUREAU) :", bureaux)

    # 4. تصفية البيانات لعرض معدات المكتب المختار فقط
    df_bureau = df[df['N° BUREAU'] == selected_bureau].copy()
    # Resetting index is necessary to properly hide it and save horizontal space on mobile screens
    df_bureau = df_bureau.reset_index(drop=True) 

    st.subheader(f"📋 قائمة المعدات الخاصة بالمكتب رقم: {selected_bureau}")
    st.info("💡 يمكنك النقر على أي خلية لتعديلها. كما يمكنك إضافة صفوف جديدة أو حذفها من الأسفل.")

    # 5. عرض البيانات في جدول قابل للتعديل (Data Editor)
    # num_rows="dynamic" تسمح للمستخدم بإضافة أو حذف صفوف
    edited_df = st.data_editor(
        df_bureau, 
        num_rows="dynamic", 
        width="stretch",
        hide_index=True,
        column_config={
            "ORDRE": None,
            "OLDCODE": None,
            "CODE COMPTBLE": None,
            "CATEGORIE": None
        }
    )
    
    # عرض إجمالي عدد المعدات (Total Articles)
    st.markdown(f"**Total articles : {len(edited_df)}**")

    st.markdown("---")
    with st.expander("⚙️ إعدادات ملف PDF (خيارات الطباعة)"):
        st.write("يمكنك تعديل عرض الأعمدة واتجاه الصفحة ليناسب محتوى الطباعة:")
        col_w1, col_w2, col_w3, col_w4 = st.columns(4)
        w_col1 = col_w1.number_input("عرض كود محاسبي", min_value=10, max_value=100, value=40)
        w_col2 = col_w2.number_input("عرض OLDCODE2", min_value=10, max_value=100, value=45)
        w_col3 = col_w3.number_input("عرض DESIGNATION", min_value=10, max_value=200, value=120)
        w_col4 = col_w4.number_input("عرض OBSERVATION", min_value=10, max_value=100, value=72)
        
        col_opt1, col_opt2 = st.columns(2)
        orientation_choice = col_opt1.radio("اتجاه الصفحة", ["Paysage (أفقي)", "Portrait (عمودي)"], horizontal=True)
        l_height = col_opt2.number_input("ارتفاع سطر الجدول (لضغط المساحة)", min_value=3.0, max_value=15.0, value=5.0, step=0.5)

    # 6. أزرار الإجراءات
    import io
    from fpdf import FPDF
    col1, col2, col3 = st.columns(3)
    
    with col1:
        save_clicked = st.button("💾 Enregistrer les modifications", use_container_width=True)
        
    with col2:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            edited_df.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 Exporter la liste en excel",
            data=buffer.getvalue(),
            file_name=f"Inventaire_Bureau_{selected_bureau}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col3:
        def generate_pdf():
            ori = 'L' if "Paysage" in orientation_choice else 'P'
            pdf = FPDF(orientation=ori, unit='mm', format='A4')
            
            # Reduce margins and enable auto page break with small margin
            pdf.set_margins(left=5, top=5, right=5)
            pdf.set_auto_page_break(auto=True, margin=5)
            
            pdf.add_page()
            
            # Reduce header size to fit more items
            pdf.set_font("Arial", 'B', 14)
            
            dep_val = str(edited_df['DEP'].iloc[0]) if not edited_df.empty and 'DEP' in edited_df.columns else ""
            occ_val = str(edited_df['OCC'].iloc[0]) if not edited_df.empty and 'OCC' in edited_df.columns else ""
            bureau_val = str(selected_bureau)
            
            pdf.cell(0, 6, txt="FICHE DE BUREAU", ln=True, align='C')
            
            # En-tête sur une seule ligne (plus compact)
            pdf.set_font("Arial", 'B', 11)
            info_line = f"DEPARTEMENT: {dep_val}   |   N\xb0 BUREAU: {bureau_val}   |   OCCUPANT: {occ_val}   |   TOTAL ARTICLES: {len(edited_df)}"
            pdf.cell(0, 6, txt=info_line.encode('latin-1', 'replace').decode('latin-1'), ln=True, align='C')
            pdf.ln(2) # Small space before table
            
            col_map = {
                "CODE COMPT.": "CODE COMPTBLE" if "CODE COMPTBLE" in edited_df.columns else ("OLDCODE" if "OLDCODE" in edited_df.columns else "OLD CODE COMPTBLE"),
                "OLDCODE2": "OLDCODE2" if "OLDCODE2" in edited_df.columns else "OLDCODE2",
                "DESIGNATION": "DESIGNATION DU MAT\xc9RIEL" if "DESIGNATION DU MAT\xc9RIEL" in edited_df.columns else "DESIGNATION",
                "OBSERVATION": "OBSERVATION" if "OBSERVATION" in edited_df.columns else "OBSERVATION",
            }
            
            headers = ["CODE COMPT.", "OLDCODE2", "DESIGNATION", "OBSERVATION"]
            widths = [w_col1, w_col2, w_col3, w_col4]
            
            # Table Header
            pdf.set_font("Arial", 'B', 9)
            for i, h in enumerate(headers):
                pdf.cell(widths[i], 6, txt=h, border=1, align='C')
            pdf.ln()
            
            # Hauteur de ligne dynamique via les paramètres
            line_height = l_height
            pdf.set_font("Arial", '', 8)
            for _, row in edited_df.iterrows():
                v1 = str(row[col_map["CODE COMPT."]]) if col_map["CODE COMPT."] in edited_df.columns and pd.notna(row[col_map["CODE COMPT."]]) else ""
                v2 = str(row[col_map["OLDCODE2"]]) if col_map["OLDCODE2"] in edited_df.columns and pd.notna(row[col_map["OLDCODE2"]]) else ""
                v3 = str(row[col_map["DESIGNATION"]]) if col_map["DESIGNATION"] in edited_df.columns and pd.notna(row[col_map["DESIGNATION"]]) else ""
                v4 = str(row[col_map["OBSERVATION"]]) if col_map["OBSERVATION"] in edited_df.columns and pd.notna(row[col_map["OBSERVATION"]]) else ""
                
                # Truncation augmentée grâce au format paysage
                v1 = v1.encode('latin-1', 'replace').decode('latin-1')[:20]
                v2 = v2.encode('latin-1', 'replace').decode('latin-1')[:25]
                v3 = v3.encode('latin-1', 'replace').decode('latin-1')[:75] 
                v4 = v4.encode('latin-1', 'replace').decode('latin-1')[:45] 
                
                pdf.cell(widths[0], line_height, txt=v1, border=1)
                pdf.cell(widths[1], line_height, txt=v2, border=1)
                pdf.cell(widths[2], line_height, txt=v3, border=1)
                pdf.cell(widths[3], line_height, txt=v4, border=1)
                pdf.ln()
                
            out = pdf.output(dest='S')
            return out.encode('latin1') if isinstance(out, str) else out
            
        occ_val_for_name = str(edited_df['OCC'].iloc[0]) if not edited_df.empty and 'OCC' in edited_df.columns else "INCONNU"
        occ_val_for_name = occ_val_for_name.replace(" ", "_").replace("/", "_").replace("\\", "_")
        pdf_filename = f"FICHE_DE_BUREAU_N{selected_bureau}_OCC_{occ_val_for_name}.pdf"
        
        st.download_button(
            label="📄 Exporter en PDF",
            data=generate_pdf(),
            file_name=pdf_filename,
            mime="application/pdf",
            use_container_width=True
        )

    if save_clicked:
        try:
            # حذف البيانات القديمة الخاصة بهذا المكتب من الجدول الرئيسي
            df = df[df['N° BUREAU'] != selected_bureau]
            
            # إضافة البيانات الجديدة/المعدلة إلى الجدول الرئيسي
            df = pd.concat([df, edited_df], ignore_index=True)
            
            # إعادة ترتيب البيانات (اختياري)
            df = df.sort_values(by=['N° BUREAU'])
            
            # حفظ الملف
            df.to_excel(FILE_PATH, index=False)
            
            st.success("✅ تم حفظ التغييرات بنجاح!")
            # مسح التخزين المؤقت لتحديث البيانات في التطبيق
            st.cache_data.clear()
            
        except Exception as e:
            st.error(f"❌ حدث خطأ أثناء الحفظ: {e}")