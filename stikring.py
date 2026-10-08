import streamlit as st
import pandas as pd
import base64
import os

st.set_page_config(page_title="نظام طباعة ملصقات أوربت", layout="wide")

st.title("🖨️ نظام إنشاء وطباعة ملصقات التعبئة (التصميم العريض)")

# --- خيارات التحكم ---
st.write("### ⚙️ إعدادات الطباعة")

# 1. خيار عنوان الملصق (نوع الكويل)
coil_type = st.radio(
    "📌 اختر نوع الكويل (العنوان الرئيسي للملصق):",
    options=["Mill Finish Aluminum Coils", "Coated Coils for Aluminum Structure","Coated Coils for Steel Structure","Mill Finish Steel Coils"],
    horizontal=True
)

st.markdown("<br>", unsafe_allow_html=True)

# 2. خيارات الشعار والمنشأ
col_opt1, col_opt2 = st.columns(2)
with col_opt1:
    show_logo = st.checkbox("إظهار شعار الشركة (صورة Orbit)", value=True)
with col_opt2:
    show_origin = st.checkbox("إظهار بلد المنشأ (Made in Jordan)", value=True)

st.info("💡 تأكد من وجود صورة الشعار باسم 'orbit.jpg' في نفس المجلد الذي يحتوي على هذا البرنامج ليتم عرضها في الملصقات.")
st.markdown("---")

uploaded_file = st.file_uploader("ارفع ملف الإكسل (Excel) هنا", type=['xlsx', 'xls'])

def get_image_base64(image_path):
    if os.path.exists(image_path):
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode('utf-8')
    return None

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #e0e0e0;
        }}
        .sticker {{
            width: 260mm;
            height: 130mm;
            border: 3px solid black;
            margin: 10mm auto;
            background-color: white;
            display: flex;
            flex-direction: column;
            box-sizing: border-box;
            page-break-after: always;
        }}
        
        .top-section {{
            display: flex;
            height: 35%;
            border-bottom: 3px solid black;
        }}
        .logo-box {{
            width: 22%;
            border-right: 2px solid black;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            padding: 5px;
            text-align: center;
        }}
        
        .mid-top {{ width: 48%; border-right: 2px solid black; display: flex; flex-direction: column; }}
        .title-bar {{ font-size: 20px; font-weight: bold; text-align: center; padding: 10px; border-bottom: 2px dotted #555; flex: 1; display: flex; align-items: center; justify-content: center; text-transform: capitalize; }}
        .mid-info {{ display: flex; height: 50%; }}
        .info-cell {{ border-right: 2px dotted #555; padding: 5px 10px; display: flex; flex-direction: column; justify-content: center; flex: 1; }}
        .info-cell:last-child {{ border-right: none; }}
        .info-cell span {{ font-size: 14px; color: #555; margin-bottom: 3px; }}
        .info-cell strong {{ display: block; word-wrap: break-word; line-height: 1.1; }}

        .right-top {{ width: 30%; display: flex; flex-direction: column; }}
        .rt-row {{ display: flex; flex: 1; border-bottom: 2px dotted #555; }}
        .rt-row:last-child {{ border-bottom: none; }}
        .rt-cell {{ flex: 1; border-right: 2px dotted #555; padding: 5px 10px; display: flex; flex-direction: column; justify-content: center; }}
        .rt-cell:last-child {{ border-right: none; }}
        .rt-cell span {{ font-size: 14px; color: #555; margin-bottom: 3px; }}
        
        .bottom-section {{ display: flex; height: 65%; }}
        
        .weights-col {{ width: 22%; border-right: 3px solid black; display: flex; flex-direction: column; }}
        .weight-item {{ display: flex; flex: 1; border-bottom: 2px solid black; }}
        .weight-item.deduction {{ flex-direction: column; border-bottom: none; }}
        .w-vert {{ background-color: black; color: white; writing-mode: vertical-rl; transform: rotate(180deg); text-align: center; padding: 5px; font-weight: bold; font-size: 16px; width: 35px; display: flex; align-items: center; justify-content: center; }}
        .w-horiz {{ background-color: black; color: white; text-align: center; padding: 5px; font-weight: bold; font-size: 16px; border-bottom: 2px solid black; }}
        
        .w-val {{ flex: 1; display: flex; flex-direction: column; align-items: center; justify-content: center; position: relative; }}
        .w-val-kg {{ font-size: 32px; font-weight: 900; line-height: 1; color: #000; }}
        .w-val-kg sub {{ font-size: 12px; font-weight: normal; margin-left: 2px; }}
        .w-val-lbs {{ font-size: 28px; font-weight: 900; color: #000; margin-top: 6px; line-height: 1; }}
        .w-val-lbs sub {{ font-size: 12px; font-weight: normal; margin-left: 2px; }}

        .details-col {{ width: 78%; padding: 15px; display: flex; flex-direction: column; gap: 15px; justify-content: space-between; }}
        .d-row {{ display: flex; gap: 15px; }}
        .d-cell {{ border: 2px dotted #888; padding: 10px; flex: 1; display: flex; flex-direction: column; justify-content: center; }}
        .d-cell span {{ font-size: 15px; color: #555; margin-bottom: 5px; }}
        
        @media print {{
            @page {{ size: landscape; margin: 0; }}
            body {{ background-color: white; }}
            .sticker {{ margin: 0; width: 100%; height: 100vh; border: none; box-shadow: none; }}
        }}
    </style>
</head>
<body>
    {labels_content}
    <script>
        window.onload = function() {{ window.print(); }}
    </script>
</body>
</html>
"""

SINGLE_LABEL = """
    <div class="sticker">
        <div class="top-section">
            <div class="logo-box">
                {logo_content}
            </div>
            <div class="mid-top">
                <div class="title-bar">{coil_type_title}</div>
                <div class="mid-info">
                    <div class="info-cell"><span>Destination</span><strong style="font-size: 22px;">{destination}</strong></div>
                    <div class="info-cell" style="flex:1.5;"><span>Sales Order</span><strong style="font-size: 22px;">{sales_order}</strong></div>
                    <div class="info-cell"><span>PO</span><strong style="font-size: 34px;">{po}</strong></div>
                </div>
            </div>
            <div class="right-top">
                <div class="rt-row">
                    <div class="rt-cell" style="flex:1.5;"><span>Package #</span><strong style="font-size: 38px; line-height: 1;">{package}</strong></div>
                    <div class="rt-cell" style="align-items:center;"><strong style="font-size: 18px;">{origin_content}</strong></div>
                </div>
                <div class="rt-row">
                    <div class="rt-cell"><span>Pallet Size</span><strong style="font-size: 34px;">{pallet_size}</strong></div>
                    <div class="rt-cell"><span>Slits #</span><strong style="font-size: 26px;">{slits}</strong></div>
                </div>
            </div>
        </div>
        
        <div class="bottom-section">
            <div class="weights-col">
                <div class="weight-item">
                    <div class="w-vert">Net</div>
                    <div class="w-val">
                        <div class="w-val-kg"><strong>{net_kg}</strong><sub>KG</sub></div>
                        <div class="w-val-lbs"><strong>{net_lbs}</strong><sub>LBS</sub></div>
                    </div>
                </div>
                <div class="weight-item">
                    <div class="w-vert">Gross</div>
                    <div class="w-val">
                        <div class="w-val-kg"><strong>{gross_kg}</strong><sub>KG</sub></div>
                        <div class="w-val-lbs"><strong>{gross_lbs}</strong><sub>LBS</sub></div>
                    </div>
                </div>
                <div class="weight-item deduction">
                    <div class="w-horiz">Deduction</div>
                    <div class="w-val">
                        <div class="w-val-kg" style="margin-top: 10px;"><strong>{deduction}</strong><sub>KG</sub></div>
                    </div>
                </div>
            </div>
            
            <div class="details-col">
                <div class="d-row">
                    <div class="d-cell" style="width: 100%;">
                        <span>Customer Name</span><strong style="font-size: 26px;">{customer}</strong>
                    </div>
                </div>
                <div class="d-row">
                    <div class="d-cell" style="flex: 2;"><span>Description</span><strong style="font-size: 28px; line-height: 1.1;">{description}</strong></div>
                    <div class="d-cell"><span>Item Number</span><strong style="font-size: 36px;">{item_number}</strong></div>
                </div>
                <div class="d-row">
                    <div class="d-cell"><span>CustReference</span><strong style="font-size: 26px;">{cust_ref}</strong></div>
                    <div class="d-cell"><span>Coils</span><strong style="font-size: 34px;">{coil_details}</strong></div>
                    <div class="d-cell" style="flex:1.5;"><span>Dimensions Inch</span><strong style="font-size: 28px;">{dimensions}</strong></div>
                </div>
            </div>
        </div>
    </div>
"""

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.subheader("عينة من البيانات المقروءة:")
        st.dataframe(df.head(3))
        
        labels_html = ""
        
        # معالجة صورة الشعار
        logo_html = ""
        if show_logo:
            b64_logo = get_image_base64("orbit.jpg")
            if b64_logo:
                logo_html = f'<img src="data:image/jpeg;base64,{b64_logo}" style="max-width: 95%; max-height: 100px; object-fit: contain;">'
            else:
                logo_html = "<span style='color:red; font-size:12px;'>الصورة orbit.jpg مفقودة</span>"
        
        def get_val(row, possible_columns, default=""):
            row_keys = {str(k).replace(" ", "").strip().lower(): k for k in row.index}
            for col in possible_columns:
                col_norm = str(col).replace(" ", "").strip().lower()
                if col_norm in row_keys:
                    val = str(row[row_keys[col_norm]]).strip()
                    if val.lower() not in ["", "nan", "nat", "none", "null"]:
                        if val.endswith('.0'): val = val[:-2]
                        return val
            return default

        for index, row in df.iterrows():
            coil_num = get_val(row, ['Coil Number Details', 'CoilNumberDetails', 'Coil Number'])
            segment = get_val(row, ['Segments', 'Segments ', 'Segment', 'Segmen'])
            coil_details = f"{coil_num} {segment}" if coil_num or segment else ""
            
            # جلب الأوزان ومعالجتها بدون كسور
            net_kg = get_val(row, ['Net Weight', 'NetWeight'])
            gross_kg = get_val(row, ['Gross Weight', 'GrossWeight'])
            net_lbs = get_val(row, ['NET LBS', 'NETLBS'])
            gross_lbs = get_val(row, ['GRS LBS', 'GRSLBS', 'Gross LBS'])
            deduction_val = get_val(row, ['Deduction', 'Deductio'], '0')

            try: net_kg = str(int(float(net_kg)))
            except: pass
            
            try: gross_kg = str(int(float(gross_kg)))
            except: pass
            
            try: net_lbs = str(int(float(net_lbs)))
            except: pass
            
            try: gross_lbs = str(int(float(gross_lbs)))
            except: pass
            
            try: deduction_val = str(int(float(deduction_val)))
            except: pass
            
            origin_content = "Made In Jordan" if show_origin else ""

            label = SINGLE_LABEL.format(
                logo_content=logo_html,
                coil_type_title=coil_type, 
                origin_content=origin_content,
                customer=get_val(row, ['Customer Name', 'CustomerName']),
                destination=get_val(row, ['Destination', 'Destinatio']),
                sales_order=get_val(row, ['Sales Order', 'SalesOrde']),
                po=get_val(row, ['Customer PO #', 'Customer PO', 'PO']),
                package=get_val(row, ['PackageId', 'Package', 'Package#']),
                pallet_size=get_val(row, ['Pallet Size', 'PalletSiz']),
                item_number=get_val(row, ['Item Number', 'ItemNumbe']),
                description=get_val(row, ['Description']),
                dimensions=get_val(row, ['Dimensions Inch', 'DimensionsInch', 'Dimensions']),
                slits=get_val(row, ['Number Of Slits', 'NumberOfSli', 'Slits']),
                net_kg=net_kg,
                gross_kg=gross_kg,
                net_lbs=net_lbs,
                gross_lbs=gross_lbs,
                deduction=deduction_val,
                cust_ref=get_val(row, ['Custreference', 'Cust Reference', 'Cust Ref']),
                coil_details=coil_details
            )
            labels_html += label
            
        final_html = HTML_TEMPLATE.format(labels_content=labels_html)
        
        b64 = base64.b64encode(final_html.encode('utf-8')).decode('utf-8')
        href = f'<a href="data:text/html;base64,{b64}" download="Orbit_Labels_HugeFonts.html" target="_blank" style="text-decoration:none;"><button style="background-color:#4CAF50; color:white; padding:12px 24px; border:none; border-radius:5px; cursor:pointer; font-size:16px; font-weight:bold;">📥 تحميل الملصقات للطباعة</button></a>'
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(href, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"حدث خطأ أثناء قراءة البيانات: {e}")
        
