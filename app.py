import streamlit as st
import pandas as pd
from weasyprint import HTML
import io

st.set_page_config(page_title="Holiday Energy Request Generator", layout="wide")

st.title("📊 เว็บไซต์แปลงข้อมูล Holiday Energy Request")
st.write("อัปโหลดไฟล์ Excel ดิบเพื่อแปลงเป็นตารางสรุปแผนงานประจำสัปดาห์ (วันเสาร์-อาทิตย์) ในรูปแบบ PDF ได้ทันที")

def process_excel(file):
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip()
    
    df['PERMIT_DATE_FROM'] = pd.to_datetime(df['PERMIT_DATE_FROM'], format='mixed')
    df['PERMIT_DATE_TO'] = pd.to_datetime(df['PERMIT_DATE_TO'], format='mixed')
    
    # ดึงวันเสาร์และอาทิตย์ที่เป็นกลุ่มหลักของข้อมูลในไฟล์อัตโนมัติ
    all_dates = pd.concat([df['PERMIT_DATE_FROM'].dt.date, df['PERMIT_DATE_TO'].dt.date]).unique()
    all_dates = sorted([d for d in all_dates if d.weekday() in [5, 6]])
    
    if len(all_dates) < 2:
        sat_date = pd.Timestamp('2026-07-04').date()
        sun_date = pd.Timestamp('2026-07-05').date()
    else:
        sat_date = all_dates[0]
        sun_date = all_dates[1]
        
    df['SHOP_CLEAN'] = df['PERMIT_LOCATION_NAME'].str.replace(' shop', '', case=False).str.strip()
    
    df['FORMATTED_TASK'] = df.apply(lambda r: f"-{r['PERMIT_NAME']} : {r['REQ_SECTION_CODE']}", axis=1)
    
    sat_mask = (df['PERMIT_DATE_FROM'].dt.date <= sat_date) & (df['PERMIT_DATE_TO'].dt.date >= sat_date)
    sun_mask = (df['PERMIT_DATE_FROM'].dt.date <= sun_date) & (df['PERMIT_DATE_TO'].dt.date >= sun_date)
    
    sat_grouped = df[sat_mask].groupby('SHOP_CLEAN')['FORMATTED_TASK'].apply(lambda x: '<br>'.join(x)).reset_index()
    sat_grouped.columns = ['Shop', 'Sat_Tasks']
    
    sun_grouped = df[sun_mask].groupby('SHOP_CLEAN')['FORMATTED_TASK'].apply(lambda x: '<br>'.join(x)).reset_index()
    sun_grouped.columns = ['Shop', 'Sun_Tasks']
    
    shop_list = ['Paint', 'Body', 'Frame', 'LA', 'UA', 'PT7-8', 'FR, KD']
    all_shops = pd.DataFrame({'Shop': shop_list})
    
    summary = all_shops.merge(sat_grouped, on='Shop', how='left')
    summary = summary.merge(sun_grouped, on='Shop', how='left')
    
    summary['Sat_Tasks'] = summary['Sat_Tasks'].fillna('X')
    summary['Sun_Tasks'] = summary['Sun_Tasks'].fillna('X')
    
    return summary, sat_date, sun_date

def generate_pdf_html(summary_df, sat_date, sun_date):
    sat_label = sat_date.strftime("%d %B'%y")
    sun_label = sun_date.strftime("%d %B'%y")
    header_label = f"{sat_date.strftime('%d')}-{sun_date.strftime('%d %B\'%y')}"
    
    rows_html = ""
    for idx, row in summary_df.iterrows():
        no = idx + 1
        shop = row['Shop']
        sat = row['Sat_Tasks']
        sun = row['Sun_Tasks']
        
        sat_style = 'background-color: #e6f7ff; color: #004d80;' if sat != 'X' else 'background-color: #ffcccc; color: #cc0000; text-align: center; font-weight: bold;'
        sun_style = 'background-color: #f6ffed; color: #274e13;' if sun != 'X' else 'background-color: #ffcccc; color: #cc0000; text-align: center; font-weight: bold;'
        
        sat_content = sat if sat == 'X' else f"<div style='text-align:center;font-weight:bold;margin-bottom:5px;'>O<br>(08.00-16:30 น.)</div>{sat}"
        sun_content = sun if sun == 'X' else f"<div style='text-align:center;font-weight:bold;margin-bottom:5px;'>O<br>(08.00-16:30 น.)</div>{sun}"
        
        rows_html += f'''
        <tr>
            <td style="text-align: center; width: 5%;">{no}</td>
            <td style="text-align: center; font-weight: bold; width: 10%;">{shop}</td>
            <td style="{sat_style} width: 42.5%; padding: 8px; vertical-align: top; font-size: 10pt;">{sat_content}</td>
            <td style="{sun_style} width: 42.5%; padding: 8px; vertical-align: top; font-size: 10pt;">{sun_content}</td>
        </tr>
        '''
        
    return f'''
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            @page {{ size: A4 portrait; margin: 15mm 12mm; }}
            body {{ font-family: Arial, sans-serif; color: #333; }}
            .header {{ font-size: 16pt; font-weight: bold; margin-bottom: 15px; }}
            table {{ width: 100%; border-collapse: collapse; }}
            th, td {{ border: 1px solid #333; padding: 6px; }}
            .bg-yellow {{ background-color: #fffbe6; }}
        </style>
    </head>
    <body>
        <div class="header">Energy Request Holiday : {header_label}</div>
        <table>
            <thead>
                <tr>
                    <th rowspan="2" style="width: 5%;">No.</th>
                    <th rowspan="2" style="width: 10%;">Shop</th>
                    <th class="bg-yellow" style="width: 42.5%;">Holiday</th>
                    <th class="bg-yellow" style="width: 42.5%;">Holiday</th>
                </tr>
                <tr>
                    <th class="bg-yellow">Sat ({sat_label})</th>
                    <th class="bg-yellow">Sun ({sun_label})</th>
                </tr>
            </thead>
            <tbody>{rows_html}</tbody>
        </table>
    </body>
    </html>
    '''

uploaded_file = st.file_uploader("เลือกไฟล์ตาราง Excel ดิบ (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        summary_table, sat_d, sun_d = process_excel(uploaded_file)
        st.success("ประมวลผลข้อมูลสำเร็จ!")
        
        display_df = summary_table.copy()
        display_df['Sat_Tasks'] = display_df['Sat_Tasks'].str.replace('<br>', '\n')
        display_df['Sun_Tasks'] = display_df['Sun_Tasks'].str.replace('<br>', '\n')
        st.dataframe(display_df, use_container_width=True)
        
        html_string = generate_pdf_html(summary_table, sat_d, sun_d)
        pdf_filename = f"Holiday_Energy_Request_{sat_d.strftime('%d-%b')}.pdf"
        
        pdf_buffer = io.BytesIO()
        HTML(string=html_string).write_pdf(pdf_buffer)
        
        st.download_button(
            label="📥 ดาวน์โหลดรายงานแบบ PDF (หน้าตาเหมือนรูปที่ 2)",
            data=pdf_buffer.getvalue(),
            file_name=pdf_filename,
            mime="application/pdf"
        )
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาด: {e}")
