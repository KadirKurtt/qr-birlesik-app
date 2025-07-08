import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="QR Records Birleştirici", layout="centered")
st.title("QR Records Birleştirici Web Uygulaması")

st.write("Birden fazla Excel dosyasını yükleyin, QR Records sayfalarını birleştirin ve sonucu indirin.")

uploaded_files = st.file_uploader(
    "Excel dosyalarını seçin (birden fazla seçebilirsiniz)",
    type=["xlsx"],
    accept_multiple_files=True
)

def process_qr_records_streamlit(files):
    merged_dataframes = []
    for file in files:
        try:
            df = pd.read_excel(file, sheet_name="QR Records")
            df["SourceFile"] = file.name
            merged_dataframes.append(df)
        except Exception as e:
            st.warning(f"{file.name} dosyasında sorun oluştu: {e}")
            continue
    if not merged_dataframes:
        st.error("Hiçbir dosya başarıyla okunamadı.")
        return None
    merged_df = pd.concat(merged_dataframes, ignore_index=True)
    # Total Rejects sütunu (dinamik tarihli) bul
    reject_col = [col for col in merged_df.columns if "Total Rejects" in str(col)]
    if reject_col:
        merged_df = merged_df.sort_values(by=reject_col[0], ascending=False)
    # Site Name sütununu H sütununa taşı
    columns = merged_df.columns.tolist()
    if "Site Name" in columns:
        site_data = merged_df["Site Name"]
        insert_position = 7  # H sütunu (0-indeksli)
        if columns[insert_position] != "Site Name":
            columns = [col for col in columns if col != "Site Name"]
            columns.insert(insert_position, "Site Name")
            merged_df = merged_df[columns]
            merged_df["Site Name"] = site_data
    # Gizlenecek sütunlar
    hide_cols = merged_df.columns[:2].tolist()
    extra_hide = [
        'Bin Code',
        'Plant',
        'Incoming Quality Contact',
        'Contack',
        'System',
        'Unit Of Measurement',
        'Original ShipPoint Code',
        'Sort Qty',
        'Original Plant Code',
        'Impact PPM  02-Jul-2025'
    ]
    for col in extra_hide:
        if col in merged_df.columns and col not in hide_cols:
            hide_cols.append(col)
    return merged_df, hide_cols

if uploaded_files:
    if st.button("Birleştir ve İndir"):
        result = process_qr_records_streamlit(uploaded_files)
        if result is not None:
            merged_df, hide_cols = result
            # Excel dosyasını bellekte oluştur
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                merged_df.to_excel(writer, index=False, sheet_name="QR Records")
                worksheet = writer.sheets["QR Records"]
                for col in hide_cols:
                    col_idx = merged_df.columns.get_loc(col)
                    worksheet.set_column(col_idx, col_idx, None, None, {'hidden': True})
            output.seek(0)
            st.success(f"Birleştirme tamamlandı! Toplam satır: {merged_df.shape[0]}")
            st.download_button(
                label="Sonuç Excel Dosyasını İndir",
                data=output,
                file_name=f"QR_Birlesik_{datetime.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ) 