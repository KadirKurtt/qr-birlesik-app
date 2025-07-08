import os
import pandas as pd
from datetime import datetime


def process_qr_records(file_paths, output_path):
    merged_dataframes = []

    for file_path in file_paths:
        try:
            df = pd.read_excel(file_path, sheet_name="QR Records")
            df["SourceFile"] = os.path.basename(file_path)
            merged_dataframes.append(df)
        except Exception as e:
            print(f"Hata: {file_path} dosyasında sorun oluştu:\n{e}")
            continue

    if not merged_dataframes:
        print("Hiçbir dosya başarıyla okunamadı.")
        return

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

    # Gizlenecek sütunlar (ilk iki sütun + kullanıcıdan gelenler)
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

    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        merged_df.to_excel(writer, index=False, sheet_name="QR Records")
        worksheet = writer.sheets["QR Records"]
        for col in hide_cols:
            col_idx = merged_df.columns.get_loc(col)
            worksheet.set_column(col_idx, col_idx, None, None, {'hidden': True})

    print(f"İşlem tamamlandı! Dosya oluşturuldu: {output_path}")
    print(f"Birleştirilen toplam satır sayısı: {merged_df.shape[0]}")


def main():
    print("Birleştirmek istediğiniz dosyaların yolunu girin (her satıra bir dosya, bitince boş satır bırakın):")
    file_paths = []
    while True:
        path = input().strip().strip('"').strip("'")
        if not path:
            break
        if not os.path.isfile(path):
            print(f"Hata: {path} bulunamadı. Lütfen doğru yolu girin veya tekrar deneyin.")
            continue
        file_paths.append(path)

    if len(file_paths) < 1:
        print("En az bir dosya girmelisiniz!")
        return

    output_file = os.path.join(os.path.expanduser("~"), "Desktop", f"QR_Birlesik_{datetime.today().strftime('%Y%m%d')}.xlsx")
    process_qr_records(file_paths, output_file)

if __name__ == "__main__":
    main()
