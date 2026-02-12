import pandas as pd
import sys
from pathlib import Path

def generate_report(input_path, output_path):

    df = pd.read_excel(input_path)

    df["Tarih"] = pd.to_datetime(df["Tarih"], errors="coerce")
    df["Adet"] = pd.to_numeric(df["Adet"], errors="coerce")
    df["Birim Fiyat"] = pd.to_numeric(df["Birim Fiyat"], errors="coerce")

    df.dropna(inplace=True)

    df["Toplam"] = df["Adet"] * df["Birim Fiyat"]

    total_revenue = df["Toplam"].sum()
    product_summary = df.groupby("Ürün")["Toplam"].sum().reset_index()
    
    df["Ay"] = df["Tarih"].dt.strftime("%Y-%m")

    monthly_summary = (
        df.groupby("Ay")["Toplam"]
        .sum()
        .reset_index()
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Ham Veri", index=False)
        product_summary.to_excel(writer, sheet_name="Ürün Özeti", index=False)
        monthly_summary.to_excel(writer, sheet_name="Aylık Özet", index=False)

    print("Rapor oluşturuldu.")
    print("Genel Ciro:", total_revenue)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Kullanım: python report.py input.xlsx output.xlsx")
        sys.exit(1)

    input_file = Path(sys.argv[1])
    output_file = Path(sys.argv[2])

    generate_report(input_file, output_file)