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
    total_orders = len(df)
    total_quantity = df["Adet"].sum()
    avarage_order_value = df["Toplam"].mean()
    most_profitable_product = (
        df.groupby("Ürün")["Toplam"]
        .sum()
        .idxmax()   #returns the raw label of max value
    )
    # product_summary = df.groupby("Ürün")["Toplam", "Adet"].sum().reset_index()
    # newer one is:
    toplam = df.groupby("Ürün")["Toplam"].sum().reset_index()
    satis = df.groupby("Ürün")["Adet"].sum().reset_index()
    product_summary = pd.merge(toplam, satis, how="inner")
    

    df["Ay"] = df["Tarih"].dt.strftime("%Y-%m")

    monthly_summary = (
        df.groupby("Ay")["Toplam"]
        .sum()
        .reset_index()
    )

    total_summary_data = {
        "Birim": [
            "Satışların Toplamı",
            "Sipariş Sayısı",
            "Toplam Satış Sayısı",
            "Ortalama Sipariş Değeri",
            "En Karlı Ürün"
        ],
        "Değerler": [
            total_revenue,
            total_orders,
            total_quantity,
            int(avarage_order_value),
            most_profitable_product
        ]
    }
    total_summary = pd.DataFrame(total_summary_data)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Ham Veri", index=False)
        product_summary.to_excel(writer, sheet_name="Ürün Özeti", index=False)
        monthly_summary.to_excel(writer, sheet_name="Aylık Özet", index=False)
        total_summary.to_excel(writer, sheet_name="Genel Özet", index=False)

    # for CLI usage
    # print("Rapor oluşturuldu.")
    # print("Genel Ciro:", total_revenue)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Kullanım: python report.py input.xlsx output.xlsx")
        sys.exit(1)

    input_file = Path(sys.argv[1])
    output_file = Path(sys.argv[2])

    generate_report(input_file, output_file)