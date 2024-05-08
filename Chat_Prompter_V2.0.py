import pandas as pd
from docx import Document

def clean_percentage(value):
    return float(str(value).replace('%', '').replace(',', '.')) * 100

def generate_text(df, filename, is_region=False):
    doc = Document()
    grouped_df = df.groupby('Marka')
    for name, group in grouped_df:
        doc.add_heading('Bölge' if is_region else "Türkiye", level=1)  # Bölge veya Türkiye başlığı ekleniyor
        doc.add_heading(f"Ürün Adı: {name}", level=2)  # İlaç adına göre ana başlık oluşturuluyor
        df = df.sort_values(['Bölge', 'Marka'])

        for index, row in group.iterrows():
            text = f"""
            {row['Marka']} (YTD’DE) %{clean_percentage(row['PP_3.AY']):.2f} pazar payına sahiptir. Geçtiğimiz üç aylık (Ocak-Şubat-Mart) dönemde pp değişimleri sırasıyla %{clean_percentage(row['pp_değişim_1.ay']):.2f} , %{clean_percentage(row['pp_değişim_2.ay']):.2f}  ve %{clean_percentage(row['pp_değişim_3.ay']):.2f} olarak gerçekleşmiştir. Bu süreçte {row['Marka']}'in Türkiye'deki ortalama değişim oranı %{clean_percentage(row['pp_marg_ortalama']):.2f} olarak kaydedilmiştir. Geçtiğimiz yılın mart ayındaki Pazar payı yüzdesi %{clean_percentage(row['Mart_2023']):.2f}'dır. Bu da önceki yılın aynı dönemine göre pazar payımızda %{clean_percentage(row['PP_3.AY']) - clean_percentage(row['Mart_2023']):.2f} lik bir değişimin sağlandığını göstermektedir.İncelemeye alınan son üç aylık dönemde Gelişim İndeksi sırasıyla %{clean_percentage(row['GI_1.AY']):.2f}, %{clean_percentage(row['GI_2.AY']):.2f} ve %{clean_percentage(row['GI_3.AY']):.2f} olarak tespit edilmiştir.
            """
            # Mevcut trendlerin devam etmesi durumunda, 3 aylık süreçte {row['Marka']}'un net kutu hacminin {row['AylıkTahmin']} olacağı öngörülmektedir. ve ortalama olarak %{clean_percentage(row['GI_değişim_ortalama']):.2f} hesaplanmış
            doc.add_paragraph(text)
    doc.save(filename)

def process_data(df, filename, is_region=False):
    generate_text(df, filename, is_region)

def main():
    turkiye_df = pd.read_excel('C:\\Users\\melih.kalkan\\Desktop\\Turkiye_Bolge_Brick.xlsx', sheet_name='TURKIYE DATA')
    bolge_df = pd.read_excel('C:\\Users\\melih.kalkan\\Desktop\\Turkiye_Bolge_Brick.xlsx', sheet_name='BOLGE DATA')

    process_data(turkiye_df, "C:\\Users\\melih.kalkan\\Desktop\\Turkiye.docx")
    process_data(bolge_df, "C:\\Users\\melih.kalkan\\Desktop\\Bolge.docx", is_region=True)

if __name__ == "__main__":
    main()
