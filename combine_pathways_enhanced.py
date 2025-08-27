import pandas as pd
import os
import glob
from collections import Counter

def combine_pathway_names():
    """
    data0.xlsx'den data29.xlsx'e kadar olan dosyaların 3. sütunundaki 
    Pathway Name verilerini alıp yeni bir Excel dosyasında birleştirir.
    """
    
    # Excel dosyalarını bul (data0.xlsx'den data29.xlsx'e kadar)
    excel_files = []
    for i in range(30):  # 0'dan 29'a kadar
        filename = f"data{i}.xlsx"
        if os.path.exists(filename):
            excel_files.append(filename)
    
    # Dosyaları sırala
    excel_files.sort()
    
    print(f"Bulunan Excel dosyaları: {excel_files}")
    
    # Tüm Pathway Name verilerini saklamak için liste
    all_pathways = []
    all_pathway_names = []  # Tüm pathway isimlerini toplamak için
    
    # Her dosyayı oku ve 3. sütunu al
    for file in excel_files:
        try:
            # Excel dosyasını oku
            df = pd.read_excel(file)
            
            # 3. sütunu al (index 2, çünkü Python 0'dan başlar)
            if len(df.columns) >= 3:
                pathway_column = df.iloc[:, 2]  # 3. sütun (index 2)
                all_pathways.append(pathway_column)
                
                # Pathway isimlerini toplamaya ekle (NaN değerleri hariç)
                pathway_names = pathway_column.dropna().tolist()
                all_pathway_names.extend(pathway_names)
                
                print(f"{file}: {len(pathway_column)} satır Pathway Name verisi alındı")
            else:
                print(f"Uyarı: {file} dosyasında 3. sütun bulunamadı!")
                all_pathways.append(pd.Series([f"Veri yok - {file}"]))
                
        except Exception as e:
            print(f"Hata: {file} dosyası okunamadı - {e}")
            all_pathways.append(pd.Series([f"Hata - {file}"]))
    
    # Tüm verileri bir DataFrame'de birleştir
    if all_pathways:
        # En uzun sütunun uzunluğunu bul
        max_length = max(len(series) for series in all_pathways)
        
        # Tüm sütunları aynı uzunluğa getir
        normalized_pathways = []
        for i, series in enumerate(all_pathways):
            if len(series) < max_length:
                # Eksik satırları NaN ile doldur
                extended_series = series.reindex(range(max_length))
                normalized_pathways.append(extended_series)
            else:
                normalized_pathways.append(series)
        
        # DataFrame oluştur
        result_df = pd.DataFrame(normalized_pathways).T
        
        # Sütun isimlerini ayarla
        column_names = [f"Pathway_Names_{file.replace('.xlsx', '')}" for file in excel_files]
        result_df.columns = column_names
        
        # Sonucu Excel dosyasına kaydet
        output_filename = "combined_pathway_names.xlsx"
        result_df.to_excel(output_filename, index=False)
        
        print(f"\n✅ Başarılı! Tüm Pathway Name verileri '{output_filename}' dosyasına kaydedildi.")
        print(f"📊 Toplam {len(excel_files)} dosyadan veri alındı.")
        print(f"📋 Her sütun {max_length} satır içeriyor.")
        
        # İlk birkaç satırı göster
        print("\n📋 İlk 5 satır:")
        print(result_df.head())
        
        # En çok tekrar eden pathway'leri analiz et
        analyze_most_common_pathways(all_pathway_names)
        
        return output_filename
    else:
        print("❌ Hiçbir veri alınamadı!")
        return None

def analyze_most_common_pathways(pathway_names):
    """
    En çok tekrar eden pathway'leri analiz eder ve yazdırır.
    """
    if not pathway_names:
        print("❌ Analiz edilecek pathway verisi bulunamadı!")
        return
    
    # Pathway sayılarını hesapla
    pathway_counter = Counter(pathway_names)
    
    print(f"\n🔍 PATHWAY ANALİZİ:")
    print(f"📊 Toplam benzersiz pathway sayısı: {len(pathway_counter)}")
    print(f"�� Toplam pathway tekrarı: {sum(pathway_counter.values())}")
    
    # En çok tekrar eden 20 pathway'yi göster
    print(f"\n�� EN ÇOK TEKRAR EDEN PATHWAY'LER (İlk 20):")
    print("=" * 80)
    print(f"{'Sıra':<4} {'Tekrar Sayısı':<15} {'Pathway Adı':<60}")
    print("=" * 80)
    
    for i, (pathway, count) in enumerate(pathway_counter.most_common(20), 1):
        print(f"{i:<4} {count:<15} {pathway:<60}")
    
    # Tekrar eden pathway'lerin yüzdesini hesapla
    total_pathways = sum(pathway_counter.values())
    unique_pathways = len(pathway_counter)
    
    print(f"\n📈 İSTATİSTİKLER:")
    print(f"• Toplam pathway tekrarı: {total_pathways}")
    print(f"• Benzersiz pathway sayısı: {unique_pathways}")
    print(f"• Ortalama tekrar sayısı: {total_pathways/unique_pathways:.2f}")
    
    # Sadece 1 kez geçen pathway'ler
    single_occurrence = sum(1 for count in pathway_counter.values() if count == 1)
    print(f"• Sadece 1 kez geçen pathway sayısı: {single_occurrence}")
    print(f"• Birden fazla kez geçen pathway sayısı: {unique_pathways - single_occurrence}")
    
    # En çok tekrar eden pathway'leri Excel dosyasına da kaydet
    save_most_common_to_excel(pathway_counter)

def save_most_common_to_excel(pathway_counter):
    """
    En çok tekrar eden pathway'leri Excel dosyasına kaydeder.
    """
    # DataFrame oluştur
    df_most_common = pd.DataFrame([
        {'Sıra': i, 'Pathway_Adı': pathway, 'Tekrar_Sayısı': count}
        for i, (pathway, count) in enumerate(pathway_counter.most_common(), 1)
    ])
    
    # Excel dosyasına kaydet
    output_filename = "most_common_pathways.xlsx"
    df_most_common.to_excel(output_filename, index=False)
    
    print(f"\n�� En çok tekrar eden pathway'ler '{output_filename}' dosyasına kaydedildi.")

if __name__ == "__main__":
    print("🚀 Pathway Name verilerini birleştirme işlemi başlıyor...")
    result_file = combine_pathway_names()
    
    if result_file:
        print(f"\n🎉 İşlem tamamlandı! Sonuç dosyası: {result_file}")
    else:
        print("\n❌ İşlem başarısız oldu!")
