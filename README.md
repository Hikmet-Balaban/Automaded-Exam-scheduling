# Automaded-Exam-scheduling
## Açıklama 
Bu python projesi önceden belirlenen sınavlar ve Çakışma matrixini inceleyerek optimal bir sınav programı hazırlar.

## Özellikler
- Çakışmasız sınav programı hazırlama
- Dinamik zaman slotları oluşturma
- Bir gün için aynı prefixten 3den fazla sınav yerleştirlmemesi.
- Outputları Excel dosyası olarak oluşturma

# Yükleme
1. repository'yi clonlayın
2. gerekli paketleri yükleyin(pandas,openpyxl)
3. Programı çalıştırın

# Kullanım 
-Predefined exam excel dosyasının yolunu (predefined_schedule_path = r'C:\...\Dekanlık Sınav programı (1).xlsx') şeklinde 5.line da tanımlayın.(pathin yazılım şekli aynı olmalı)
-Çakışma matrixinin yolunu (conflict_matrix_path = r'C:\....\Kitle Sınavları - Çakışma Matrisi Güz 2023.xlsx') formatında 38.line a yazın.
-Liste halindeki output için (output_file_path = r'C:\...\Yeni_Ders_Programı.xlsx') şeklinde output yolunu belirleyin ve 101.line ı ona göre düzenleyin.
-Takvim formatındaki output için 134.line a (output_file_path = r'C:\...\Final_Ders_Programı.xlsx') formatında yolu kopyalayın
Kodu çalıştırın ve oluşturulan excel dosyalarından verileri inceleyin.
!!!NOT!!!: DOSYALARIN PATH'i BELİRTİLEN ŞEKİLDE TANIMLANMALI!!!!!

