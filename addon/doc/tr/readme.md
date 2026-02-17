# EasyTableCopy - Kullanıcı Kılavuzu

**EasyTableCopy**, Web'den veya Windows listelerinden Word, Excel gibi uygulamalara veri aktarırken yaşanan format bozulmalarını, kaybolan kenarlıkları ve yapısal hataları çözmek için tasarlanmış bir NVDA eklentisidir.

## 1. Temel Özellikler

* **Çift Kopyalama Motoru:** Görsel kalite için "Standart" mod, yapısal garanti için "Yeniden Oluşturma" modu.
* **Akıllı Seçim:** Bitişik olmayan satırları veya sadece belirli sütunları kolayca seçip kopyalayın.
* **Masaüstü Desteği:** Dosya Gezgini veya uygulama listelerini anında düzenli bir tabloya dönüştürerek kopyalayın.
* **Sesli Geri Bildirim:** İşlem başladığında ve bittiğinde farklı seslerle sizi bilgilendirir.

## 2. Web Modu (Tarayıcılar)

Bu özellikler Chrome, Edge, Firefox gibi tarayıcılarda "Tarama Kipi" (Browse Mode) açıkken çalışır.

### A. Kopyalama Menüsü

Herhangi bir tablonun içindeyken `NVDA + Alt + T` tuşlarına basarak menüyü açabilirsiniz.

* **Tabloyu Kopyala (Standart):**
Tarayıcının kendi kopyalama yöntemini kullanır.
*Avantajı:* Yazı tiplerini, renkleri ve linkleri orijinal haliyle korur.
*Dezavantajı:* Tarayıcıya bağımlı olduğu için bazı boş hücreler Word'de kaybolabilir.
* **Tabloyu Kopyala (Yeniden Oluşturulan):**
"Güvenli Mod"dur. Eklenti tabloyu hücre hücre okur ve sıfırdan inşa eder.
*Avantajı:* **Boş hücrelerin korunmasını** ve kenarlıkların (border) Word'de kesinlikle görünmesini garanti eder.
*Dezavantajı:* Çok karmaşık görsel stilleri (yuvarlak köşeler vb.) aktarmayabilir.
* **Geçerli Satırı Kopyala:** İmlecin bulunduğu satırı kopyalar.
* **Geçerli Sütunu Kopyala:** İmlecin bulunduğu dikey sütunu manuel olarak ayıklar ve kopyalar.

### B. Akıllı Seçim (İşaretleme)

Tablonun tamamını değil, sadece istediğiniz kısımları alabilirsiniz. **Not:** Aynı anda hem satır hem sütun seçilemez.

#### Belirli Satırları Kopyalama:

1. İstediğiniz satıra gelin.
2. `Kontrol + Alt + Boşluk` tuşuna basın. NVDA "Satır İşaretlendi" diyecektir.
3. Başka bir satıra gidin ve tekrarlayın.
4. Menüyü açın (`NVDA + Alt + T`) ve **"İşaretlileri Kopyala"** seçeneğine tıklayın.

#### Belirli Sütunları Kopyalama:

1. İstediğiniz sütundaki bir hücreye gelin.
2. `Kontrol + Alt + Shift + Boşluk` tuşuna basın. NVDA "Sütun İşaretlendi" diyecektir.
3. Diğer sütunlar için tekrarlayın.
4. Menüyü açın ve **"İşaretlileri Kopyala"** deyin.

> **Not:** Seçimleri temizlemek için `Kontrol + Alt + Windows + Boşluk` tuşuna basabilir veya menüden "Seçimleri Temizle" diyebilirsiniz.

## 3. Masaüstü Modu (Windows)

EasyTableCopy, Dosya Gezgini ve standart liste görünümlerinde de çalışır.

1. Bir listeye odaklanın (Örneğin bir klasörün içine girin).
2. `NVDA + Alt + T` tuşuna basın.
3. **Sonuç:** Liste; "Ad", "Tarih", "Tür" gibi başlıkları olan düzenli bir tabloya dönüştürülür ve panoya kopyalanır.

## 4. Kısayol Listesi

*(Bu kısayolları NVDA Menüsü -> Tercihler -> Girdi Hareketleri -> EasyTableCopy kategorisinden değiştirebilirsiniz)*

| Kısayol | İşlev | Bağlam |
| --- | --- | --- |
| `NVDA + Alt + T` | Menüyü Aç (Web) / Listeyi Kopyala (Masaüstü) | Her Yerde |
| `Ctrl + Alt + Boşluk` | Satırı İşaretle/Kaldır | Sadece Web |
| `Ctrl + Alt + Shift + Boşluk` | Sütunu İşaretle/Kaldır | Sadece Web |
| `Ctrl + Alt + Win + Boşluk` | Seçimleri Temizle | Sadece Web |