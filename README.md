# UDF Wizard

Word (.docx) ve PDF dosyalarını UYAP UDF formatına, UDF dosyalarını da Word/PDF'e
çeviren, **tamamen çevrimdışı** çalışan masaüstü uygulaması. Hiçbir veri internete
gönderilmez; her şey kendi bilgisayarınızda işlenir.

## Desteklenen dönüşümler

| Girdi | Çıktı seçenekleri |
|-------|-------------------|
| .udf  | .docx, .pdf       |
| .docx | .udf, .pdf        |
| .pdf  | .udf, .docx       |

Korunan biçim bilgisi: paragraf yapısı, hizalama (sol/orta/sağ/iki yana yasla)
ve paragraf düzeyinde kalınlık. **Toplu dönüştürme** desteklenir (aynı türden
birden çok dosyayı tek seferde seçip bir klasöre dönüştürebilirsiniz).

---

## Kurulum ve Çalıştırma

### Windows

1. Python 3 kurulu olmalı: https://www.python.org/downloads/
   Kurulum sırasında **"Add Python to PATH"** kutusunu işaretleyin.
2. `Baslat_Windows.bat` dosyasına çift tıklayın.
   - Gerekli kütüphaneler otomatik kurulur.
   - Masaüstüne **UDF Wizard** kısayolu otomatik oluşturulur (yalnızca ilk seferinde).
   - Sonraki açılışlarda masaüstü kısayolunu kullanabilirsiniz.

### Mac / Linux

1. Python 3 kurulu olmalı.
   - Mac: https://www.python.org/downloads/
   - Linux: `sudo apt install python3 python3-pip`
2. Terminalde bir kez çalıştırın:
   ```
   chmod +x Baslat_Mac_Linux.command
   ```
3. `Baslat_Mac_Linux.command` dosyasına çift tıklayın.
   - Gerekli kütüphaneler otomatik kurulur.
   - Masaüstüne kısayol otomatik oluşturulur (yalnızca ilk seferinde).

---

## Kullanım

1. **Dosya Seç** — bir veya birden çok dosya seçin (toplu için hepsi aynı türde olmalı).
2. **Hedef format** — açılır listeden çıktı türünü seçin.
3. **Dönüştür** — tek dosyada kayıt yeri, çoklu dosyada çıktı klasörü sorulur.

---

## Komut satırından (isteğe bağlı)

```
python udf_core.py girdi.docx cikti.udf
python udf_core.py belge.udf belge.pdf
```

### Elle kurulum

```
pip install customtkinter python-docx pdfplumber reportlab pillow
python udf_donusturucu.py
```

Arayüz **CustomTkinter** ile modern bir görünüme sahiptir; sağ üstteki düğmeyle
açık/koyu tema arasında geçiş yapabilirsiniz.

---

## Dosyalar

| Dosya | Açıklama |
|-------|----------|
| `udf_donusturucu.py` | Grafik arayüz — çalıştırılacak ana dosya |
| `udf_core.py` | Dönüştürme motoru — aynı klasörde olmalı |
| `Baslat_Windows.bat` | Windows başlatıcı (kütüphane kurulumu + kısayol) |
| `Baslat_Mac_Linux.command` | Mac/Linux başlatıcı (kütüphane kurulumu + kısayol) |
| `UDF_LOGO.png` | Uygulama ikonu (kaynak) |
| `UDF_LOGO.ico` | Windows kısayolu ikonu (otomatik oluşturulur) |

---

## Notlar ve sınırlar

- **PDF → UDF:** PDF'ten metin çıkarılır. Sütunlu/tablolu karmaşık görsel
  düzenler birebir korunmayabilir; metin akışı düz paragraflara dönüşür.
- UDF akışkan bir metin formatıdır; sayfa kırılımı, üstbilgi gibi öğeler
  PDF'teki gibi sabit değildir.
- Oluşan UDF dosyasını UYAP Editör'de açıp gözden geçirmeniz, gerekirse
  imzalamadan önce biçimi kontrol etmeniz önerilir.
- UYAP Editör sürümleri arasında küçük şema farkları olabilir; bir dosya
  açılmazsa UYAP Editör'de boş belgeye yapıştırıp UDF olarak kaydetmek en
  güvenli yöntemdir.
