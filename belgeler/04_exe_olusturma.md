# Windows için Kurulumsuz Uygulama Üretme

`kod/exe_olusturucu.py` dosyasının tamamı, Google Colab'da **tek bir hücreye**
yapıştırılıp çalıştırıldığında kurulum gerektirmeyen bir Windows `.exe` üretir.
Analiz kodu bu dosyanın içine base64 olarak gömülüdür; ayrıca dosya yüklemeniz
gerekmez.

## Kullanım

1. [colab.research.google.com](https://colab.research.google.com) adresinde yeni bir
   not defteri açın.
2. `kod/exe_olusturucu.py` içeriğinin tamamını tek bir hücreye yapıştırın.
3. **Çalışma zamanı → Tümünü çalıştır**.
4. İşlem 10–15 dakika sürer; bitince `.exe` dosyası tarayıcınıza iner.

## Adımlar

Hücre beş adımda ilerler ve her adımı ekrana yazar:

| Adım | İşlem | Süre |
|---|---|---|
| 1/5 | Kaynak kod hazırlanır ve derlenebilirliği sınanır | saniyeler |
| 2/5 | Wine kurulur (Windows programlarını Linux'ta çalıştırma katmanı) | 3–5 dk |
| 3/5 | Windows için Python kurulur (MSI bileşenleriyle) | 3–5 dk |
| 4/5 | openpyxl, xlrd ve PyInstaller kurulur | 2–4 dk |
| 5/5 | `.exe` derlenir ve indirilir | 3–5 dk |

## Neden Wine gerekiyor

Colab Linux üzerinde çalışır ve doğrudan Windows çalıştırılabiliri üretemez. Wine,
Windows Python'u ve PyInstaller'ı Linux içinde çalıştırmayı sağlar.

İki ayrıntı bu akışın takılmadan çalışmasını sağlar ve değiştirilmemelidir:

- Windows Python, tam kurulum `.exe`'si yerine **MSI bileşenleriyle** (`core`, `exe`,
  `lib`, `tcltk`) sessiz kurulur. Tam kurulum dosyası Wine içinde görünmez bir
  kurulum penceresi açar ve süreç yanıt vermeyi bırakır.
- `WINEDLLOVERRIDES="mscoree,mshtml="` ortam değişkeni Mono ve Gecko kurulum
  diyaloglarını kapatır.

Analiz yalnızca `openpyxl` ve `xlrd` kullandığı, `pandas` ve `numpy` içermediği için
derleme hızlıdır ve üretilen dosya küçüktür (yaklaşık 20–30 MB).

## Bilinmesi gerekenler

**SmartScreen uyarısı.** Windows imzasız programları engeller. "Daha fazla bilgi" →
"Yine de çalıştır" ile geçilir.

**Antivirüs.** PyInstaller ile paketlenmiş programlar sık sık yanlış pozitif verir.
Kurum bilgisayarında bilgi işlemden istisna tanımlanması gerekebilir.

**Bir adım takılırsa.** Colab oturumunu yeniden başlatıp baştan çalıştırın; Wine
kurulumu ara sıra ilk denemede eksik kalır. Her komutta zaman aşımı tanımlıdır, bu
nedenle süreç sonsuza kadar beklemez, hata verip durur.

## Alternatif

Windows bilgisayarınızda Python kuruluysa Colab'a hiç gerek yoktur:

```
pip install openpyxl xlrd pyinstaller
pyinstaller --onefile --windowed --name IhaleAnaliz kod/ihale_analiz.py
```

`.exe` dosyası `dist/` klasöründe oluşur.
