# Hastane Eczanesinde Bulunurluğa Göre Düzeltilmiş İlaç Talebi Tahmini

[![DOI](https://zenodo.org/badge/DOI/XX.XXXX/zenodo.XXXXXXX.svg)](https://doi.org/XX.XXXX/zenodo.XXXXXXX)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Hastane eczanelerinde yıllık ihale miktarının belirlenmesi için, rutin HBYS stok
hareket verisinden **bulunurluğa göre düzeltilmiş** talep tahmini üreten açık kaynaklı
araç ve çözümleme kodu.

> **Sorun.** İhale miktarı genellikle geçmiş tüketimin takvim süresine bölünmesiyle
> hesaplanır. Tüketim kaydı yalnızca ilaç eldeyken oluştuğundan, stoksuz geçen günler
> "talep yoktu" olarak sayılır ve gerçek talep sistematik biçimde düşük görünür. Bir
> kez stoksuz kalan kalem, ertesi dönem yine yetersiz alınır.
>
> **Çözüm.** Tüketimi takvim süresine değil, ilacın **fiilen stokta bulunduğu süreye**
> bölmek.
>
> ```
> Aylık Ortalama = (Toplam Çıkış ÷ Fiilen Stokta Bulunulan Gün) × 30
> ```

Yöntem, bir üniversite hastanesinin üç dönemlik verisiyle (398.401 hareket satırı,
875 etken madde) geliştirilmiş ve iki katmanlı kayan köken doğrulamasıyla sınanmıştır.
Ayrıntılar için ilgili makaleye bakınız (bkz. [Atıf](#atıf)).

---

## İçindekiler

| Klasör | İçerik |
|---|---|
| `kod/` | Analiz motoru, doğrulama betikleri, şekil üreteci, exe oluşturucu |
| `sorgu/` | HBYS'den veri çekme SQL sorguları ve doğrulama sorguları |
| `ornek_veri/` | Sentetik örnek veri ve üreteci |
| `belgeler/` | Yöntem belgesi, karar eşikleri, kullanım kılavuzu |

---

## Hızlı başlangıç

### Gereksinimler

```bash
pip install openpyxl xlrd
```

Yalnızca bu iki kütüphane gereklidir. `pandas`, `numpy` ya da makine öğrenmesi
kütüphanesi **kullanılmaz**. (Doğrulama ve şekil betikleri ayrıca `numpy`, `scipy` ve
`matplotlib` ister; ana analiz için gerekmez.)

### Örnek veriyle deneme

```bash
cd ornek_veri
python sentetik_veri_uret.py          # örnek dosyaları üretir
cd ..
python kod/ihale_analiz.py ornek_veri/ornek_hareket_2024.xlsx \
                           ornek_veri/ornek_hareket_2025.xlsx \
                           ornek_veri/ornek_hareket_2026.xlsx
```

Çıktı, çalışılan klasöre `ihale_analizi_<yıllar>_<zaman>.xlsx` adıyla yazılır.

### Arayüzle kullanım

```bash
python kod/ihale_analiz.py
```

Argümansız çalıştırıldığında masaüstü arayüzü açılır: dosyaları ekleyin,
**Tamamlandı** düğmesine basın.

### Windows için kurulumsuz uygulama

`kod/exe_olusturucu.py` içeriği, Google Colab'da tek hücre olarak çalıştırıldığında
kurulum gerektirmeyen bir Windows `.exe` üretir. Ayrıntı: `belgeler/exe_olusturma.md`.

---

## Gereken veri

HBYS'den aşağıdaki **yedi alan** çekilmelidir. Her etken madde kalemi için her güne
bir satır olması önerilir; satır bulunmayan günlerde yazılım son bilinen bakiyeyi
taşır.

| Alan | Tip | Açıklama |
|---|---|---|
| `DAY_ID` | Tarih | Hareket günü |
| `ETKEN_MADDE_KODU` | Sayı | Kalemin tekil kodu |
| `ETKEN_MADDE` | Metin | Doz ve formu içeren ad |
| `DEVIR` | Sayı | Dönem başı devreden miktar |
| `GIRIS` | Sayı | O gün stoğa giren |
| `CIKIS` | Sayı | O gün servise çıkan |
| `KALAN_SON` | Sayı | Gün sonu yürüyen stok |

`AYIN_GUNU` sütunu varsa tarih biçiminin doğrulanmasında kullanılır. `KALAN` (günlük
net hareket) hesaplamada kullanılmaz.

**Önemli iki nokta.** Çıkış miktarı **depodan servise fiilen çıkan** miktarı
yansıtmalıdır; hasta bazında faturalanan kullanım kaydı farklı bir büyüklüktür.
Analiz **etken madde kodu** üzerinden yürütülmelidir; ticari ürün düzeyindeki kayıt
sunum değişikliklerini gizler.

Hazır sorgular için `sorgu/01_veri_cekme_sorgulari.sql` dosyasına bakınız. Sorgular
Oracle tabanlı bir HBYS için yazılmıştır ve kendi şemanıza uyarlanmalıdır.

---

## Yöntem özeti

Yazılım her kalem için günlük stok serisini yeniden kurar ve aşağıdaki kurallarla
bulunurluk hesaplar.

**Nominal bulunurluk.** Gün başı bakiye pozitif, gün sonu bakiye pozitif, o gün çıkış
var ya da o gün giriş varsa gün stokta sayılır.

**Negatif bakiye düzeltmesi.** Bakiyenin kesintisiz negatif kaldığı dönemde çıkış
yapılmışsa ilaç fiziken raftadır; ilk çıkıştan son çıkışa kadarki günler stokta
sayılır. Kayıt gecikmesinden doğan yapay stoksuzluğu giderir.

**Donmuş bakiye kuralı.** Kalan bakiye, kalemin tipik günlük çıkışının altındaysa ve
bu durum hiç hareket görmeden en az üç gün sürüyorsa bakiye dağıtılamaz kabul edilir.
Eşik kalemin kendi hareketine göre ölçeklenir.

**Molekül düzeyi analiz.** Aynı etken maddenin farklı doz ve formları gruplanarak,
kalem düzeyindeki değişimin gerçek talep değişimi mi yoksa sunum kayması mı olduğu
ayırt edilir.

Tüm karar eşikleri `belgeler/karar_esikleri.md` dosyasında listelenmiştir ve kodun
başındaki sabitlerden değiştirilebilir. Uyarı metinleri belirlenimci kurallardan
üretilir; aynı veri her çalıştırmada aynı çıktıyı verir.

---

## Çıktı

Üretilen çalışma kitabı şu sayfaları içerir:

- **BIRLESIK** — Dönem karşılaştırmalı ana tablo, değişim yüzdeleri, durum
  sınıflandırması, uyarılar ve molekül analizi
- **MOLEKUL_OZET** — Molekül bazında etkin madde toplamları ve değerlendirme
- **<yıl>_AYLIK** — Her dönem için ay ay çıkış ve bulunurluk günü
- **BILGI** — Kapsam, kural tanımları ve veri kalitesi uyarıları

Tek dönem yüklenirse yalnızca o döneme ait aylık tablo üretilir.

---

## Doğrulama

`kod/dogrulama.py`, üç tahminciyi (naif, nominal düzeltme, fiilî düzeltme) iki
katmanlı kayan köken tasarımıyla karşılaştırır. `kod/sekil.py` makaledeki şekli
üretir. Her iki betik `numpy`, `scipy` ve `matplotlib` gerektirir.

---

## Veri hakkında

Depodaki örnek veri **sentetiktir**; gerçek kurum verisi içermez. `TOHUM` sabiti
nedeniyle yeniden üretilebilirdir. Çalışmanın dayandığı ham hareket kayıtları kurumsal
veri niteliğinde olduğundan paylaşılmamaktadır.

---

## Atıf

Bu yazılımı ya da yöntemi kullanıyorsanız lütfen atıf veriniz:

> Dağdeviren YK, Kara FM. Hastane eczanesinde ihale miktarının belirlenmesinde
> bulunurluğa göre düzeltilmiş talep tahmini: Rutin HBYS verisiyle çok dönemli
> otomatik analiz ve geriye dönük doğrulama. *(Dergi bilgisi eklenecek.)*

Yazılımın kendisi için `CITATION.cff` dosyasına ya da Zenodo DOI'sine bakınız.

---

## Lisans

MIT Lisansı — bkz. [LICENSE](LICENSE). Serbestçe kullanılabilir, değiştirilebilir ve
dağıtılabilir.

---

## Sorumluluk reddi

Bu yazılım bir karar **destek** aracıdır. Ürettiği miktarlar, klinik ve idari
değerlendirmenin yerine geçmez. İhale kararlarının sorumluluğu kullanan kuruma aittir.
Üretilen uyarılar belirlenimci kurallara dayanır ve yorum gerektirir.

---

<a name="english"></a>

# English

**Availability-Adjusted Pharmaceutical Demand Estimation for Hospital Pharmacy Tenders**

Open-source tooling that estimates pharmaceutical demand from routine hospital
information system (HIS) stock movement data, correcting for the days a product was
actually out of stock.

Tender quantities are commonly computed by dividing past consumption by calendar time.
Because consumption is recorded only while stock is available, stockout days enter the
record as zero demand and true demand is systematically underestimated. The correction
divides consumption by the number of days the item was **effectively in stock**.

The method was developed on three periods of data from a university hospital pharmacy
(398,401 movement rows, 875 active substances) and validated with a two-fold
rolling-origin design.

**Requirements:** `openpyxl`, `xlrd` (analysis); `numpy`, `scipy`, `matplotlib`
(validation and figures). No pandas, no machine learning components.

**Quick start:**

```bash
pip install openpyxl xlrd
cd ornek_veri && python sentetik_veri_uret.py && cd ..
python kod/ihale_analiz.py ornek_veri/ornek_hareket_*.xlsx
```

Running `python kod/ihale_analiz.py` without arguments opens a desktop interface.

**Minimum data set:** date, active substance code, active substance name, opening
balance, receipts, issues, closing balance. Issues must reflect quantities dispensed
from the pharmacy store to wards, not patient-level charges.

Sample data in this repository is **synthetic**. Source code is released under the MIT
License.
