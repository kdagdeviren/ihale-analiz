# Karar Eşikleri

Yazılımın ürettiği uyarılar, istatistiksel bir modelden değil önceden tanımlanmış
eşik karşılaştırmalarından üretilir. Aynı veri her çalıştırmada aynı metni üretir ve
her uyarı tek bir kurala kadar izlenebilir.

Eşikler `kod/ihale_analiz.py` dosyasının başındaki sabitlerden ve ilgili fonksiyonlardan
değiştirilebilir. Donmuş bakiye kuralının iki parametresi duyarlılık analiziyle
sınanmıştır; diğer eşikler ihale planlamasında dikkat çekmesi beklenen durumları
yakalayacak biçimde uzman görüşüyle belirlenmiş olup veriden türetilmemiştir.

**Önemli:** Bu eşiklerin değiştirilmesi hesaplanan tüketim miktarlarını değil,
yalnızca hangi kalemlerin işaretleneceğini etkiler.

## Eşik tablosu

| # | Kural | Eşik | Düzey |
|---|---|---|---|
| **Bulunurluk hesabı** ||||
| 1 | Donmuş bakiye: kalan bakiye / tipik günlük çıkış | < 1,0 | — |
| 2 | Donmuş bakiyenin kesintisiz süresi | ≥ 3 gün | — |
| 3 | Negatif bakiye döneminde çıkış varsa stokta sayma | — | — |
| **Tek dönem uyarıları** ||||
| 4 | Stokta bulunulan gün sayısı | < 30 gün | Yüksek |
| 5 | Stokta bulunulan gün / dönem günü | < %40 | Düşük |
| 6 | Toplam çıkışın tek aydaki payı (≥ 2 ay stoklu ise) | ≥ %70 | Orta |
| 7 | En yoğun ayın hızı / genel ortalama | ≥ 3 kat | Orta |
| 8 | Stokta olduğu hâlde çıkış görülmeyen ay sayısı | ≥ 3 ay | Düşük |
| 9 | Donmuş bakiye gün sayısı | ≥ 15 gün | Orta |
| **Çok dönem uyarıları** ||||
| 10 | Ardışık dönem değişimlerinin tamamı | ≥ +%25 | Orta |
| 11 | Ardışık dönem değişimlerinin tamamı | ≤ −%25 | Orta |
| 12 | Son dönem değişimi (mutlak) | ≥ %60 | Düşük |
| 13 | Bulunurluk oranı düşük olan dönem sayısı (< %60) | ≥ 2 dönem | Orta |
| 14 | Tüm dönemler toplamında donmuş bakiye | ≥ 15 gün | Orta |
| 15 | Tüm dönemler toplamında stokta bulunulan gün | < 60 gün | Yüksek |
| 16 | Değişim yüzdesi (mutlak) ve başlangıç dönemi ortalaması | ≥ %300 ve < 5 | Düşük |
| **Molekül düzeyi (sunum kayması)** ||||
| 17 | Kalem değişimi ≤ −%35 ve molekül değişimi ±%20 içinde | — | Sunum kayması |
| 18 | Kalem değişimi ≥ +%50 ve molekül değişimi ±%20 içinde | — | Devralan sunum |
| 19 | Kalem ve molekül değişimi birlikte ≤ −%35 | — | Gerçek düşüş |
| 20 | Kalem ve molekül değişimi birlikte ≥ +%50 | — | Gerçek artış |
| 21 | Kalem ≤ −%35 ve molekül düşüşü kalemden ≥ 25 puan az | — | Kısmi kayma |
| 22 | Kalem ≥ +%50 ve molekül artışı kalemden ≥ 40 puan az | — | Kısmi kayma |

Bir kalem birden fazla kuralı karşılayabilir; bu durumda uyarılar birikmekte ve
kaleme atanan düzey en yüksek olanı almaktadır.

Eşik değerleri iki kaynaktan belirlenmiştir. Donmuş bakiye kuralının iki parametresi
(1 ve 2 numaralı kurallar) duyarlılık analiziyle sınanmış, sonuçlar Bulgular bölümünde
sunulmuştur. Diğer eşikler, ihale planlamasında dikkat çekmesi beklenen durumları
yakalayacak biçimde uzman görüşüyle belirlenmiş olup veriden türetilmemiştir; bu
eşiklerin değiştirilmesi, hesaplanan tüketim miktarlarını değil yalnızca hangi
kalemlerin işaretleneceğini etkiler.

