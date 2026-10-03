# İlaç Tüketim Analizi — SQL Sorgularının Gözden Geçirilmesi

**Hazırlayan:** Dr. Ecz. Yusuf Kağan Dağdeviren — DEÜ Hastane Eczanesi
**Konu:** Etken madde bazlı ilaç tüketim sorgularında tespit edilen sorunlar ve düzeltme önerileri
**Tarih:** 29.09.2026

---

## Neden bu gözden geçirme yapıldı

Hastane eczanesinin ihale miktarlarını belirlemek amacıyla 2024–2026 dönemine ait
hareket verileri üzerinde çok yıllı bir tüketim analizi yürütülmektedir. Analizin
temel metriği şudur:

```
Aylık Ortalama = (Toplam Çıkış / İlacın fiilen stokta bulunduğu gün sayısı) × 30
```

Paydadaki **"stokta bulunulan gün"** kavramı, yöntemin çekirdeğidir. Bir ilaç yılın
yarısında stoksuz kaldıysa, yıllık toplam çıkış gerçek talebin yarısını gösterir ve
o rakamla ihaleye çıkmak stoksuzluğu bir yıl daha sürdürür.

Bu çalışma kapsamında mevcut iki sorgu incelenmiş, çıktıları ham veriyle
karşılaştırılarak doğrulanmıştır. Aşağıdaki tespitler bu karşılaştırmalara
dayanmaktadır.

> **Not:** Aşağıdaki düzeltilmiş sorgular canlı veritabanında çalıştırılmamıştır.
> Devreye almadan önce mevcut sorguların çıktısıyla karşılaştırılarak doğrulanması
> gerekir. Bölüm 3'te bunun için kontrol sorguları verilmiştir.

---

## 1. GÜNLÜK HAREKET SORGUSU

Bu sorgu, analizin **ana veri kaynağıdır**. Aşağıdaki dört tespitten ilki kritik
önemdedir.

### 1.1 Takvim ile ilaç listesi çaprazlanmıyor — ÖNCELİKLİ

Mevcut sorguda dış birleştirme yalnızca tarih üzerinden kurulmuş:

```sql
And Et.TARIH(+) = T.Day_Id
```

Bu, ilaç listesiyle bir çapraz birleşim içermediği için, **bir ilacın hareketsiz
geçirdiği günler için satır üretilmiyor.**

Ölçülen etki: 2026 Ocak–Temmuz döneminde 47.140 takvim günü çıktıda hiç yer almıyor
ve bu günlerin 37.953'ünde ilacın stoğu pozitif. Yani ilaç raftaydı ama veri bunu
göstermiyor. Analiz tarafında bu boşluklar "son bilinen stoğu taşıma" mantığıyla
elle dolduruluyor; bu, veritabanında çözülmesi gereken bir işi uygulama katmanına
taşımak anlamına geliyor.

**Çözüm:** Takvim ile ilaç listesi `CROSS JOIN` ile çaprazlanmalı, hareketler bu
çarpıma `LEFT JOIN` ile bağlanmalıdır. Düzeltilmiş sorgu Bölüm 4'tedir.

### 1.2 Yürüyen stok, etken madde ADI ile bölümleniyor

```sql
Sum(Nvl(Et.Kalan,0)) Over (PARTITION BY et.Etken_Madde ...)
```

`PARTITION BY` ifadesinde kod yerine ad kullanılmış. 2024–2026 verileri
karşılaştırıldığında, aynı etken madde kodunun yıllar içinde farklı adlarla
kaydedildiği 225 kalem tespit edilmiştir (örnek: "MR KONTRAST 15 ML" →
"GADOBUTROL MR KONTRAST 15 ML"). Adla bölümleme, yürüyen stoğu hatalı biçimde
böler ya da birleştirir.

**Çözüm:** `PARTITION BY et.Etken_Madde_Kodu`

### 1.3 Analitik fonksiyonun ORDER BY'ında ROWNUM var

```sql
Order By t.Day_Id, Rownum, Nvl(Et.Kalan, 0)
```

`ROWNUM`, dış sorguda satırlar üretildikçe atanır ve birleştirme sırasına göre
değişebilir. Kümülatif toplamın sıralamasını belirsiz hale getirir. Ayrıca
`Nvl(Et.Kalan,0)` değerinin sıralama anahtarı olarak kullanılması, tarihe göre
yürüyen bir toplamda anlamsızdır.

İç sorguda zaten `GROUP BY ... TARIH` yapıldığı için ilaç-gün çifti tekildir;
ek sıralama anahtarına gerek yoktur.

**Çözüm:** `ORDER BY t.Day_Id` (tek başına yeterli)

### 1.4 Tarih karşılaştırması metin üzerinden yapılıyor

```sql
And T.Day_Id >= '01012026'
And T.Day_Id < '01012027'
```

Bu ifade oturumun `NLS_DATE_FORMAT` ayarına bağımlıdır. Farklı bir oturumda ya da
farklı bir istemciden çalıştırıldığında sessizce hatalı sonuç verebilir veya hata
fırlatabilir.

**Çözüm:** `DATE '2026-01-01'` biçiminde tarih değişmezi ya da
`TO_DATE('01012026','DDMMYYYY')`

### 1.5 Küçük notlar

- `Ayın_Gunu` takma adında Türkçe karakter var; tırnaksız kullanımda taşınabilirlik
  sorunu çıkarabilir. `AYIN_GUNU` tercih edilmeli.
- `AMBAR_KODU = 7` sorguya gömülü. Bu filtre çıktının **depodan servise çıkan
  miktarı** verdiğini belirler (hasta bazında faturalanan kullanımı değil).
  Bu ayrım yayında raporlanacağı için parametre haline getirilmesi ve
  belgelenmesi gerekiyor.

---

## 2. AYLIK ÖZET SORGUSU

Bu sorgu analizin veri kaynağı **değildir**; çapraz kontrol amacıyla kullanılmaktadır.
Stok kavramı içermediği için (ne `DEVIR`, ne `KALAN_SON`, ne yürüyen bakiye)
bulunurluk hesabı bu sorgudan yapılamaz.

### 2.1 "AGUN" takma adı iki kez kullanılmış — KESİN HATA

```sql
COUNT(DISTINCT ...month = 8 ...) AGUN     -- Ağustos
COUNT(DISTINCT ...month = 12 ...) AGUN    -- Aralık
```

Aynı takma ad iki sütuna verilmiş. Excel'e aktarılan çıktıda sütunlar `AGUN` ve
`AGUN1` olarak gelmiş; yani araç kendi düzeltmesini yapmış. Başka bir istemciye
ya da rapor aracına aktarımda sütunlardan biri üzerine yazılabilir.

**Çözüm:** Aralık için `ARGUN` gibi ayrı bir ad.

### 2.2 ISLEM_TIPI filtresi yok

```sql
SUM(CASE WHEN EXTRACT(MONTH FROM ik.Tarih) = 1 THEN ik.Adet ELSE 0 END) OCAK
```

`ik.Adet` hangi işlem tipine ait olduğu belirtilmeden toplanıyor. Görünümde
`ISLEM_TIPI` alanı D (devir), G (giriş) ve Ç (çıkış) değerlerini alıyor — nitekim
günlük sorguda `DECODE` ile bu üçü ayrıştırılıyor.

Çıktı üzerinde yapılan kontrol:

| Ölçüm | Değer |
|---|---|
| Aylık rapor Ocak–Temmuz toplamı | 24.816.980 |
| Günlük veride ÇIKIŞ toplamı | 25.378.076 |
| Günlük veride DEVİR+GİRİŞ+ÇIKIŞ | 52.628.306 |

Rapor, üç tipin toplamına değil yalnızca çıkışa yakın bir değer üretiyor. Yani
pratikte doğru çalışıyor görünüyor — ancak bu, **belgelenmemiş bir görünüm
davranışına dayanıyor.** Sorguya bakan biri hangi hareketin toplandığını
göremiyor.

**İstenen:** Bu davranışın nedeni açıklanmalı (görünüm bu bağlamda neden yalnızca
çıkış döndürüyor?) ve niyet sorguda açık yazılmalı:
`DECODE(ik.islem_tipi, 'Ç', ik.adet, 0)`

Sonuç değişmese bile, yöntem yayına konu olacağı için sorgunun kendi kendini
açıklaması gerekiyor.

### 2.3 Gruplama etken madde ADI üzerinden, kod yorum satırında

```sql
-- ET.ETKEN_MADDE_KODU,
GROUP BY ET.ETKEN_MADDE
```

Aynı ada sahip iki farklı kod tek satırda birleşir. Çok yıllı karşılaştırma
yapıldığı için kod alanı zorunludur.

**Çözüm:** Kod satırlarındaki yorum işaretleri kaldırılmalı, `GROUP BY` ifadesine
kod eklenmeli.

### 2.4 "Gün" sütunlarının anlamı — kavramsal, hata değil

```sql
COUNT(DISTINCT CASE WHEN ... THEN TRUNC(ik.Tarih) END) OGUN
```

Bu ifade **o ay çıkış yapılan gün sayısını** verir. Ham veriyle karşılaştırıldığında
%90 oranında "çıkış yapılan gün" tanımıyla örtüşüyor.

Bu, ilacın **stokta bulunduğu gün sayısı değildir.** Örnek: Adenosin 20 mg ampulde
Ocak ayında 11 gün çıkış yapılmış, ancak ilaç 31 günün tamamında stoktaydı.

İki tanım payda olarak kullanıldığında bambaşka sonuç verir. Aylık talep hesabında
"çıkış yapılan gün" kullanılırsa talep olduğundan yüksek çıkar (bu örnekte aylık 64
yerine 175).

**İstenen:** Sütun adları anlamı yansıtacak biçimde değiştirilmeli
(`OCAK_HAREKETLI_GUN` vb.), böylece raporu kullanan kişi yanlış yorumlamaz.

### 2.5 Eski usul birleştirme

`FROM a, b, c WHERE ...` biçimindeki virgüllü birleştirmeler ANSI `JOIN`
söz dizimine çevrilirse okunabilirlik ve bakım kolaylığı artar.

---

## 3. DEVREYE ALMADAN ÖNCE ÇALIŞTIRILACAK KONTROL SORGULARI

Düzeltilmiş sorgular devreye alınmadan önce aşağıdaki dört kontrolün çalıştırılması
ve sonuçlarının paylaşılması rica olunur.

```sql
-- K1. KALAN_SON gerçekten yürüyen stok mu?  (0 dönmelidir)
WITH v AS (
    SELECT kod, gun, kalan_son, giris, cikis,
           LAG(kalan_son) OVER (PARTITION BY kod ORDER BY gun) AS onceki
      FROM (<DÜZELTİLMİŞ GÜNLÜK SORGU>)
)
SELECT COUNT(*) AS tutarsiz_satir
  FROM v
 WHERE onceki IS NOT NULL
   AND kalan_son <> onceki + giris - cikis;

-- K2. Her ilaç için her gün satır üretiliyor mu?  (0 dönmelidir)
SELECT COUNT(DISTINCT kod) * COUNT(DISTINCT day_id) - COUNT(*) AS eksik_satir
  FROM (<DÜZELTİLMİŞ GÜNLÜK SORGU>);

-- K3. DEVIR yalnızca dönemin ilk gününde mi dolu?
SELECT TO_CHAR(day_id,'MM-DD') AS tarih, COUNT(*) AS satir
  FROM (<DÜZELTİLMİŞ GÜNLÜK SORGU>)
 WHERE devir <> 0
 GROUP BY TO_CHAR(day_id,'MM-DD')
 ORDER BY 1;

-- K4. İki sorgunun çıkış toplamları tutuyor mu?
SELECT (SELECT SUM(cikis) FROM (<DÜZELTİLMİŞ GÜNLÜK SORGU>))   AS gunluk_toplam,
       (SELECT SUM(ocak+subat+mart+nisan+mayis+haziran+temmuz
                  +agustos+eylul+ekim+kasim+aralik)
          FROM (<DÜZELTİLMİŞ AYLIK SORGU>))                    AS aylik_toplam
  FROM dual;
```

---

## 4. DÜZELTİLMİŞ SORGULAR

### 4.1 Günlük hareket dökümü

```sql
WITH parametre AS (
    SELECT DATE '2026-01-01' AS bas,
           DATE '2027-01-01' AS bit,
           7                 AS ambar
      FROM dual
),
hareket AS (
    SELECT TRUNC(vm.tarih)                              AS gun,
           i.etken_madde_kodu                           AS kod,
           SUM(DECODE(vm.islem_tipi, 'D', vm.adet, 0))  AS devir,
           SUM(DECODE(vm.islem_tipi, 'G', vm.adet, 0))  AS giris,
           SUM(DECODE(vm.islem_tipi, 'Ç', vm.adet, 0))  AS cikis
      FROM hastane.v_ilac_cikis_giris_durum vm
      JOIN ilac i ON i.ilac_no = vm.ilac_no
     CROSS JOIN parametre p
     WHERE vm.ambar_kodu = p.ambar
       AND vm.tarih >= p.bas
       AND vm.tarih <  p.bit
     GROUP BY TRUNC(vm.tarih), i.etken_madde_kodu
),
ilac_listesi AS (
    SELECT h.kod, MIN(h.gun) AS ilk_gun, et.etken_madde AS ad
      FROM hareket h
      JOIN hastane.etken_madde et ON et.etken_madde_kodu = h.kod
     GROUP BY h.kod, et.etken_madde
),
takvim AS (
    SELECT (SELECT bas FROM parametre) + LEVEL - 1 AS gun
      FROM dual
   CONNECT BY LEVEL <= (SELECT bit - bas FROM parametre)
)
SELECT t.gun                                            AS day_id,
       il.kod                                           AS etken_madde_kodu,
       il.ad                                            AS etken_madde,
       NVL(h.devir, 0)                                  AS devir,
       NVL(h.giris, 0)                                  AS giris,
       NVL(h.cikis, 0)                                  AS cikis,
       NVL(h.devir,0) + NVL(h.giris,0) - NVL(h.cikis,0) AS kalan,
       SUM(NVL(h.devir,0) + NVL(h.giris,0) - NVL(h.cikis,0))
           OVER (PARTITION BY il.kod
                 ORDER BY t.gun
                 ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW)
                                                        AS kalan_son,
       TO_CHAR(t.gun, 'DD')                             AS ayin_gunu,
       CASE WHEN h.kod IS NULL THEN 0 ELSE 1 END        AS hareket_var
  FROM takvim t
 CROSS JOIN ilac_listesi il
  LEFT JOIN hareket h
         ON h.gun = t.gun
        AND h.kod = il.kod
 WHERE t.gun >= il.ilk_gun
 ORDER BY il.kod, t.gun;
```

**Satır sayısı uyarısı:** Bu sürüm hareketsiz günleri de ürettiği için satır sayısı
yaklaşık iki katına çıkar (≈ 875 ilaç × 365 gün). Excel'e aktarımda **mutlaka
`.xlsx` biçimi** kullanılmalıdır; eski `.xls` biçiminde 65.536 satır sınırı vardır
ve veri sessizce birden fazla sayfaya bölünür. 2024 dosyasında bu yaşanmış, veri
üç sayfaya bölünmüştü ve bölünme ilaçların ortasından geçmişti.

### 4.2 Aylık özet (çapraz kontrol)

```sql
WITH p AS (
    SELECT TRUNC(SYSDATE, 'YYYY')                  AS bas,
           ADD_MONTHS(TRUNC(SYSDATE, 'YYYY'), 12)  AS bit,
           7                                       AS ambar
      FROM dual
),
cikislar AS (
    SELECT i.etken_madde_kodu                         AS kod,
           et.etken_madde                             AS ad,
           EXTRACT(MONTH FROM ik.tarih)               AS ay,
           TRUNC(ik.tarih)                            AS gun,
           DECODE(ik.islem_tipi, 'Ç', ik.adet, 0)     AS adet
      FROM v_ilac_cikis_giris_durum ik
      JOIN ilac i                 ON i.ilac_no = ik.ilac_no
      JOIN hastane.etken_madde et ON et.etken_madde_kodu = i.etken_madde_kodu
     CROSS JOIN p
     WHERE ik.ambar_kodu = p.ambar
       AND ik.tarih >= p.bas
       AND ik.tarih <  p.bit
)
SELECT kod AS etken_madde_kodu,
       ad  AS etken_madde,
       SUM(CASE WHEN ay = 1 THEN adet ELSE 0 END)  AS ocak,
       COUNT(DISTINCT CASE WHEN ay = 1 AND adet > 0 THEN gun END)  AS ocak_hareketli_gun,
       SUM(CASE WHEN ay = 2 THEN adet ELSE 0 END)  AS subat,
       COUNT(DISTINCT CASE WHEN ay = 2 AND adet > 0 THEN gun END)  AS subat_hareketli_gun,
       SUM(CASE WHEN ay = 3 THEN adet ELSE 0 END)  AS mart,
       COUNT(DISTINCT CASE WHEN ay = 3 AND adet > 0 THEN gun END)  AS mart_hareketli_gun,
       SUM(CASE WHEN ay = 4 THEN adet ELSE 0 END)  AS nisan,
       COUNT(DISTINCT CASE WHEN ay = 4 AND adet > 0 THEN gun END)  AS nisan_hareketli_gun,
       SUM(CASE WHEN ay = 5 THEN adet ELSE 0 END)  AS mayis,
       COUNT(DISTINCT CASE WHEN ay = 5 AND adet > 0 THEN gun END)  AS mayis_hareketli_gun,
       SUM(CASE WHEN ay = 6 THEN adet ELSE 0 END)  AS haziran,
       COUNT(DISTINCT CASE WHEN ay = 6 AND adet > 0 THEN gun END)  AS haziran_hareketli_gun,
       SUM(CASE WHEN ay = 7 THEN adet ELSE 0 END)  AS temmuz,
       COUNT(DISTINCT CASE WHEN ay = 7 AND adet > 0 THEN gun END)  AS temmuz_hareketli_gun,
       SUM(CASE WHEN ay = 8 THEN adet ELSE 0 END)  AS agustos,
       COUNT(DISTINCT CASE WHEN ay = 8 AND adet > 0 THEN gun END)  AS agustos_hareketli_gun,
       SUM(CASE WHEN ay = 9 THEN adet ELSE 0 END)  AS eylul,
       COUNT(DISTINCT CASE WHEN ay = 9 AND adet > 0 THEN gun END)  AS eylul_hareketli_gun,
       SUM(CASE WHEN ay = 10 THEN adet ELSE 0 END) AS ekim,
       COUNT(DISTINCT CASE WHEN ay = 10 AND adet > 0 THEN gun END) AS ekim_hareketli_gun,
       SUM(CASE WHEN ay = 11 THEN adet ELSE 0 END) AS kasim,
       COUNT(DISTINCT CASE WHEN ay = 11 AND adet > 0 THEN gun END) AS kasim_hareketli_gun,
       SUM(CASE WHEN ay = 12 THEN adet ELSE 0 END) AS aralik,
       COUNT(DISTINCT CASE WHEN ay = 12 AND adet > 0 THEN gun END) AS aralik_hareketli_gun
  FROM cikislar
 GROUP BY kod, ad
 ORDER BY NLSSORT(ad, 'NLS_SORT=XTURKISH');
```

---

## 5. ÖZET — YAPILMASI İSTENENLER

| # | Sorgu | Konu | Öncelik |
|---|---|---|---|
| 1 | Günlük | Takvim × ilaç çapraz birleşimi eklensin | **Yüksek** |
| 2 | Günlük | `PARTITION BY` kod ile yapılsın | Yüksek |
| 3 | Günlük | Analitik `ORDER BY`'dan `ROWNUM` çıkarılsın | Orta |
| 4 | Günlük | Tarih karşılaştırması `DATE` değişmeziyle yapılsın | Orta |
| 5 | Aylık | `AGUN` takma ad çakışması giderilsin | **Yüksek** |
| 6 | Aylık | `ISLEM_TIPI` filtresi açıkça yazılsın / davranış açıklansın | **Yüksek** |
| 7 | Aylık | Gruplama koda taşınsın | Yüksek |
| 8 | Aylık | Gün sütunları `..._HAREKETLI_GUN` olarak adlandırılsın | Orta |
| 9 | Her ikisi | `AMBAR_KODU` parametre haline getirilsin | Düşük |
| 10 | Her ikisi | Bölüm 3'teki kontrol sorguları çalıştırılıp sonuç paylaşılsın | **Yüksek** |

---

## 6. EK BİLGİ — ANALİZİN İHTİYAÇ DUYDUĞU ASGARİ ALANLAR

Yöntemin çalışması için günlük sorgudan yalnızca şu yedi alan gereklidir:

| Alan | Tip | Açıklama |
|---|---|---|
| `gun` | DATE | Hareket günü |
| `etken_madde_kodu` | NUMBER | Etken madde kodu |
| `etken_madde` | VARCHAR2 | Etken madde adı |
| `devir` | NUMBER | Dönem başı devreden |
| `giris` | NUMBER | O gün stoğa giren |
| `cikis` | NUMBER | O gün servise çıkan |
| `kalan_son` | NUMBER | Gün sonu yürüyen stok |

`KALAN` (günlük net hareket) sütunu hesaplamada kullanılmamaktadır; isteğe bağlıdır.
