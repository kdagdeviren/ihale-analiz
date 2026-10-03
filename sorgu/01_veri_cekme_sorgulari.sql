--------------------------------------------------------------------------------
--  İHALE ANALİZİ İÇİN VERİ ÇEKME SORGULARI — GÖZDEN GEÇİRİLMİŞ SÜRÜM
--------------------------------------------------------------------------------
--  Mevcut sorgularda tespit edilen sorunlar ve düzeltmeleri.
--  Her düzeltme, ilgili satırda "DÜZELTME" notuyla işaretlenmiştir.
--
--  DİKKAT: Bu betik gerçek bir Oracle örneğinde ÇALIŞTIRILMAMIŞTIR.
--  Kendi ortamınızda önce mevcut sorgunun çıktısıyla karşılaştırarak
--  doğrulayın.
--------------------------------------------------------------------------------


--------------------------------------------------------------------------------
--  SORGU 1 — GÜNLÜK HAREKET DÖKÜMÜ  (ana veri kaynağı)
--------------------------------------------------------------------------------
--  Mevcut sorgudaki sorunlar:
--
--  S1. TAKVİM İLE İLAÇ ÇAPRAZLANMIYOR.  En önemli sorun budur.
--      Mevcut sorguda dış birleştirme (Et.TARIH(+) = T.Day_Id) yalnızca
--      TARİH üzerinden kurulmuş; ilaç listesiyle çapraz birleşim yok.
--      Sonuç: bir ilacın hareketsiz geçirdiği günler için satır üretilmez.
--      2026 verisinde 47.140 takvim günü bu yüzden eksikti ve bunların
--      37.953'ünde stok pozitifti. Aşağıdaki sürüm takvim x ilaç çaprazı
--      kurarak her ilaç için her güne satır üretir.
--
--  S2. KALAN_SON, ETKEN_MADDE (AD) İLE BÖLÜMLENİYOR.
--      PARTITION BY et.Etken_Madde yerine ETKEN_MADDE_KODU kullanılmalıdır.
--      Aynı ad farklı kodlara ya da aynı kod yıllar içinde farklı adlara
--      sahip olabilir; adla bölümlemek yürüyen stoğu hatalı birleştirir.
--
--  S3. ANALİTİK FONKSİYONUN ORDER BY'INDA ROWNUM VAR.
--      Order By t.Day_Id, Rownum, Nvl(Et.Kalan,0) — ROWNUM dış sorguda
--      satırlar üretildikçe atanır ve birleştirme sırasına göre değişir;
--      kümülatif toplamı belirsiz hale getirir. Gruplama sonrası ilaç-gün
--      çifti zaten tekil olduğu için ORDER BY t.gun yeterlidir.
--
--  S4. TARİH KARŞILAŞTIRMASI METİNLE YAPILIYOR.
--      T.Day_Id >= '01012026' ifadesi NLS_DATE_FORMAT ayarına bağımlıdır ve
--      farklı oturumda sessizce bozulur. DATE değişmezi kullanılmalıdır.
--
--  S5. AMBAR_KODU = 7 GÖMÜLÜ.
--      Bu, deponun servise çıkışını verir (hasta bazlı kullanımı değil).
--      Parametre haline getirildi; yayında da açıkça belirtilmelidir.
--------------------------------------------------------------------------------

WITH parametre AS (
    SELECT DATE '2026-01-01' AS bas,          -- DÜZELTME S4
           DATE '2027-01-01' AS bit,
           7                 AS ambar         -- DÜZELTME S5
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
    -- Dönem içinde en az bir hareketi olan etken maddeler
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
           OVER (PARTITION BY il.kod                    -- DÜZELTME S2
                 ORDER BY t.gun                         -- DÜZELTME S3
                 ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW)
                                                        AS kalan_son,
       TO_CHAR(t.gun, 'DD')                             AS ayin_gunu,
       CASE WHEN h.kod IS NULL THEN 0 ELSE 1 END        AS hareket_var
  FROM takvim t
 CROSS JOIN ilac_listesi il                             -- DÜZELTME S1
  LEFT JOIN hareket h
         ON h.gun = t.gun
        AND h.kod = il.kod
 WHERE t.gun >= il.ilk_gun        -- ilacın ilk hareketinden önceki günler atlanır
 ORDER BY il.kod, t.gun;

--  NOT: Bu sürüm daha fazla satır üretir (yaklaşık 875 ilaç x 365 gün).
--       Excel'e aktarırken MUTLAKA .xlsx kullanın; eski .xls biçiminde
--       65.536 satır sınırı vardır ve veri sessizce birden fazla sayfaya
--       bölünür (2024 dosyasında bu yaşanmıştı, üç sayfaya bölünmüştü).
--
--  NOT: Bu sürümle hareketsiz günler de dolduğu için, analiz tarafındaki
--       "son bilinen stoğu taşıma" adımı gereksiz hale gelir.


--------------------------------------------------------------------------------
--  SORGU 2 — AYLIK ÖZET  (kontrol amaçlı)
--------------------------------------------------------------------------------
--  Mevcut sorgudaki sorunlar:
--
--  S6. ISLEM_TIPI FİLTRESİ YOK.  SUM(ik.Adet) devir, giriş ve çıkışı birlikte
--      toplar. Yalnızca çıkış isteniyorsa DECODE(ik.islem_tipi,'Ç',...)
--      gerekir. MUTLAKA DOĞRULAYIN: mevcut çıktı bizim hesapladığımızdan
--      %2,2 DÜŞÜK geliyordu; eğer üç tipi birlikte toplasaydı çok daha
--      YÜKSEK olmalıydı. Yani görünüşe göre bu görünüm bağlamında yalnızca
--      çıkış satırları dönüyor, ama bu varsayıma bırakılmamalıdır.
--
--  S7. "AGUN" TAKMA ADI İKİ KEZ KULLANILMIŞ.
--      Hem Ağustos (month=8) hem Aralık (month=12) gün sayısı AGUN olarak
--      adlandırılmış. Excel'e aktarımda sütunlardan biri yeniden adlandırılır
--      ya da üzerine yazılır. Aralık için ARGUN kullanıldı.
--
--  S8. GRUPLAMA ETKEN_MADDE (AD) ÜZERİNDEN.  Kod alanı yorum satırına
--      alınmış. Aynı ada sahip iki kod tek satırda birleşir. Kod eklendi.
--
--  S9. "GÜN" SÜTUNLARININ ANLAMI.
--      COUNT(DISTINCT TRUNC(tarih)) = o ay HAREKET GÖRÜLEN gün sayısıdır,
--      ilacın STOKTA BULUNDUĞU gün sayısı DEĞİLDİR. İkisi karıştırılıp
--      payda olarak kullanılırsa aylık talep olduğundan yüksek hesaplanır.
--      Sütun adları bunu yansıtacak şekilde değiştirildi.
--------------------------------------------------------------------------------

WITH p AS (
    SELECT TRUNC(SYSDATE, 'YYYY')                  AS bas,
           ADD_MONTHS(TRUNC(SYSDATE, 'YYYY'), 12)  AS bit,
           7                                       AS ambar
      FROM dual
),
cikislar AS (
    SELECT i.etken_madde_kodu                  AS kod,
           et.etken_madde                      AS ad,
           EXTRACT(MONTH FROM ik.tarih)        AS ay,
           TRUNC(ik.tarih)                     AS gun,
           DECODE(ik.islem_tipi, 'Ç', ik.adet, 0) AS adet   -- DÜZELTME S6
      FROM v_ilac_cikis_giris_durum ik
      JOIN ilac i                 ON i.ilac_no = ik.ilac_no
      JOIN hastane.etken_madde et ON et.etken_madde_kodu = i.etken_madde_kodu
     CROSS JOIN p
     WHERE ik.ambar_kodu = p.ambar
       AND ik.tarih >= p.bas
       AND ik.tarih <  p.bit
)
SELECT kod                                          AS etken_madde_kodu,  -- DÜZELTME S8
       ad                                           AS etken_madde,
       SUM(CASE WHEN ay = 1 THEN adet ELSE 0 END)   AS ocak,
       COUNT(DISTINCT CASE WHEN ay = 1 AND adet > 0 THEN gun END)  AS ocak_hareketli_gun,
       SUM(CASE WHEN ay = 2 THEN adet ELSE 0 END)   AS subat,
       COUNT(DISTINCT CASE WHEN ay = 2 AND adet > 0 THEN gun END)  AS subat_hareketli_gun,
       SUM(CASE WHEN ay = 3 THEN adet ELSE 0 END)   AS mart,
       COUNT(DISTINCT CASE WHEN ay = 3 AND adet > 0 THEN gun END)  AS mart_hareketli_gun,
       SUM(CASE WHEN ay = 4 THEN adet ELSE 0 END)   AS nisan,
       COUNT(DISTINCT CASE WHEN ay = 4 AND adet > 0 THEN gun END)  AS nisan_hareketli_gun,
       SUM(CASE WHEN ay = 5 THEN adet ELSE 0 END)   AS mayis,
       COUNT(DISTINCT CASE WHEN ay = 5 AND adet > 0 THEN gun END)  AS mayis_hareketli_gun,
       SUM(CASE WHEN ay = 6 THEN adet ELSE 0 END)   AS haziran,
       COUNT(DISTINCT CASE WHEN ay = 6 AND adet > 0 THEN gun END)  AS haziran_hareketli_gun,
       SUM(CASE WHEN ay = 7 THEN adet ELSE 0 END)   AS temmuz,
       COUNT(DISTINCT CASE WHEN ay = 7 AND adet > 0 THEN gun END)  AS temmuz_hareketli_gun,
       SUM(CASE WHEN ay = 8 THEN adet ELSE 0 END)   AS agustos,
       COUNT(DISTINCT CASE WHEN ay = 8 AND adet > 0 THEN gun END)  AS agustos_hareketli_gun,
       SUM(CASE WHEN ay = 9 THEN adet ELSE 0 END)   AS eylul,
       COUNT(DISTINCT CASE WHEN ay = 9 AND adet > 0 THEN gun END)  AS eylul_hareketli_gun,
       SUM(CASE WHEN ay = 10 THEN adet ELSE 0 END)  AS ekim,
       COUNT(DISTINCT CASE WHEN ay = 10 AND adet > 0 THEN gun END) AS ekim_hareketli_gun,
       SUM(CASE WHEN ay = 11 THEN adet ELSE 0 END)  AS kasim,
       COUNT(DISTINCT CASE WHEN ay = 11 AND adet > 0 THEN gun END) AS kasim_hareketli_gun,
       SUM(CASE WHEN ay = 12 THEN adet ELSE 0 END)  AS aralik,
       COUNT(DISTINCT CASE WHEN ay = 12 AND adet > 0 THEN gun END) AS aralik_hareketli_gun  -- DÜZELTME S7
  FROM cikislar
 GROUP BY kod, ad
 ORDER BY NLSSORT(ad, 'NLS_SORT=XTURKISH');


--------------------------------------------------------------------------------
--  DOĞRULAMA SORGULARI  (yayın öncesi mutlaka çalıştırın)
--------------------------------------------------------------------------------

--  D1. KALAN_SON gerçekten yürüyen stok mu?  (0 dönmelidir)
/*
WITH v AS (SELECT kod, gun, kalan_son, giris, cikis,
                  LAG(kalan_son) OVER (PARTITION BY kod ORDER BY gun) AS onceki
             FROM (<SORGU 1>))
SELECT COUNT(*) AS tutarsiz_satir
  FROM v
 WHERE onceki IS NOT NULL
   AND kalan_son <> onceki + giris - cikis;
*/

--  D2. Her ilaç için her gün satır var mı?  (fark 0 olmalıdır)
/*
SELECT COUNT(DISTINCT kod) * COUNT(DISTINCT day_id) - COUNT(*) AS eksik_satir
  FROM (<SORGU 1>);
*/

--  D3. DEVIR yalnızca dönemin ilk gününde mi dolu?
/*
SELECT TO_CHAR(day_id,'MM-DD') AS tarih, COUNT(*) AS satir
  FROM (<SORGU 1>) WHERE devir <> 0
 GROUP BY TO_CHAR(day_id,'MM-DD') ORDER BY 1;
*/

--  D4. Sorgu 2'nin toplamı Sorgu 1'in çıkış toplamına eşit mi?
/*
SELECT (SELECT SUM(cikis) FROM (<SORGU 1>)) AS sorgu1,
       (SELECT SUM(ocak+subat+mart+nisan+mayis+haziran+temmuz
                  +agustos+eylul+ekim+kasim+aralik) FROM (<SORGU 2>)) AS sorgu2
  FROM dual;
*/


--------------------------------------------------------------------------------
--  ASGARİ VERİ KÜMESİ  (yayında "aktarılabilirlik" bölümü için)
--------------------------------------------------------------------------------
--  Yöntemin çalışması için gereken yedi alan:
--
--    gun        DATE      hareket günü
--    kod        NUMBER    etken madde kodu   (ATC kodu da kullanılabilir)
--    ad         VARCHAR2  etken madde adı
--    devir      NUMBER    dönem başı devreden
--    giris      NUMBER    o gün stoğa giren
--    cikis      NUMBER    o gün servise çıkan
--    kalan_son  NUMBER    gün sonu yürüyen stok
--
--  KALAN (günlük net hareket) sütunu hesaplamada KULLANILMAZ.
--  Depo kodu, çıkışın hangi ambardan yapıldığını belirler ve mutlaka
--  raporlanmalıdır: hasta bazlı kullanım kaydı ile depo çıkışı farklı
--  büyüklüklerdir.
--------------------------------------------------------------------------------
