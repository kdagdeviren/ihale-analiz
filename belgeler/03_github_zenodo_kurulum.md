# GitHub ve Zenodo Kurulumu — Adım Adım

Bu belge, deponun GitHub'a yüklenmesi ve Zenodo üzerinden kalıcı bir DOI alınması
için izlenecek adımları içerir. Toplam süre yaklaşık 30 dakikadır.

---

## 1. Depoyu GitHub'a yükle

### 1.1 Yeni depo oluştur

GitHub'da **New repository**:

- **Repository name:** `ihale-analiz`
- **Description:** Hastane eczanesinde bulunurluğa göre düzeltilmiş ilaç talebi tahmini
- **Public** (Zenodo yalnızca açık depoları arşivler)
- **Add a README file** işaretini **KALDIRIN** (zaten var)
- **Add .gitignore** → None (zaten var)
- **Choose a license** → None (zaten var)

### 1.2 Dosyaları yükle

En kolay yol tarayıcıdan: depo sayfasında **uploading an existing file** bağlantısına
tıklayıp tüm klasörü sürükleyin.

Komut satırını tercih ederseniz:

```bash
cd ihale-analiz
git init
git add .
git commit -m "İlk sürüm: analiz motoru, doğrulama, sorgular, sentetik örnek veri"
git branch -M main
git remote add origin https://github.com/<kullanıcı-adı>/ihale-analiz.git
git push -u origin main
```

### 1.3 Yüklemeden önce son kontrol

> **Kritik.** Gerçek kurum verisinin yüklenmediğinden emin olun. `.gitignore` dosyası
> `*_Etken_Madde_Analiz.*`, `Ilac_Etken_analiz.*` ve `SONUÇ*.xlsm` kalıplarını
> engeller, ancak farklı adlandırılmış bir dosya yakalanmaz.

```bash
git status --short        # yüklenecek dosyaları listeler, gözle kontrol edin
```

Yüklenmesi gereken dosyalar:

```
README.md  LICENSE  CITATION.cff  requirements.txt  .gitignore  .zenodo.json
kod/        ihale_analiz.py  dogrulama.py  sekil.py  exe_olusturucu.py
sorgu/      01_veri_cekme_sorgulari.sql
ornek_veri/ sentetik_veri_uret.py
belgeler/   01_karar_esikleri.md  02_sql_notlari.md  03_github_zenodo_kurulum.md
```

Örnek `.xlsx` dosyaları `.gitignore` kapsamındadır; kullanıcı üreteciyle kendisi
oluşturur. Depoyu daha kullanışlı kılmak için yüklemek isterseniz `.gitignore`
içindeki ilgili satırı kaldırın.

---

## 2. Zenodo bağlantısını kur

### 2.1 Hesap ve yetkilendirme

1. [zenodo.org](https://zenodo.org) adresinde **Log in with GitHub** ile giriş yapın.
2. Sağ üstten **GitHub** → gelen listede `ihale-analiz` deposunu bulun.
3. Depo yanındaki anahtarı **ON** konumuna getirin.

> Anahtarı açmak yalnızca bundan **sonraki** sürümleri arşivler. Bu nedenle sürüm
> yayımlama adımı anahtar açıldıktan sonra yapılmalıdır.

### 2.2 ORCID bağlantısı

Zenodo profilinizde ORCID numaranızı bağlayın; böylece arşivlenen yazılım ORCID
kaydınıza otomatik düşer.

---

## 3. Sürüm yayımla

GitHub deposunda **Releases** → **Create a new release**:

- **Tag:** `v1.0.0`
- **Release title:** `v1.0.0 — İlk yayın`
- **Description:**

```
Makale ile birlikte yayımlanan ilk sürüm.

İçerik:
- Analiz motoru (bulunurluk düzeltmesi, negatif bakiye düzeltmesi, donmuş bakiye
  kuralı, molekül düzeyi sunum kayması analizi)
- İki katmanlı kayan köken doğrulama betiği
- HBYS veri çekme sorguları ve doğrulama sorguları
- Sentetik örnek veri üreteci
- Karar eşikleri belgesi
```

**Publish release** düğmesine basın. Zenodo birkaç dakika içinde arşivleyip DOI atar.

---

## 4. DOI'yi yerleştir

Zenodo deponuzun sayfasında iki DOI görürsünüz:

| DOI türü | Kullanım |
|---|---|
| **Concept DOI** | Tüm sürümleri temsil eder — **makalede bunu kullanın** |
| Sürüm DOI'si | Yalnızca v1.0.0'ı gösterir |

Concept DOI'yi şu üç yere yazın:

1. **README.md** — en üstteki rozet satırındaki `XX.XXXX/zenodo.XXXXXXX` kısmını
   değiştirin. Rozet kodunu Zenodo sayfasındaki **DOI badge** bölümünden
   kopyalayabilirsiniz.
2. **CITATION.cff** — `repository-code` satırının altına ekleyin:
   ```yaml
   doi: 10.5281/zenodo.XXXXXXX
   ```
3. **Makale** — "Veri ve Kod Erişilebilirliği" bölümündeki boşluğu doldurun:

   > Çalışmada kullanılan analiz kodu ve masaüstü uygulaması, veri çekme
   > sorgularıyla birlikte açık erişimle paylaşılmıştır:
   > https://github.com/<kullanıcı-adı>/ihale-analiz
   > (arşiv: https://doi.org/10.5281/zenodo.XXXXXXX).

Değişiklikleri kaydettikten sonra isterseniz `v1.0.1` olarak yeni bir sürüm
yayımlayabilirsiniz; Concept DOI değişmez.

---

## 5. Makale gönderiminden sonra

Makale kabul edilince `CITATION.cff` dosyasına makale künyesini ekleyin:

```yaml
preferred-citation:
  type: article
  title: >-
    Hastane eczanesinde ihale miktarının belirlenmesinde bulunurluğa göre
    düzeltilmiş talep tahmini
  authors:
    - given-names: Yusuf Kağan
      family-names: Dağdeviren
    - given-names: Fatih Mahmut
      family-names: Kara
  journal: Hacettepe Sağlık İdaresi Dergisi
  year: 2026
  volume:
  issue:
  start:
  end:
  doi:
```

README'nin **Atıf** bölümünü de güncelleyin.

---

## Kontrol listesi

- [ ] Depo **public** olarak oluşturuldu
- [ ] `git status` ile gerçek veri dosyası olmadığı doğrulandı
- [ ] Dosyalar yüklendi, README doğru görüntüleniyor
- [ ] Zenodo'da depo anahtarı açıldı
- [ ] Zenodo profiline ORCID bağlandı
- [ ] `v1.0.0` sürümü yayımlandı
- [ ] Zenodo DOI alındı
- [ ] Concept DOI README, CITATION.cff ve makaleye yazıldı
- [ ] Depo bağlantısı ikinci yazarla paylaşıldı
