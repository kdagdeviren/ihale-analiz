# -*- coding: utf-8 -*-
"""
KAYAN KÖKEN DOĞRULAMASI (rolling-origin validation)
====================================================
Zaman sırası korunarak iki katman kurulur:

  Katman 1 : eğitim 2024            → test 2025 (334 gün)
  Katman 2 : eğitim 2024 + 2025     → test 2026 (212 gün)

Her katmanda test kümesi, hedef değişkenin sansürlenmemiş olmasını güvence altına
almak için test döneminde fiilen kesintisiz stokta kalan kalemlerle sınırlandırılır.

Üç tahminci:
  A) Naif            : toplam çıkış / takvim günü × test günü
  B) Nominal düzeltme: toplam çıkış / nominal stokta olunan gün × test günü
  C) Fiilî düzeltme  : toplam çıkış / fiilî stokta olunan gün × test günü
"""

import importlib.util
import json
import os
import numpy as np
from scipy import stats

KLASOR = "/home/claude/exe"
DOSYA = {"2024": "2024_Etken_Madde_Analiz.xlsx",
         "2025": "2025_Etken_Madde_Analiz.xlsx",
         "2026": "Ilac_Etken_analiz.xlsx"}


def motor(esik):
    s = importlib.util.spec_from_file_location(f"m{esik}", os.path.join(KLASOR, "app2.py"))
    m = importlib.util.module_from_spec(s)
    s.loader.exec_module(m)
    m.FIILEN_STOKSUZ_ESIK = esik
    return m


def hesapla_tum():
    m_f, m_n = motor(1.0), motor(0.0)
    fii, nom, don = {}, {}, {}
    for yil, dosya in DOSYA.items():
        u = []
        s = m_f.yil_isle(os.path.join(KLASOR, dosya), lambda *a: None, u)
        fii[yil] = {r["KOD"]: r for r in s["satirlar"]}
        don[yil] = s["donem_gun"]
        s2 = m_n.yil_isle(os.path.join(KLASOR, dosya), lambda *a: None, [])
        nom[yil] = {r["KOD"]: r for r in s2["satirlar"]}
    return fii, nom, don


def katman(fii, nom, don, egitim, test):
    """Bir katmanın tahminlerini üretir."""
    t_gun = don[test]
    eg_gun = sum(don[y] for y in egitim)

    # test kümesi: test döneminde fiilen kesintisiz stokta kalanlar
    kume = [k for k, r in fii[test].items()
            if r["toplam"] > 0 and r["toplam_gun"] >= 0.95 * t_gun and r["fiilen"] == 0]
    # eğitim döneminde hareketi olanlar
    kume = [k for k in kume
            if all(k in fii[y] for y in egitim)
            and sum(fii[y][k]["toplam"] for y in egitim) > 0]

    kayit = []
    for k in kume:
        toplam = sum(fii[y][k]["toplam"] for y in egitim)
        nom_gun = sum(nom[y][k]["toplam_gun"] for y in egitim)
        fii_gun = sum(fii[y][k]["toplam_gun"] for y in egitim)
        donmus = sum(fii[y][k]["fiilen"] for y in egitim)
        kayit.append({
            "kod": k, "ad": fii[test][k]["AD"], "katman": f"{'+'.join(egitim)}→{test}",
            "gercek": fii[test][k]["toplam"],
            "A": toplam / eg_gun * t_gun,
            "B": (toplam / nom_gun * t_gun) if nom_gun > 0 else 0.0,
            "C": (toplam / fii_gun * t_gun) if fii_gun > 0 else 0.0,
            "donmus": donmus,
            "sansurlu": nom_gun < 0.95 * eg_gun or donmus > 0,
        })
    return kayit


# ----------------------------- ölçütler -------------------------------------
def olc(veri, yontem):
    y = np.array([r["gercek"] for r in veri], dtype=float)
    p = np.array([r[yontem] for r in veri], dtype=float)
    m = (p > 0) & (y > 0)
    L = np.log(p[m] / y[m])
    return {
        "n": len(veri),
        "wmape": np.abs(p - y).sum() / y.sum() * 100,
        "medyan": float(np.exp(np.median(L))),
        "mae_log": float(np.abs(L).mean()),
        "p_yanlilik": float(stats.wilcoxon(L).pvalue) if len(L) > 5 else float("nan"),
        "_L": L,
    }


def tablo(baslik, veri):
    print("=" * 76)
    print(f"{baslik}   (n = {len(veri)})")
    print("=" * 76)
    print(f"{'Yöntem':<22}{'WMAPE %':>10}{'Medyan tah/ger':>17}{'MAE(log)':>11}{'Yanlılık p':>14}")
    print("-" * 76)
    adlar = {"A": "A) Naif", "B": "B) Nominal düzeltme", "C": "C) Fiilî düzeltme"}
    o = {}
    for y in ("A", "B", "C"):
        o[y] = olc(veri, y)
        print(f"{adlar[y]:<22}{o[y]['wmape']:>10.1f}{o[y]['medyan']:>17.3f}"
              f"{o[y]['mae_log']:>11.3f}{o[y]['p_yanlilik']:>14.3g}")
    # yöntemler arası karşılaştırma
    pAC = stats.wilcoxon(np.abs(o["A"]["_L"]), np.abs(o["C"]["_L"])).pvalue
    pBC = stats.wilcoxon(np.abs(o["B"]["_L"]), np.abs(o["C"]["_L"])).pvalue
    print(f"\n  |log| karşılaştırması:  A↔C p = {pAC:.3g}   |   B↔C p = {pBC:.3g}\n")


def main():
    print("Veriler hesaplanıyor…\n")
    fii, nom, don = hesapla_tum()
    for y in DOSYA:
        print(f"  {y}: {don[y]} gün, {len(fii[y])} kalem")

    k1 = katman(fii, nom, don, ["2024"], "2025")
    k2 = katman(fii, nom, don, ["2024", "2025"], "2026")
    print(f"\nKatman 1 (2024→2025): {len(k1)} kalem | "
          f"sansürlü {sum(r['sansurlu'] for r in k1)}")
    print(f"Katman 2 (2024+2025→2026): {len(k2)} kalem | "
          f"sansürlü {sum(r['sansurlu'] for r in k2)}\n")

    for ad, veri in [("KATMAN 1 — eğitim 2024, test 2025", k1),
                     ("KATMAN 2 — eğitim 2024+2025, test 2026", k2)]:
        s = [r for r in veri if r["sansurlu"]]
        t = [r for r in veri if not r["sansurlu"]]
        tablo(f"{ad} — SANSÜRLÜ", s)
        tablo(f"{ad} — KONTROL", t)

    # havuzlanmış
    hepsi = k1 + k2
    tablo("HAVUZLANMIŞ (iki katman) — SANSÜRLÜ",
          [r for r in hepsi if r["sansurlu"]])
    tablo("HAVUZLANMIŞ (iki katman) — KONTROL",
          [r for r in hepsi if not r["sansurlu"]])

    # donmuş bakiye alt kümesi — güç artışı
    dd = [r for r in hepsi if r["donmus"] > 0]
    tablo("HAVUZLANMIŞ — EĞİTİMDE DONMUŞ BAKİYE VAR", dd)

    with open(os.path.join(KLASOR, "dogrulama2_sonuc.json"), "w", encoding="utf-8") as f:
        json.dump(hepsi, f, ensure_ascii=False)
    print(f"Kaydedildi: dogrulama2_sonuc.json ({len(hepsi)} satır)")


if __name__ == "__main__":
    main()
