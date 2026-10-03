# -*- coding: utf-8 -*-
"""
SENTETİK ÖRNEK VERİ ÜRETİCİ
============================
Depodaki örnek veri setini üretir. Üretilen dosyalar GERÇEK KURUM VERİSİ DEĞİLDİR;
yalnızca yazılımın denenmesi ve veri biçiminin gösterilmesi amacıyla rastgele
üretilmiştir.

Üretilen dosyalar gerçek HBYS çıktısının sütun yapısını birebir taklit eder:
    DAY_ID, ETKEN_MADDE_KODU, ETKEN_MADDE, DEVIR, GIRIS, CIKIS, KALAN, KALAN_SON, AYIN_GUNU

Kullanım:
    python sentetik_veri_uret.py
"""

import calendar
import datetime as dt
import random

from openpyxl import Workbook

TOHUM = 20261002          # yeniden üretilebilirlik için sabit tohum
random.seed(TOHUM)

# Sentetik ilaç listesi: (ad, tipik günlük çıkış, stoksuzluğa yatkınlık)
ILACLAR = [
    ("ADENOZIN 20 MG AMPUL",                     2,   0.05),
    ("ADRENALIN 1 MG AMPUL",                   180,   0.02),
    ("AMOKSISILIN + KLAVULANIK ASIT 1000 MG TABLET", 320, 0.10),
    ("AMOKSISILIN + KLAVULANIK ASIT 625 MG TABLET",  210, 0.10),
    ("ASETILSALISILIK ASIT 100 MG TABLET",      450,  0.04),
    ("ASETILSALISILIK ASIT 300 MG TABLET",       90,  0.06),
    ("BUPIVAKAIN %0,5 20 ML FLAKON",             50,  0.45),   # kronik stoksuz örnek
    ("DILTIAZEM 60 MG TABLET",                   40,  0.08),
    ("DILTIAZEM 90 MG TABLET",                   25,  0.55),   # sunum kayması örneği
    ("DILTIAZEM 120 MG TABLET",                  15,  0.60),   # sunum kayması örneği
    ("ENOKSAPARIN 4000 ANTI-XA IU ENJEKTOR",    260,  0.07),
    ("FENTANIL 0,05 MG/ML 10 ML AMPUL",          75,  0.12),
    ("KT FLOROURASIL 5000 MG FLAKON",         28000,  0.05),   # mg cinsinden
    ("KT FLOROURASIL 1000 MG FLAKON",          4200,  0.35),
    ("KT GEMSITABIN 2000 MG FLAKON",           6500,  0.09),
    ("METILPREDNIZOLON 20 MG AMPUL",             85,  0.06),
    ("METILPREDNIZOLON 250 MG AMPUL",            60,  0.15),
    ("METILPREDNIZOLON 16 MG TABLET",           140,  0.05),
    ("PANTOPRAZOL 40 MG FLAKON",                380,  0.06),
    ("PARASETAMOL 10 MG/ML 100 ML FLAKON",      520,  0.03),
    ("PIPERASILIN + TAZOBAKTAM 4,5 G FLAKON",   240,  0.25),
    ("SEFTRIAKSON 1000 MG FLAKON",              430,  0.08),
    ("SER. %0,9 SODYUM KLORUR 100 ML TORBA",   1100,  0.03),
    ("SER. %0,9 SODYUM KLORUR 1000 ML TORBA",   640,  0.04),
    ("TRAMADOL 100 MG/2 ML AMPUL",              120,  0.10),
]

DONEMLER = [
    (2024, 1, 1, 12, 31),      # tam yıl
    (2025, 1, 1, 12, 31),      # tam yıl
    (2026, 1, 1,  7, 31),      # kısmi yıl — çok dönemli analizi göstermek için
]

BUYUME = {2024: 1.00, 2025: 1.06, 2026: 1.12}     # yıllık hacim artışı


def gun_carpani(tarih):
    """Hafta sonu düşüşü ve rastgele dalgalanma."""
    c = 0.22 if tarih.weekday() >= 5 else 1.0
    return c * random.uniform(0.6, 1.4)


def donem_uret(yil, bas_ay, bas_gun, bit_ay, bit_gun):
    satirlar = []
    basla = dt.date(yil, bas_ay, bas_gun)
    bitir = dt.date(yil, bit_ay, bit_gun)

    for kod_no, (ad, tipik, risk) in enumerate(ILACLAR, start=1001):
        gunluk = tipik * BUYUME[yil]
        # açılış stoğu
        stok = round(gunluk * random.uniform(8, 25))
        devir_yazildi = False
        # stoksuz dönemler: rastgele aralıklar
        stoksuz_gunler = set()
        t = basla
        while t <= bitir:
            if random.random() < risk / 60:
                uzunluk = random.randint(5, 45)
                for i in range(uzunluk):
                    g = t + dt.timedelta(days=i)
                    if g <= bitir:
                        stoksuz_gunler.add(g)
                t += dt.timedelta(days=uzunluk)
            t += dt.timedelta(days=1)

        t = basla
        while t <= bitir:
            devir = stok if not devir_yazildi else 0
            devir_yazildi = True
            giris = 0
            cikis = 0

            if t in stoksuz_gunler:
                stok = 0
            else:
                if stok <= gunluk * 2 and random.random() < 0.35:
                    giris = round(gunluk * random.uniform(10, 30))
                talep = gunluk * gun_carpani(t)
                cikis = min(round(talep), stok + giris)
                if cikis < 0:
                    cikis = 0
                stok = stok + giris - cikis
                # HBYS kayıt gecikmesi: ara sıra negatif bakiye
                if random.random() < 0.004:
                    cikis += round(gunluk * random.uniform(1, 3))
                    stok = stok + giris - cikis

            if devir == 0 and giris == 0 and cikis == 0 and random.random() < 0.55:
                t += dt.timedelta(days=1)       # hareketsiz gün: satır yazılmaz
                continue

            satirlar.append([
                dt.datetime(t.year, t.month, t.day),
                kod_no, ad,
                devir, giris, cikis,
                devir + giris - cikis,
                stok,
                f"{t.day:02d}",
            ])
            t += dt.timedelta(days=1)
    return satirlar


def yaz(yil, satirlar):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet 1"
    ws.append(["DAY_ID", "ETKEN_MADDE_KODU", "ETKEN_MADDE", "DEVIR", "GIRIS",
               "CIKIS", "KALAN", "KALAN_SON", "AYIN_GUNU"])
    for r in sorted(satirlar, key=lambda x: (x[0], x[1])):
        ws.append(r)
    ad = f"ornek_hareket_{yil}.xlsx"
    wb.save(ad)
    print(f"  {ad}: {len(satirlar):,} satır")


if __name__ == "__main__":
    print("Sentetik örnek veri üretiliyor (gerçek kurum verisi DEĞİLDİR)…")
    for yil, ba, bg, bia, big in DONEMLER:
        yaz(yil, donem_uret(yil, ba, bg, bia, big))
    print("Tamamlandı.")
