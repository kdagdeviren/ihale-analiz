# -*- coding: utf-8 -*-
"""
Şekil 1 — Tahmin/gerçek oranının yöntemlere ve alt kümelere göre dağılımı
Dergi için gri tonlamalı, 600 dpi, TIFF + PNG + PDF.
"""
import json
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 9,
    "axes.linewidth": 0.8,
    "xtick.direction": "out",
    "ytick.direction": "out",
})

d = json.load(open("dogrulama2_sonuc.json", encoding="utf-8"))
g = np.array([r["gercek"] for r in d])
A = np.array([r["A"] for r in d])
B = np.array([r["B"] for r in d])
C = np.array([r["C"] for r in d])
sans = np.array([r["sansurlu"] for r in d])

ETIKET = ["A\nNaif", "B\nNominal\ndüzeltme", "C\nFiilî\ndüzeltme"]
GRI = ["#b0b0b0", "#707070", "#383838"]


def oranlar(p, y):
    m = (p > 0) & (y > 0)
    return p[m] / y[m]


fig, axes = plt.subplots(1, 2, figsize=(7.0, 3.9), sharey=True)

for ax, (mask, baslik) in zip(axes, [
        (sans,  f"(a) Eğitim döneminde sansürlenmiş\nkalemler (n = {sans.sum()})"),
        (~sans, f"(b) Eğitim döneminde kesintisiz\nkalemler — kontrol (n = {(~sans).sum()})")]):

    y = g[mask]
    veri = [oranlar(A[mask], y), oranlar(B[mask], y), oranlar(C[mask], y)]

    bp = ax.boxplot(veri, widths=0.55, showfliers=False, patch_artist=True,
                    medianprops=dict(color="black", linewidth=1.6),
                    whiskerprops=dict(linewidth=0.8),
                    capprops=dict(linewidth=0.8),
                    boxprops=dict(linewidth=0.8))
    for kutu, renk in zip(bp["boxes"], GRI):
        kutu.set_facecolor(renk)

    # tarafsızlık çizgisi
    ax.axhline(1.0, color="black", linestyle="--", linewidth=0.9, zorder=0)

    # medyan değerlerini yaz
    for i, v in enumerate(veri, start=1):
        med = np.median(v)
        ax.text(i, 2.24, f"{med:.3f}".replace(".", ","), ha="center", va="top",
                fontsize=8.5, fontweight="bold")

    ax.set_xticks([1, 2, 3])
    ax.set_xticklabels(ETIKET, fontsize=8.5)
    ax.set_ylim(0.0, 2.3)
    ax.set_title(baslik, fontsize=9, pad=8)
    ax.grid(axis="y", linewidth=0.4, alpha=0.35)
    ax.set_axisbelow(True)
    for yan in ("top", "right"):
        ax.spines[yan].set_visible(False)

axes[0].set_ylabel("Tahmin / gerçek talep oranı")
axes[0].yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:.1f}".replace(".", ",")))

fig.tight_layout(rect=[0, 0.045, 1, 1])
fig.text(0.5, 0.012,
         "Kutular çeyrekler arası aralığı, yatay çizgi medyanı gösterir; uç değerler "
         "gösterilmemiştir. Kesikli çizgi tam isabeti (1,0) belirtir.\n"
         "1,0'ın altındaki değerler eksik tahmini ifade eder. Medyan değerler "
         "kutuların üzerinde verilmiştir.",
         ha="center", fontsize=7.2, linespacing=1.5)

for uzanti, kwargs in [("png", dict(dpi=600)),
                       ("pdf", {}),
                       ("tif", dict(dpi=600, pil_kwargs={"compression": "tiff_lzw"}))]:
    fig.savefig(f"Sekil1_tahmin_gercek_orani.{uzanti}", bbox_inches="tight", **kwargs)
print("Şekil 1 üretildi: png, pdf, tif")
