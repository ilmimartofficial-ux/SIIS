"""
SMART INVENTORY INTELLIGENCE SYSTEM (SIIS) — iPOS 5 ENGINE
Author  : AI Senior Data Engineer
Version : 1.5.0  (dead-stock NAMA fix+master · FIFO layer HJ1/HJ2 margin · WA button+toggle · Rec HJ rekomendasi)
"""

import io
import urllib.parse
import warnings
from collections import deque
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
# THEME / CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="SIIS – iPOS 5 Engine",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main-header{font-size:2rem;font-weight:700;color:#1a7f37;margin-bottom:0}
    .sub-header{font-size:.9rem;color:#666;margin-bottom:1.5rem}
    .metric-card{background:#f8fff9;border:1px solid #c6e6c9;border-radius:10px;padding:1rem;text-align:center}
    .metric-warn{background:#fffde7;border:1px solid #ffe082;border-radius:10px;padding:1rem;text-align:center}
    .metric-crit{background:#fff3f3;border:1px solid #f5b8b8;border-radius:10px;padding:1rem;text-align:center}
    .section-title{font-size:1.2rem;font-weight:600;color:#1a7f37;border-left:4px solid #1a7f37;padding-left:.6rem;margin:1.2rem 0 .6rem}
    .stAlert{border-radius:8px}
    </style>
    """,
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────────────────────────
# SECTION 1 — DATA LOADING & CLEANING
# ─────────────────────────────────────────────────────────────────

def _clean_number(val):
    """Convert Indonesian-formatted number string to float."""
    if pd.isna(val):
        return np.nan
    s = str(val).replace(".", "").replace(",", ".").strip()
    try:
        return float(s)
    except ValueError:
        return np.nan


def load_item_perjumlah(f) -> pd.DataFrame:
    """
    ITEM_PERJUMLAH_IPOS_5.xlsx — header baris 0, tiap item bisa 2 baris:
      baris konversi=1  → satuan dasar (PCS)
      baris konversi>1  → satuan besar (DUS / SLOF / LSN dll)
    Kolom: Kode Item, Nama Item, Jenis, Konversi, Satuan, Harga Pokok,
           Jml 1, Harga Jml 1, Jml 2, Harga Jml 2
    """
    df = pd.read_excel(f, header=0, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    rename = {
        "Kode Item"  : "KODE",   "Nama Item"   : "NAMA",
        "Jenis"      : "JENIS",  "Merek"        : "MEREK",
        "Konversi"   : "KONVERSI","Satuan"       : "SATUAN",
        "Harga Pokok": "HARGA_POKOK",
        "Jml 1"      : "JML1",   "Harga Jml 1"  : "HARGA_JML1",
        "Jml 2"      : "JML2",   "Harga Jml 2"  : "HARGA_JML2",
    }
    df.rename(columns={k: v for k, v in rename.items() if k in df.columns}, inplace=True)
    df = df.dropna(subset=["KODE"]).copy()
    df["KODE"] = df["KODE"].str.strip()
    df = df[df["KODE"].str.len() > 0]
    for col in ["HARGA_POKOK","KONVERSI","JML1","HARGA_JML1","JML2","HARGA_JML2"]:
        if col in df.columns:
            df[col] = df[col].apply(_clean_number)
    df["HARGA_POKOK"] = df["HARGA_POKOK"].fillna(0)
    df["KONVERSI"]    = df["KONVERSI"].fillna(1)

    # Pisahkan baris PCS (konversi=1) dan satuan besar (konversi>1)
    df_pcs  = df[df["KONVERSI"] == 1].copy()
    df_bulk = (df[df["KONVERSI"] > 1][["KODE","SATUAN","KONVERSI"]]
               .drop_duplicates("KODE")
               .rename(columns={"SATUAN":"SAT_BESAR","KONVERSI":"KONV_BESAR"}))

    # Item yang hanya punya baris bulk → tetap dimasukkan
    only_bulk = set(df_bulk["KODE"]) - set(df_pcs["KODE"])
    if only_bulk:
        df_pcs = pd.concat([df_pcs, df[df["KODE"].isin(only_bulk)]], ignore_index=True)

    df_pcs = df_pcs.merge(df_bulk, on="KODE", how="left")
    return df_pcs.reset_index(drop=True)


def load_pembelian(f) -> pd.DataFrame:
    """
    PEMBELIAN_IPOS_5.xlsx — format iPOS 5 (30 kolom, 0-indexed):
      Header transaksi : col[1]=no_transaksi, col[7]=tanggal dd/mm/yyyy,
                         col[10]=dept,        col[12]=kode_supp,
                         col[16]=nama_supplier
      Baris kolom-item : col[1]='No.', col[2]='Kd. Item'
      Data item        : col[1]=no_urut, col[2]=kode, col[7]=nama,
                         col[17]=jml,    col[19]=satuan,
                         col[24]=harga,  col[25]=diskon%,  col[29]=total
    """
    raw = pd.read_excel(f, header=None, dtype=str)
    records=[]; current_date=None; current_sup=None; in_data=False

    for _, row in raw.iterrows():
        vals = [str(v).strip() if pd.notna(v) else "" for v in row]
        while len(vals) < 30:
            vals.append("")

        # ── Header blok transaksi: col[7] = tanggal dd/mm/yyyy ────
        c7 = vals[7]
        if len(c7) == 10 and c7[2] == "/" and c7[5] == "/":
            try:
                current_date = pd.to_datetime(c7, dayfirst=True)
                current_sup  = vals[16] if vals[16] else current_sup
                in_data = False; continue
            except Exception:
                pass

        # ── Baris kolom-header item ───────────────────────────────
        if vals[2].lower() in ("kd. item", "kd.item", "kode item", "kode"):
            in_data = True; continue

        if not in_data:
            continue

        kode = vals[2]
        if not kode or kode.lower() in ("nan", "none", ""):
            continue
        # Skip baris subtotal/total (nama item kosong)
        if not vals[7]:
            continue

        nama     = vals[7]
        qty_s    = vals[17]
        satuan   = vals[19]
        harga_s  = vals[24]
        diskon_s = vals[25] if vals[25] else "0"
        total_s  = vals[29]

        qty   = _clean_number(qty_s)
        harga = _clean_number(harga_s)
        diskon= _clean_number(diskon_s) or 0.0
        total = _clean_number(total_s)

        if pd.isna(qty) or pd.isna(harga): continue
        if qty <= 0 or harga <= 0: continue

        harga_net = harga * (1 - diskon / 100)
        if pd.isna(total) or total <= 0:
            total = qty * harga_net

        records.append({
            "TANGGAL"   : current_date, "SUPPLIER"  : current_sup,
            "KODE"      : kode.strip(), "NAMA"      : nama.strip(),
            "QTY"       : qty,          "SATUAN"    : satuan.strip(),
            "HARGA"     : harga,        "DISKON_PCT": diskon,
            "HARGA_NET" : harga_net,    "TOTAL"     : total,
        })

    df = pd.DataFrame(records)
    if df.empty:
        return df
    return df.sort_values("TANGGAL").reset_index(drop=True)


def load_penjualan(f) -> pd.DataFrame:
    """
    PENJUALAN_IPOS_5.xlsx — format iPOS 5 (30 kolom, 0-indexed):
      Header transaksi : col[1]=no_transaksi, col[7]=tanggal dd/mm/yyyy,
                         col[10]=dept,         col[12]=kode_pelanggan,
                         col[16]=nama_pelanggan
      Baris kolom-item : col[1]='No.', col[2]='Kd. Item'
      Data item        : col[1]=no_urut, col[2]=kode, col[7]=nama,
                         col[17]=jml,    col[19]=satuan,
                         col[24]=harga,  col[25]=diskon%, col[29]=total
    """
    raw = pd.read_excel(f, header=None, dtype=str)
    records=[]; current_date=None; current_pel=None; in_data=False

    for _, row in raw.iterrows():
        vals = [str(v).strip() if pd.notna(v) else "" for v in row]
        while len(vals) < 30:
            vals.append("")

        # ── Header blok transaksi: col[7] = tanggal dd/mm/yyyy ────
        c7 = vals[7]
        if len(c7) == 10 and c7[2] == "/" and c7[5] == "/":
            try:
                current_date = pd.to_datetime(c7, dayfirst=True)
                current_pel  = vals[16] if vals[16] else current_pel
                in_data = False; continue
            except Exception:
                pass

        # ── Baris kolom-header item ───────────────────────────────
        if vals[2].lower() in ("kd. item", "kd.item", "kode item", "kode"):
            in_data = True; continue

        if not in_data:
            continue

        kode = vals[2]
        if not kode or kode.lower() in ("nan", "none", ""):
            continue
        # Skip baris subtotal/total (nama item kosong)
        if not vals[7]:
            continue

        nama     = vals[7]
        qty_s    = vals[17]
        satuan   = vals[19]
        harga_s  = vals[24]
        diskon_s = vals[25] if vals[25] else "0"
        total_s  = vals[29]

        qty    = _clean_number(qty_s)
        harga  = _clean_number(harga_s)
        diskon = _clean_number(diskon_s) or 0.0
        total  = _clean_number(total_s)

        if (pd.isna(qty) or qty == 0) and (pd.isna(total) or total == 0):
            continue

        qty   = qty   if not pd.isna(qty)   else 0.0
        harga = harga if not pd.isna(harga) else 0.0
        total = total if not pd.isna(total) else 0.0
        harga_jual = total / qty if qty > 0 else harga

        records.append({
            "TANGGAL"    : current_date,
            "PELANGGAN"  : current_pel,
            "KODE"       : kode.strip(),
            "NAMA"       : nama.strip(),
            "QTY"        : qty,
            "SATUAN"     : satuan.strip(),
            "HARGA_JUAL" : harga_jual,
            "DISKON_PCT" : diskon,
            "TOTAL"      : total,
        })

    df = pd.DataFrame(records)
    if df.empty:
        return df
    return df.sort_values("TANGGAL").reset_index(drop=True)


def load_mutasi(f) -> pd.DataFrame:
    """
    MUTASI_ITEM_IPOS_5.xlsx — 26 kolom (0-indexed):
      Periode  : baris yang mengandung "PERIODE : dd/mm/yyyy - dd/mm/yyyy"
      Header   : col[1]='Kode Item' (baris ~13)
      Data kolom: col[1]=kode, col[5]=nama, col[10]=satuan,
                  col[11]=awal_qty,  col[13]=awal_nilai,
                  col[15]=masuk_qty, col[16]=masuk_nilai,
                  col[19]=keluar_qty,col[22]=keluar_nilai,
                  col[24]=akhir_qty, col[25]=akhir_nilai
    """
    raw = pd.read_excel(f, header=None, dtype=str)

    # Deteksi periode
    periode = None
    for _, row in raw.iterrows():
        for v in row:
            s = str(v)
            if "PERIODE" in s.upper() and "/" in s:
                try:
                    part = s.split(":")[1].strip()
                    tgl  = part.split("-")[0].strip()
                    periode = pd.to_datetime(tgl, dayfirst=True); break
                except Exception:
                    pass
        if periode:
            break

    # Cari header: col[1] == "Kode Item"
    header_row = None
    for i, row in raw.iterrows():
        v1 = str(row.iloc[1]).strip().lower() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
        if v1 == "kode item":
            header_row = i; break
    if header_row is None:
        header_row = 13

    records = []
    for _, row in raw.iloc[header_row + 1:].iterrows():
        vals = [
            str(v).strip() if pd.notna(v) and str(v).strip() not in ("nan","None") else ""
            for v in row
        ]
        while len(vals) < 26:
            vals.append("")

        kode = vals[1]
        if not kode or kode.lower() in ("kode item","total","sub total","nan",""):
            continue

        def _g(idx):
            return _clean_number(vals[idx]) if len(vals) > idx and vals[idx] else np.nan

        records.append({
            "TANGGAL"     : periode,
            "KODE"        : kode.strip(),
            "NAMA"        : vals[5],
            "SATUAN"      : vals[10],
            "AWAL_QTY"   : _g(11),  "AWAL_NILAI"  : _g(13),
            "MASUK_QTY"  : _g(15),  "MASUK_NILAI" : _g(16),
            "KELUAR_QTY" : _g(19),  "KELUAR_NILAI": _g(22),
            "AKHIR_QTY"  : _g(24),  "AKHIR_NILAI" : _g(25),
        })

    df = pd.DataFrame(records)
    if df.empty:
        return df
    df = df[df["KODE"].str.len() > 0]
    num_cols = ["AWAL_QTY","AWAL_NILAI","MASUK_QTY","MASUK_NILAI",
                "KELUAR_QTY","KELUAR_NILAI","AKHIR_QTY","AKHIR_NILAI"]
    for c in num_cols:
        df[c] = df[c].fillna(0)
    return df.reset_index(drop=True)



# ─────────────────────────────────────────────────────────────────
# SECTION 2 — FIFO ENGINE
# ─────────────────────────────────────────────────────────────────

def run_fifo(df_beli: pd.DataFrame, df_jual: pd.DataFrame,
             margin_min: float = 5.0, df_master: pd.DataFrame = None):
    """
    FIFO cost calculation with proactive layer-conflict alerts.

    ★ FIX v1.2 — Baris `konv = konversi_map.get(k, 1.0)` diganti dengan
    `konv = _get_konv(k, satuan_beli)` sehingga konversi SLOF/DUS/LSN → PCS
    benar-benar diterapkan saat:
      1) Normalisasi harga beli ke per-PCS
      2) Konversi qty beli ke PCS untuk antrian FIFO
      3) Perbandingan harga modal PCS vs harga jual PCS

    Returns:
        df_jual_fifo   : penjualan dengan HPP_FIFO, MARGIN_PCT, STATUS
        fifo_queues    : sisa antrian stok per item (dalam satuan PCS)
        df_remaining   : sisa stok per item beserta nilainya
        df_layer_alert : peringatan konflik lapisan FIFO
    """
    # ── Bangun lookup konversi: {KODE: {SAT_BESAR: faktor}} ─────
    # Contoh: {"8999909028234": {"SLOF": 10, "PCS": 1}}
    konversi_map: dict[str, dict] = {}
    if df_master is not None and not df_master.empty:
        for _, mr in df_master.iterrows():
            k = str(mr.get("KODE","")).strip()
            if not k:
                continue
            if k not in konversi_map:
                konversi_map[k] = {"PCS": 1.0}
            sat_besar  = str(mr.get("SAT_BESAR","") or "").strip().upper()
            konv_besar = mr.get("KONV_BESAR", None)
            try:
                konv_besar = float(konv_besar) if konv_besar and not pd.isna(konv_besar) else None
            except (ValueError, TypeError):
                konv_besar = None
            if sat_besar and konv_besar and konv_besar > 1:
                konversi_map[k][sat_besar] = konv_besar

    def _get_konv(kode: str, satuan_beli: str) -> float:
        """
        Faktor konversi satuan beli → PCS.
        Contoh: SLOF → 10, DUS → 12, LSN → 20, PCS → 1
        Default 1 jika tidak ditemukan di master.
        """
        sat = satuan_beli.strip().upper()
        mapping = konversi_map.get(kode, {})
        if sat in mapping:
            return mapping[sat]
        # Fuzzy match substring (misal "SLOF" cocok dengan "SLOF")
        for k_sat, v in mapping.items():
            if k_sat and sat and (k_sat in sat or sat in k_sat):
                return v
        return 1.0

    # ── Harga jual dari MASTER (HARGA_JML1 & HARGA_JML2) ───────────
    #    JML1 = harga eceran (PCS), JML2 = harga grosir (satuan besar)
    harga_jml1_map: dict[str, float] = {}
    harga_jml2_map: dict[str, float] = {}
    if df_master is not None and not df_master.empty:
        for _, mr in df_master.iterrows():
            k = str(mr.get("KODE","")).strip()
            if not k:
                continue
            h1 = mr.get("HARGA_JML1", None)
            h2 = mr.get("HARGA_JML2", None)
            try:
                if h1 is not None and not pd.isna(h1) and float(h1) > 0:
                    harga_jml1_map[k] = float(h1)
            except (TypeError, ValueError):
                pass
            try:
                if h2 is not None and not pd.isna(h2) and float(h2) > 0:
                    harga_jml2_map[k] = float(h2)
            except (TypeError, ValueError):
                pass

    # ── Harga jual per PCS dari penjualan (rata-rata tertimbang) ─
    #    Dipakai sebagai fallback jika item tidak ada di master
    harga_jual_map: dict[str, float] = {}
    if not df_jual.empty:
        for k, grp in df_jual.groupby("KODE"):
            total_val = grp["TOTAL"].sum()
            total_qty = grp["QTY"].sum()
            harga_jual_map[k] = total_val / total_qty if total_qty > 0 else 0.0

    def _get_harga_jual_pcs(kode: str) -> float:
        """Ambil harga jual PCS: utamakan JML1 dari master, fallback ke rata2 penjualan."""
        h = harga_jml1_map.get(kode, 0.0)
        if h > 0:
            return h
        return harga_jual_map.get(kode, 0.0)

    # ── Isi antrian FIFO & deteksi konflik lapisan ────────────────
    queues: dict[str, deque] = {}
    layer_alerts: list[dict] = []

    for _, row in df_beli.sort_values("TANGGAL").iterrows():
        k = row["KODE"]
        satuan_beli = str(row.get("SATUAN", "") or "").strip()

        # ★ FIX: Gunakan _get_konv() bukan konversi_map.get()
        konv = _get_konv(k, satuan_beli)

        # Harga beli per PCS = harga beli per satuan besar / faktor konversi
        # Contoh: beli SLOF @ Rp 150.000, konversi 10 → Rp 15.000/PCS
        harga_beli_pcs = float(row["HARGA_NET"]) / konv

        # Qty dalam PCS untuk antrian FIFO
        # Contoh: beli 5 SLOF, konversi 10 → 50 PCS masuk antrian
        qty_pcs = float(row["QTY"]) * konv

        if k not in queues:
            queues[k] = deque()

        # ── CEK KONFLIK: stok lama belum habis, beli baru masuk ──
        if queues[k]:
            harga_jual_pcs = _get_harga_jual_pcs(k)   # dari master JML1 / fallback penjualan
            h_jml1         = harga_jml1_map.get(k, 0.0)
            h_jml2         = harga_jml2_map.get(k, 0.0)
            nama_item      = str(row.get("NAMA", k))
            supplier_baru  = str(row.get("SUPPLIER", ""))
            tgl_baru       = row["TANGGAL"]

            layer_lama  = queues[k][0]
            qty_lama    = layer_lama["qty"]
            harga_lama  = layer_lama["harga"]
            sup_lama    = layer_lama["supplier"]
            tgl_lama    = layer_lama["tanggal"]

            def _margin_pct(hj, hm):
                return round((hj - hm) / hj * 100, 1) if hj > 0 else 0.0

            # Harga asli sebelum konversi (per satuan beli, misal per SLOF/DUS)
            harga_beli_asli = float(row["HARGA_NET"])

            def _build_alert(alert_type, severity, keterangan):
                return {
                    "KODE"            : k,
                    "NAMA"            : nama_item,
                    "SUPPLIER_LAMA"   : sup_lama,
                    "TGL_LAMA"        : tgl_lama,
                    "QTY_SISA_LAMA"   : round(qty_lama, 1),
                    "HARGA_LAMA_PCS"  : round(harga_lama, 0),
                    "SUPPLIER_BARU"   : supplier_baru,
                    "TGL_BARU"        : tgl_baru,
                    # Harga asli per satuan beli (sebelum ÷ konv)
                    "HARGA_BELI_ASLI" : round(harga_beli_asli, 0),
                    # Harga per PCS setelah konversi (÷ konv)
                    "HARGA_BARU_PCS"  : round(harga_beli_pcs, 0),
                    # Harga jual dari master perjumlah
                    "HARGA_JML1"      : round(h_jml1, 0),
                    "HARGA_JML2"      : round(h_jml2, 0),
                    "MARGIN_JML1_PCT" : _margin_pct(h_jml1, harga_beli_pcs),
                    "MARGIN_JML2_PCT" : _margin_pct(h_jml2, harga_beli_pcs),
                    # Harga jual efektif (JML1 atau fallback)
                    "HARGA_JUAL_PCS"  : round(harga_jual_pcs, 0),
                    "MARGIN_BARU_PCT" : _margin_pct(harga_jual_pcs, harga_beli_pcs),
                    "ALERT_TYPE"      : alert_type,
                    "SEVERITY"        : severity,
                    "SATUAN_BELI"     : satuan_beli,
                    "KONV_USED"       : konv,
                    "KETERANGAN"      : keterangan,
                }

            tgl_str = pd.Timestamp(tgl_lama).strftime('%d/%m/%y') if pd.notna(tgl_lama) else '-'

            # Kondisi 1 – Modal baru > Harga Jual → pasti rugi
            if harga_jual_pcs > 0 and harga_beli_pcs > harga_jual_pcs:
                ket = (
                    f"Stok lama {sup_lama} ({qty_lama:.0f} PCS tgl {tgl_str}) belum habis. "
                    f"Beli baru {satuan_beli} ×{konv:.0f} → Rp {harga_beli_pcs:,.0f}/PCS "
                    f"> HJ1 Rp {h_jml1:,.0f} (margin {_margin_pct(h_jml1,harga_beli_pcs):.1f}%) "
                    f"| HJ2 Rp {h_jml2:,.0f} (margin {_margin_pct(h_jml2,harga_beli_pcs):.1f}%) → RUGI!"
                )
                layer_alerts.append(_build_alert("🔴 MODAL > HARGA JUAL", "KRITIS", ket))

            # Kondisi 2 – Modal baru < Harga Jual TAPI margin < threshold
            elif harga_jual_pcs > 0 and harga_beli_pcs <= harga_jual_pcs:
                margin_baru = _margin_pct(harga_jual_pcs, harga_beli_pcs)
                if 0 < margin_baru < margin_min:
                    ket = (
                        f"Stok lama {sup_lama} ({qty_lama:.0f} PCS) belum habis. "
                        f"Beli baru {satuan_beli} ×{konv:.0f} → Rp {harga_beli_pcs:,.0f}/PCS. "
                        f"HJ1 Rp {h_jml1:,.0f} → margin {_margin_pct(h_jml1,harga_beli_pcs):.1f}% | "
                        f"HJ2 Rp {h_jml2:,.0f} → margin {_margin_pct(h_jml2,harga_beli_pcs):.1f}% "
                        f"(< target {margin_min:.0f}%)"
                    )
                    layer_alerts.append(_build_alert(
                        f"🟡 MARGIN DI BAWAH {margin_min:.0f}%", "WASPADA", ket))

            # Kondisi 3 – Modal baru naik signifikan (>15%) dibanding layer lama
            if harga_lama > 0:
                kenaikan_pct = (harga_beli_pcs - harga_lama) / harga_lama * 100
                if kenaikan_pct > 15:
                    already = any(
                        a["KODE"] == k and a["TGL_BARU"] == tgl_baru
                        for a in layer_alerts
                    )
                    if not already:
                        ket = (
                            f"Modal lama Rp {harga_lama:,.0f}/PCS ({qty_lama:.0f} PCS) belum habis. "
                            f"Beli baru {satuan_beli} ×{konv:.0f} naik {kenaikan_pct:.1f}% "
                            f"→ Rp {harga_beli_pcs:,.0f}/PCS. "
                            f"HJ1 Rp {h_jml1:,.0f} → margin {_margin_pct(h_jml1,harga_beli_pcs):.1f}% | "
                            f"HJ2 Rp {h_jml2:,.0f} → margin {_margin_pct(h_jml2,harga_beli_pcs):.1f}%"
                        )
                        layer_alerts.append(_build_alert(
                            f"🟠 KENAIKAN MODAL {kenaikan_pct:.0f}%", "WASPADA", ket))

        # Masukkan layer baru ke antrian (sudah dinormalisasi per PCS)
        queues[k].append({
            "qty"     : qty_pcs,
            "harga"   : harga_beli_pcs,
            "supplier": str(row.get("SUPPLIER", "")),
            "tanggal" : row["TANGGAL"],
        })

    # ── Proses penjualan dengan FIFO ─────────────────────────────
    #    Penjualan biasanya sudah dalam satuan PCS
    results = []
    for _, row in df_jual.iterrows():
        k = row["KODE"]
        qty_needed   = float(row["QTY"])
        hpp_total    = 0.0
        qty_fulfilled= 0.0
        warning      = ""

        if k not in queues or len(queues[k]) == 0:
            warning   = "NO_STOCK"
            hpp_total = 0.0
        else:
            q         = queues[k]
            remaining = qty_needed
            while remaining > 0 and q:
                layer = q[0]
                take  = min(remaining, layer["qty"])
                hpp_total    += take * layer["harga"]
                qty_fulfilled += take
                layer["qty"] -= take
                remaining    -= take
                if layer["qty"] <= 0:
                    q.popleft()
            if remaining > 0:
                warning = f"STOK_MINUS({remaining:.1f})"

        hpp_per_unit = hpp_total / qty_fulfilled if qty_fulfilled > 0 else 0
        total_jual   = float(row.get("TOTAL", 0))
        margin_rp    = total_jual - hpp_total
        margin_pct   = (margin_rp / total_jual * 100) if total_jual > 0 else 0

        status = "KRITIS" if margin_pct < margin_min else "AMAN"

        results.append({
            **row.to_dict(),
            "HPP_FIFO_UNIT" : round(hpp_per_unit, 0),
            "HPP_FIFO_TOTAL": round(hpp_total, 0),
            "MARGIN_RP"     : round(margin_rp, 0),
            "MARGIN_PCT"    : round(margin_pct, 1),
            "STATUS"        : status,
            "FIFO_WARNING"  : warning,
        })

    # ── Sisa stok (remaining layers) ─────────────────────────────
    remaining_rows = []
    for k, q in queues.items():
        for layer in q:
            remaining_rows.append({
                "KODE"       : k,
                "QTY_SISA"   : layer["qty"],           # dalam PCS
                "HARGA_LAYER": layer["harga"],          # per PCS
                "NILAI_LAYER": layer["qty"] * layer["harga"],
                "SUPPLIER"   : layer["supplier"],
                "TGL_MASUK"  : layer["tanggal"],
            })

    df_fifo      = pd.DataFrame(results)
    df_remaining = pd.DataFrame(remaining_rows) if remaining_rows else pd.DataFrame(
        columns=["KODE","QTY_SISA","HARGA_LAYER","NILAI_LAYER","SUPPLIER","TGL_MASUK"]
    )
    df_layer_alert = pd.DataFrame(layer_alerts) if layer_alerts else pd.DataFrame(
        columns=["KODE","NAMA","SUPPLIER_LAMA","TGL_LAMA","QTY_SISA_LAMA",
                 "HARGA_LAMA_PCS","SUPPLIER_BARU","TGL_BARU","HARGA_BARU_PCS",
                 "HARGA_JML1","HARGA_JML2","MARGIN_JML1_PCT","MARGIN_JML2_PCT",
                 "HARGA_JUAL_PCS","MARGIN_BARU_PCT","ALERT_TYPE","SEVERITY",
                 "SATUAN_BELI","KONV_USED","KETERANGAN"]
    )
    return df_fifo, queues, df_remaining, df_layer_alert


# ─────────────────────────────────────────────────────────────────
# SECTION 3 — ABC MOVEMENT ANALYSIS
# ─────────────────────────────────────────────────────────────────

def abc_analysis(df_jual: pd.DataFrame) -> pd.DataFrame:
    """ABC classification by total sales value."""
    if df_jual.empty:
        return pd.DataFrame()

    grp = (
        df_jual.groupby("KODE")
        .agg(
            NAMA=("NAMA", "first"),
            TOTAL_QTY=("QTY", "sum"),
            TOTAL_OMZET=("TOTAL", "sum"),
            TRX_COUNT=("KODE", "count"),
        )
        .reset_index()
        .sort_values("TOTAL_OMZET", ascending=False)
    )

    grp["OMZET_CUM_PCT"] = grp["TOTAL_OMZET"].cumsum() / grp["TOTAL_OMZET"].sum() * 100
    grp["KELAS_ABC"] = grp["OMZET_CUM_PCT"].apply(
        lambda x: "A" if x <= 80 else ("B" if x <= 95 else "C")
    )
    return grp.reset_index(drop=True)


def detect_dead_stock(df_jual: pd.DataFrame, df_remaining: pd.DataFrame,
                      df_mutasi: pd.DataFrame = None, cutoff_days: int = 30,
                      df_master: pd.DataFrame = None) -> pd.DataFrame:
    """
    Deteksi Dead Stock untuk format iPOS (tanpa tanggal per item di penjualan).
    NAMA diambil dari penjualan → mutasi → master → fallback kode.
    """
    if df_remaining.empty:
        return pd.DataFrame()

    stock_summary = (
        df_remaining.groupby("KODE")
        .agg(STOK_SISA=("QTY_SISA", "sum"), NILAI_SISA=("NILAI_LAYER", "sum"))
        .reset_index()
    )

    # ★ Bangun nama_map berlapis: penjualan → mutasi → master → ""
    nama_map: dict = {}
    if not df_jual.empty and "NAMA" in df_jual.columns:
        nama_map = (df_jual.dropna(subset=["KODE"])
                    .groupby("KODE")["NAMA"].first().to_dict())
    # Fallback ke mutasi untuk item yang tidak ada di penjualan
    if df_mutasi is not None and not df_mutasi.empty and "NAMA" in df_mutasi.columns:
        for kode, nama in (df_mutasi.dropna(subset=["KODE"])
                           .groupby("KODE")["NAMA"].first().items()):
            if kode not in nama_map or not nama_map[kode] or str(nama_map[kode]).lower() in ("nan","none",""):
                nama_map[kode] = nama
    # Fallback ke master item perjumlah
    if df_master is not None and not df_master.empty and "NAMA" in df_master.columns:
        for kode, nama in (df_master.dropna(subset=["KODE"])
                           .groupby("KODE")["NAMA"].first().items()):
            if kode not in nama_map or not nama_map[kode] or str(nama_map[kode]).lower() in ("nan","none",""):
                nama_map[kode] = nama

    def _safe_nama(kode):
        n = nama_map.get(kode, "")
        if not n or str(n).lower() in ("nan","none",""):
            return f"[{kode}]"   # tampilkan kode dalam kurung siku, bukan "None"
        return str(n).strip()

    stock_summary["NAMA"] = stock_summary["KODE"].apply(_safe_nama)

    kode_jual = set(df_jual["KODE"].unique()) if not df_jual.empty else set()

    kode_mutasi_bergerak = set()
    if df_mutasi is not None and not df_mutasi.empty and "KELUAR_QTY" in df_mutasi.columns:
        kode_mutasi_bergerak = set(
            df_mutasi[df_mutasi["KELUAR_QTY"] > 0]["KODE"].unique()
        )

    dead = []
    for _, r in stock_summary.iterrows():
        k = r["KODE"]
        if r["STOK_SISA"] <= 0:
            continue
        is_dead = (k not in kode_jual) and (k not in kode_mutasi_bergerak)
        if is_dead:
            dead.append({
                "KODE"         : k,
                "NAMA"         : r["NAMA"],
                "STOK_SISA"    : r["STOK_SISA"],
                "NILAI_SISA"   : r["NILAI_SISA"],
                "STATUS_JUAL"  : "TIDAK ADA",
                "STATUS_MUTASI": "TIDAK BERGERAK",
                "STATUS"       : "DEAD",
            })

    return pd.DataFrame(dead)
# ─────────────────────────────────────────────────────────────────
# SECTION 4 — AUDIT & ANOMALY DETECTION
# ─────────────────────────────────────────────────────────────────

def audit_stock(df_mutasi: pd.DataFrame) -> pd.DataFrame:
    """Validate: Awal + Masuk - Keluar == Akhir."""
    if df_mutasi.empty:
        return pd.DataFrame()
    df = df_mutasi.copy()
    df["CALC_AKHIR"] = df["AWAL_QTY"] + df["MASUK_QTY"] - df["KELUAR_QTY"]
    df["SELISIH"] = df["CALC_AKHIR"] - df["AKHIR_QTY"]
    df["AUDIT_STATUS"] = df["SELISIH"].apply(lambda x: "ERROR" if abs(x) > 0.01 else "OK")
    return df[df["AUDIT_STATUS"] == "ERROR"][
        ["KODE","NAMA","AWAL_QTY","MASUK_QTY","KELUAR_QTY","AKHIR_QTY","CALC_AKHIR","SELISIH","AUDIT_STATUS"]
    ].reset_index(drop=True)


def anomaly_harga(df_beli: pd.DataFrame, threshold: float = 0.30) -> pd.DataFrame:
    """Flag purchase prices > threshold% above mean per item."""
    if df_beli.empty:
        return pd.DataFrame()
    if "HARGA_NET" not in df_beli.columns:
        return pd.DataFrame()
    stats = df_beli.groupby("KODE")["HARGA_NET"].agg(["mean", "std"]).reset_index()
    stats.columns = ["KODE", "HARGA_MEAN", "HARGA_STD"]
    df = df_beli.merge(stats, on="KODE", how="left")
    df["ANOMALY"] = df.apply(
        lambda r: abs(r["HARGA_NET"] - r["HARGA_MEAN"]) / r["HARGA_MEAN"] > threshold
        if r["HARGA_MEAN"] > 0 else False,
        axis=1,
    )
    return df[df["ANOMALY"]][
        ["TANGGAL","SUPPLIER","KODE","NAMA","SATUAN","HARGA_NET","HARGA_MEAN","HARGA_STD"]
    ].reset_index(drop=True)


# ─────────────────────────────────────────────────────────────────
# SECTION 5 — SMART REORDER
# ─────────────────────────────────────────────────────────────────

def smart_reorder(df_jual: pd.DataFrame, df_remaining: pd.DataFrame,
                  df_beli: pd.DataFrame, lead_time: int = 3) -> pd.DataFrame:
    """Calculate reorder recommendations per item."""
    if df_jual.empty or df_remaining.empty:
        return pd.DataFrame()

    grp = df_jual.groupby("KODE").agg(
        NAMA=("NAMA", "first"),
        TOTAL_QTY=("QTY", "sum"),
    ).reset_index()

    days = 30
    grp["AVG_DAILY"] = grp["TOTAL_QTY"] / days
    grp["MAX_DAILY"] = grp["TOTAL_QTY"] / max(days // 2, 1)

    stock = df_remaining.groupby("KODE")["QTY_SISA"].sum().reset_index()
    stock.columns = ["KODE", "STOK_SISA"]

    grp = grp.merge(stock, on="KODE", how="left")
    grp["STOK_SISA"] = grp["STOK_SISA"].fillna(0)

    grp["SAFETY_STOCK"] = (grp["MAX_DAILY"] - grp["AVG_DAILY"]) * lead_time
    grp["SAFETY_STOCK"] = grp["SAFETY_STOCK"].clip(lower=0)
    grp["ROP"] = grp["AVG_DAILY"] * lead_time + grp["SAFETY_STOCK"]
    grp["PERLU_ORDER"] = grp["STOK_SISA"] < grp["ROP"]
    grp["QTY_ORDER"] = (grp["ROP"] * 2 - grp["STOK_SISA"]).clip(lower=0).round()

    if not df_beli.empty:
        # Ambil supplier + harga modal terakhir per item dari pembelian
        last_beli = (
            df_beli.sort_values("TANGGAL")
            .groupby("KODE")
            .agg(SUPPLIER=("SUPPLIER", "last"),
                 HARGA_MODAL=("HARGA_NET", "last"))
            .reset_index()
        )
        grp = grp.merge(last_beli, on="KODE", how="left")
        grp["HARGA_MODAL"] = grp["HARGA_MODAL"].fillna(0)
    else:
        grp["SUPPLIER"] = ""
        grp["HARGA_MODAL"] = 0

    reorder = grp[grp["PERLU_ORDER"]].copy()
    reorder["AVG_DAILY"]    = reorder["AVG_DAILY"].round(1)
    reorder["SAFETY_STOCK"] = reorder["SAFETY_STOCK"].round(1)
    reorder["ROP"]          = reorder["ROP"].round(1)
    reorder["STOK_SISA"]    = reorder["STOK_SISA"].round(1)
    reorder["QTY_ORDER"]    = reorder["QTY_ORDER"].round(0).astype(int)
    reorder["HARGA_MODAL"]  = reorder["HARGA_MODAL"].round(0)
    return reorder.sort_values("SUPPLIER").reset_index(drop=True)


# ─────────────────────────────────────────────────────────────────
# SECTION 6 — UNIT CONVERSION VALIDATION
# ─────────────────────────────────────────────────────────────────

def validate_konversi(df_master: pd.DataFrame, df_fifo: pd.DataFrame) -> pd.DataFrame:
    """Validate: harga grosir < retail AND > HPP FIFO."""
    if df_master.empty or df_fifo.empty:
        return pd.DataFrame()

    hpp = df_fifo.groupby("KODE")["HPP_FIFO_UNIT"].mean().reset_index()
    hpp.columns = ["KODE", "HPP_MEAN"]

    merged = df_master.merge(hpp, on="KODE", how="inner")
    issues = []
    for _, r in merged.iterrows():
        if pd.isna(r.get("HARGA_JML1", np.nan)):
            continue
        h1 = r.get("HARGA_JML1", 0) or 0
        hpp_val = r.get("HPP_MEAN", 0) or 0
        hp = r.get("HARGA_POKOK", 0) or 0
        if h1 > 0 and hpp_val > 0:
            if h1 <= hpp_val:
                issues.append({
                    "KODE": r["KODE"],
                    "NAMA": r.get("NAMA", ""),
                    "HARGA_JML1": h1,
                    "HPP_FIFO": round(hpp_val, 0),
                    "HARGA_POKOK": hp,
                    "ISSUE": "HARGA GROSIR ≤ HPP FIFO",
                })
    return pd.DataFrame(issues)


# ─────────────────────────────────────────────────────────────────
# SECTION 7 — SMART RECOMMENDATIONS
# ─────────────────────────────────────────────────────────────────

def generate_recommendations(df_fifo, df_dead, df_abc, df_reorder, df_layer_alert=None,
                             margin_min: float = 5.0):
    recs = []

    def _s(row, col):
        return str(row[col]) if col in row.index and pd.notna(row[col]) else ""

    def _f(val, fmt="Rp {:,.0f}"):
        try:
            return fmt.format(float(val)) if pd.notna(val) and float(val) > 0 else "-"
        except (TypeError, ValueError):
            return "-"

    # ── FIFO Layer: KRITIS & WASPADA + rekomendasi harga JML1/JML2 ──
    if df_layer_alert is not None and not df_layer_alert.empty and "SEVERITY" in df_layer_alert.columns:
        kritis_layer = df_layer_alert[df_layer_alert["SEVERITY"] == "KRITIS"].head(5)
        for _, r in kritis_layer.iterrows():
            h1  = r.get("HARGA_JML1", 0) or 0
            h2  = r.get("HARGA_JML2", 0) or 0
            hm  = r.get("HARGA_BARU_PCS", 0) or 0
            m1  = r.get("MARGIN_JML1_PCT", 0) or 0
            m2  = r.get("MARGIN_JML2_PCT", 0) or 0

            # Hitung harga rekomendasi jika di bawah modal
            rec_hj1 = round(hm * (1 + margin_min / 100)) if (h1 <= hm and hm > 0) else h1
            rec_hj2 = round(hm * (1 + margin_min * 0.7 / 100)) if (h2 <= hm and hm > 0) else h2

            saran_hj = ""
            if h1 > 0 and h1 > hm:
                saran_hj += f" | ✅ HJ1 Rp {h1:,.0f} (margin {m1:.1f}%)"
            elif h1 > 0:
                saran_hj += f" | ❌ HJ1 Rp {h1:,.0f} (di bawah modal {m1:.1f}%) → Rec: Rp {rec_hj1:,.0f}"
            if h2 > 0 and h2 > hm:
                saran_hj += f" | ✅ HJ2 Rp {h2:,.0f} (margin {m2:.1f}%)"
            elif h2 > 0:
                saran_hj += f" | ❌ HJ2 Rp {h2:,.0f} (di bawah modal {m2:.1f}%) → Rec: Rp {rec_hj2:,.0f}"
            recs.append({
                "TYPE": "🔴 FIFO LAYER KRITIS", "KODE": r["KODE"], "NAMA": _s(r, "NAMA"),
                "REKOMENDASI": (
                    f"{_s(r,'KETERANGAN')} — "
                    f"Modal baru: Rp {hm:,.0f}/PCS.{saran_hj}. "
                    f"Naikkan harga jual atau tunda pembelian."
                ),
                "PRIORITAS": "TINGGI",
                "HARGA_JML1": h1, "HARGA_JML2": h2,
                "MARGIN_JML1_PCT": m1, "MARGIN_JML2_PCT": m2,
                "HARGA_JML1_REC": rec_hj1 if h1 <= hm else 0,
                "HARGA_JML2_REC": rec_hj2 if h2 <= hm else 0,
                "HARGA_MODAL": hm,
            })
        waspada_layer = df_layer_alert[df_layer_alert["SEVERITY"] == "WASPADA"].head(3)
        for _, r in waspada_layer.iterrows():
            h1 = r.get("HARGA_JML1", 0) or 0
            h2 = r.get("HARGA_JML2", 0) or 0
            hm = r.get("HARGA_BARU_PCS", 0) or 0
            m1 = r.get("MARGIN_JML1_PCT", 0) or 0
            m2 = r.get("MARGIN_JML2_PCT", 0) or 0
            rec_hj1 = round(hm * (1 + margin_min / 100)) if (h1 <= hm and hm > 0) else h1
            rec_hj2 = round(hm * (1 + margin_min * 0.7 / 100)) if (h2 <= hm and hm > 0) else h2
            saran_hj = ""
            if h1 > 0:
                status1 = "✅" if h1 > hm else "❌"
                saran_hj += f" | {status1} HJ1 Rp {h1:,.0f} (margin {m1:.1f}%)"
                if h1 <= hm:
                    saran_hj += f" → Rec: Rp {rec_hj1:,.0f}"
            if h2 > 0:
                status2 = "✅" if h2 > hm else "❌"
                saran_hj += f" | {status2} HJ2 Rp {h2:,.0f} (margin {m2:.1f}%)"
                if h2 <= hm:
                    saran_hj += f" → Rec: Rp {rec_hj2:,.0f}"
            recs.append({
                "TYPE": "🟡 FIFO LAYER WASPADA", "KODE": r["KODE"], "NAMA": _s(r, "NAMA"),
                "REKOMENDASI": (
                    f"{_s(r,'KETERANGAN')}{saran_hj}. "
                    f"Negosiasi harga atau cari supplier alternatif."
                ),
                "PRIORITAS": "SEDANG",
                "HARGA_JML1": h1, "HARGA_JML2": h2,
                "MARGIN_JML1_PCT": m1, "MARGIN_JML2_PCT": m2,
                "HARGA_JML1_REC": rec_hj1 if h1 <= hm else 0,
                "HARGA_JML2_REC": rec_hj2 if h2 <= hm else 0,
                "HARGA_MODAL": hm,
            })

    # ── Margin KRITIS dari FIFO ──
    if not df_fifo.empty:
        kritis = df_fifo[df_fifo["STATUS"] == "KRITIS"].groupby("KODE").agg(
            NAMA=("NAMA","first"), MARGIN_PCT=("MARGIN_PCT","mean"),
            HPP_FIFO_UNIT=("HPP_FIFO_UNIT","mean")
        ).reset_index().nsmallest(5, "MARGIN_PCT")
        for _, r in kritis.iterrows():
            recs.append({
                "TYPE": "💰 MARGIN KRITIS", "KODE": r["KODE"], "NAMA": _s(r, "NAMA"),
                "REKOMENDASI": (
                    f"Margin rata-rata {r['MARGIN_PCT']:.1f}% — HPP FIFO Rp {r['HPP_FIFO_UNIT']:,.0f}/PCS. "
                    f"Naikkan harga jual minimum di atas HPP."
                ),
                "PRIORITAS": "TINGGI",
                "HARGA_JML1": 0, "HARGA_JML2": 0,
                "MARGIN_JML1_PCT": 0, "MARGIN_JML2_PCT": 0,
                "HARGA_JML1_REC": 0, "HARGA_JML2_REC": 0, "HARGA_MODAL": r['HPP_FIFO_UNIT'],
            })

    # ── Dead Stock ──
    if not df_dead.empty:
        for _, r in df_dead.head(5).iterrows():
            recs.append({
                "TYPE": "🧊 DEAD STOCK", "KODE": r["KODE"], "NAMA": _s(r, "NAMA"),
                "REKOMENDASI": (
                    f"Sisa {r['STOK_SISA']:.0f} PCS (Rp {r['NILAI_SISA']:,.0f}) tidak terjual & "
                    f"tidak bergerak. Buat promo diskon atau ajukan retur ke supplier."
                ),
                "PRIORITAS": "SEDANG",
                "HARGA_JML1": 0, "HARGA_JML2": 0,
                "MARGIN_JML1_PCT": 0, "MARGIN_JML2_PCT": 0,
                "HARGA_JML1_REC": 0, "HARGA_JML2_REC": 0, "HARGA_MODAL": 0,
            })

    # ── ABC Kelas C ──
    if not df_abc.empty:
        kelas_c = df_abc[df_abc["KELAS_ABC"] == "C"].nsmallest(3, "TOTAL_OMZET")
        for _, r in kelas_c.iterrows():
            recs.append({
                "TYPE": "📉 OMZET RENDAH", "KODE": r["KODE"], "NAMA": _s(r, "NAMA"),
                "REKOMENDASI": f"Omzet Rp {r['TOTAL_OMZET']:,.0f} sangat kecil. Pertimbangkan stop beli.",
                "PRIORITAS": "RENDAH",
                "HARGA_JML1": 0, "HARGA_JML2": 0,
                "MARGIN_JML1_PCT": 0, "MARGIN_JML2_PCT": 0,
                "HARGA_JML1_REC": 0, "HARGA_JML2_REC": 0, "HARGA_MODAL": 0,
            })

    # ── Reorder ──
    if not df_reorder.empty:
        for _, r in df_reorder.head(5).iterrows():
            hm = r.get("HARGA_MODAL", 0) or 0
            recs.append({
                "TYPE": "🛒 REORDER", "KODE": r["KODE"], "NAMA": _s(r, "NAMA"),
                "REKOMENDASI": (
                    f"Segera order {int(r['QTY_ORDER'])} PCS ke {_s(r,'SUPPLIER')} — "
                    f"stok {r['STOK_SISA']:.1f} < ROP {r['ROP']:.1f}. "
                    f"Estimasi harga modal: Rp {hm:,.0f}/PCS."
                ),
                "PRIORITAS": "TINGGI",
                "HARGA_JML1": 0, "HARGA_JML2": 0,
                "MARGIN_JML1_PCT": 0, "MARGIN_JML2_PCT": 0,
                "HARGA_JML1_REC": 0, "HARGA_JML2_REC": 0, "HARGA_MODAL": hm,
            })

    return recs


# ─────────────────────────────────────────────────────────────────
# SECTION 8 — EXPORT
# ─────────────────────────────────────────────────────────────────

def export_excel(df_fifo, df_margin, df_dead, df_reorder, df_audit, df_abc,
                 df_layer_alert=None) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book
        fmt_hdr   = wb.add_format({"bold": True, "bg_color": "#1a7f37",
                                    "font_color": "white", "border": 1, "align": "center"})
        fmt_rp    = wb.add_format({"num_format": "#,##0"})         # Rupiah: tanpa desimal
        fmt_qty   = wb.add_format({"num_format": "#,##0.0"})       # Qty: 1 desimal
        fmt_pct   = wb.add_format({"num_format": "0.0%"})          # Persen: 1 desimal
        fmt_int   = wb.add_format({"num_format": "#,##0"})         # Integer murni

        # Kolom yang termasuk kategori masing-masing format
        COLS_RP  = {"TOTAL","HPP_FIFO_UNIT","HPP_FIFO_TOTAL","MARGIN_RP","HARGA",
                    "HARGA_NET","HARGA_JUAL","HARGA_LAYER","NILAI_LAYER","AWAL_NILAI",
                    "MASUK_NILAI","KELUAR_NILAI","AKHIR_NILAI","NILAI_SISA","TOTAL_OMZET",
                    "HARGA_POKOK","HARGA_JML1","HARGA_JML2","HPP_FIFO","HPP_MEAN",
                    "HARGA_LAMA_PCS","HARGA_BARU_PCS","HARGA_JUAL_PCS","HARGA_MODAL","HARGA_JML1","HARGA_JML2"}
        COLS_PCT = {"MARGIN_PCT","MARGIN_BARU_PCT","OMZET_CUM_PCT","DISKON_PCT","MARGIN_JML1_PCT","MARGIN_JML2_PCT"}
        COLS_QTY = {"QTY","QTY_SISA","QTY_ORDER","AWAL_QTY","MASUK_QTY","KELUAR_QTY",
                    "AKHIR_QTY","STOK_SISA","TOTAL_QTY","QTY_SISA_LAMA","SAFETY_STOCK",
                    "ROP","AVG_DAILY","CALC_AKHIR","SELISIH"}

        def write_sheet(df, name, col_widths=None):
            if df is None or df.empty:
                return
            df.to_excel(writer, sheet_name=name, index=False)
            ws  = writer.sheets[name]
            for i, col in enumerate(df.columns):
                w = col_widths.get(col, 18) if col_widths else 18
                # Tentukan format angka berdasarkan nama kolom
                if col in COLS_RP:
                    ws.set_column(i, i, w, fmt_rp)
                elif col in COLS_PCT:
                    ws.set_column(i, i, w, fmt_pct)
                elif col in COLS_QTY:
                    ws.set_column(i, i, w, fmt_qty)
                else:
                    ws.set_column(i, i, w)
                ws.write(0, i, col, fmt_hdr)

        if df_fifo is not None and not df_fifo.empty:
            write_sheet(df_fifo, "FIFO Detail")
        if df_margin is not None and not df_margin.empty:
            write_sheet(df_margin, "Margin Alert")
        if df_dead is not None and not df_dead.empty:
            write_sheet(df_dead, "Dead Stock")
        if df_reorder is not None and not df_reorder.empty:
            write_sheet(df_reorder, "Smart Reorder")
        if df_audit is not None and not df_audit.empty:
            write_sheet(df_audit, "Audit Error")
        if df_abc is not None and not df_abc.empty:
            write_sheet(df_abc, "ABC Analysis")
        if df_layer_alert is not None and not df_layer_alert.empty:
            write_sheet(df_layer_alert, "FIFO Layer Alert")

    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────
# MAIN UI
# ─────────────────────────────────────────────────────────────────

def main():
    st.markdown('<p class="main-header">📦 SMART INVENTORY INTELLIGENCE SYSTEM</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">iPOS 5 Engine v1.5.0 — FIFO · ABC · Audit · Reorder · Margin Protection · 🔔 FIFO Layer Alert</p>',
                unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════
    # SESSION STATE — hasil analisis disimpan di sini agar tidak
    # hilang saat toggle / radio / widget lain diklik
    # ══════════════════════════════════════════════════════════════
    if "siis_cache" not in st.session_state:
        st.session_state.siis_cache = None

    # ── Sidebar ──────────────────────────────────────
    with st.sidebar:
        st.image("https://img.icons8.com/color/96/warehouse.png", width=64)
        st.markdown("### 📂 Upload File iPOS 5")
        f_beli   = st.file_uploader("📥 PEMBELIAN (.xlsx)",      type=["xlsx"], key="beli")
        f_jual   = st.file_uploader("📤 PENJUALAN (.xlsx)",      type=["xlsx"], key="jual")
        f_mutasi = st.file_uploader("🔄 MUTASI ITEM (.xlsx)",    type=["xlsx"], key="mutasi")
        f_master = st.file_uploader("📋 ITEM PERJUMLAH (.xlsx)", type=["xlsx"], key="master")

        st.markdown("---")
        st.markdown("### ⚙️ Pengaturan")
        lead_time   = st.number_input("Lead Time (hari)", 1, 30, 3)
        margin_min  = st.slider("Batas Margin Minimum (%)", 1, 30, 5)
        dead_days   = st.number_input("Dead Stock Threshold (hari)", 7, 180, 30)
        anomaly_pct = st.slider("Anomaly Harga Threshold (%)", 10, 100, 30)

        st.markdown("---")
        run_btn   = st.button("🚀 ANALISA SEKARANG", type="primary", use_container_width=True)
        if st.session_state.siis_cache is not None:
            reset_btn = st.button("🔄 Reset / Ganti Data", use_container_width=True)
        else:
            reset_btn = False

    # ── Reset ─────────────────────────────────────────
    if reset_btn:
        st.session_state.siis_cache = None
        st.rerun()

    # ── Jalankan analisis hanya saat tombol diklik ────
    if run_btn:
        _errors = []
        with st.spinner("⏳ Memuat & membersihkan data..."):
            try:
                _df_master = load_item_perjumlah(f_master) if f_master else pd.DataFrame()
            except Exception as e:
                _errors.append(f"ITEM PERJUMLAH: {e}"); _df_master = pd.DataFrame()
            try:
                _df_beli = load_pembelian(f_beli) if f_beli else pd.DataFrame()
            except Exception as e:
                _errors.append(f"PEMBELIAN: {e}"); _df_beli = pd.DataFrame()
            try:
                _df_jual = load_penjualan(f_jual) if f_jual else pd.DataFrame()
            except Exception as e:
                _errors.append(f"PENJUALAN: {e}"); _df_jual = pd.DataFrame()
            try:
                _df_mutasi = load_mutasi(f_mutasi) if f_mutasi else pd.DataFrame()
            except Exception as e:
                _errors.append(f"MUTASI: {e}"); _df_mutasi = pd.DataFrame()

        _ALERT_COLS = [
            "KODE","NAMA","SUPPLIER_LAMA","TGL_LAMA","QTY_SISA_LAMA",
            "HARGA_LAMA_PCS","SUPPLIER_BARU","TGL_BARU",
            "HARGA_BELI_ASLI","HARGA_BARU_PCS",
            "HARGA_JML1","HARGA_JML2","MARGIN_JML1_PCT","MARGIN_JML2_PCT",
            "HARGA_JUAL_PCS","MARGIN_BARU_PCT","ALERT_TYPE","SEVERITY",
            "SATUAN_BELI","KONV_USED","KETERANGAN"
        ]

        with st.spinner("⚙️ Menjalankan FIFO Engine..."):
            try:
                _df_fifo, _queues, _df_remaining, _df_layer_alert = run_fifo(
                    _df_beli, _df_jual, margin_min=float(margin_min), df_master=_df_master
                )
            except Exception as e:
                st.error(f"❌ FIFO Engine error: {e}")
                import traceback; st.code(traceback.format_exc())
                _df_fifo        = pd.DataFrame()
                _df_remaining   = pd.DataFrame(
                    columns=["KODE","QTY_SISA","HARGA_LAYER","NILAI_LAYER","SUPPLIER","TGL_MASUK"])
                _df_layer_alert = pd.DataFrame(columns=_ALERT_COLS)

        with st.spinner("📊 ABC & Dead Stock..."):
            _df_abc  = abc_analysis(_df_jual)
            _df_dead = detect_dead_stock(_df_jual, _df_remaining,
                                         _df_mutasi, int(dead_days), _df_master)
        with st.spinner("🛡️ Audit & Anomaly..."):
            _df_audit   = audit_stock(_df_mutasi)
            _df_anomaly = anomaly_harga(_df_beli, anomaly_pct / 100)
        with st.spinner("🛒 Smart Reorder..."):
            _df_reorder = smart_reorder(_df_jual, _df_remaining, _df_beli, int(lead_time))
        with st.spinner("📦 Validasi Konversi..."):
            _df_konversi = validate_konversi(_df_master, _df_fifo)

        _df_margin = (_df_fifo[_df_fifo["STATUS"] == "KRITIS"].copy()
                      if not _df_fifo.empty else pd.DataFrame())
        _jml_la  = len(_df_layer_alert) if not _df_layer_alert.empty else 0
        _jml_lk  = (len(_df_layer_alert[_df_layer_alert["SEVERITY"] == "KRITIS"])
                    if not _df_layer_alert.empty and "SEVERITY" in _df_layer_alert.columns else 0)
        _recs    = generate_recommendations(_df_fifo, _df_dead, _df_abc, _df_reorder,
                                            _df_layer_alert, margin_min=float(margin_min))
        _stok_fifo = _df_remaining["NILAI_LAYER"].sum() if not _df_remaining.empty else 0
        _stok_ipos = (_df_mutasi["AKHIR_NILAI"].sum()
                      if (not _df_mutasi.empty and "AKHIR_NILAI" in _df_mutasi.columns) else 0)

        # ── Simpan ke session state ────────────────────
        st.session_state.siis_cache = dict(
            df_fifo        = _df_fifo,
            df_remaining   = _df_remaining,
            df_layer_alert = _df_layer_alert,
            df_abc         = _df_abc,
            df_dead        = _df_dead,
            df_audit       = _df_audit,
            df_anomaly     = _df_anomaly,
            df_reorder     = _df_reorder,
            df_konversi    = _df_konversi,
            df_margin      = _df_margin,
            df_master      = _df_master,
            df_beli        = _df_beli,
            df_jual        = _df_jual,
            df_mutasi      = _df_mutasi,
            recs           = _recs,
            total_stok_fifo  = _stok_fifo,
            total_stok_ipos  = _stok_ipos,
            jml_layer_alert  = _jml_la,
            jml_layer_kritis = _jml_lk,
            margin_min       = float(margin_min),
            dead_days        = int(dead_days),
            anomaly_pct      = anomaly_pct,
            errors           = _errors,
        )
        st.success("✅ Analisis selesai! Data tersimpan — silakan gunakan semua tab & toggle bebas.")

    # ── Tampilkan panduan jika belum ada data ─────────
    if st.session_state.siis_cache is None:
        st.info("👈 Upload file Excel iPOS 5 di sidebar, lalu klik **ANALISA SEKARANG**.")
        _show_sample_ui()
        return

    # ══════════════════════════════════════════════════════════════
    # AMBIL DATA DARI SESSION STATE (aman saat toggle/widget diklik)
    # ══════════════════════════════════════════════════════════════
    C = st.session_state.siis_cache
    df_fifo        = C["df_fifo"]
    df_remaining   = C["df_remaining"]
    df_layer_alert = C["df_layer_alert"]
    df_abc         = C["df_abc"]
    df_dead        = C["df_dead"]
    df_audit       = C["df_audit"]
    df_anomaly     = C["df_anomaly"]
    df_reorder     = C["df_reorder"]
    df_konversi    = C["df_konversi"]
    df_margin      = C["df_margin"]
    df_master      = C["df_master"]
    df_beli        = C["df_beli"]
    df_jual        = C["df_jual"]
    df_mutasi      = C["df_mutasi"]
    recs             = C["recs"]
    total_stok_fifo  = C["total_stok_fifo"]
    total_stok_ipos  = C["total_stok_ipos"]
    jml_layer_alert  = C["jml_layer_alert"]
    jml_layer_kritis = C["jml_layer_kritis"]
    margin_min       = C["margin_min"]
    dead_days        = C["dead_days"]
    anomaly_pct      = C["anomaly_pct"]

    for err in C.get("errors", []):
        st.error(f"❌ {err}")

    # ── KPI Cards ─────────────────────────────────────
    st.markdown('<p class="section-title">📌 Ringkasan Eksekutif</p>', unsafe_allow_html=True)
    c1, c2, c3, c4, c5, c6, c7 = st.columns(7)

    selisih_stok  = total_stok_ipos - total_stok_fifo
    jml_kritis    = len(df_margin)  if not df_margin.empty  else 0
    jml_dead      = len(df_dead)    if not df_dead.empty    else 0
    nilai_dead    = df_dead["NILAI_SISA"].sum() if not df_dead.empty else 0
    jml_audit_err = len(df_audit)   if not df_audit.empty   else 0
    total_omzet   = df_jual["TOTAL"].sum() if not df_jual.empty else 0

    with c1:
        st.metric("💰 Stok (FIFO)", f"Rp {total_stok_fifo:,.0f}")
    with c2:
        st.metric("📦 Stok (iPOS)", f"Rp {total_stok_ipos:,.0f}")
    with c3:
        pct_selisih = (selisih_stok / total_stok_ipos * 100) if total_stok_ipos > 0 else 0
        st.metric("📊 Selisih", f"Rp {selisih_stok:,.0f}",
                  delta=f"{pct_selisih:+.1f}%",
                  delta_color="inverse" if pct_selisih < -50 else "normal")
    with c4:
        st.metric("⚠️ Kr. Margin", f"{jml_kritis:,}")
    with c5:
        st.metric("🧊 Dead Stock", f"{jml_dead:,}")
    with c6:
        st.metric("📈 Omzet", f"Rp {total_omzet:,.0f}")
    with c7:
        delta_txt = f"{jml_layer_kritis} KRITIS" if jml_layer_kritis > 0 else None
        st.metric("🔔 Layer Alert", f"{jml_layer_alert:,}", delta=delta_txt,
                  delta_color="inverse" if jml_layer_kritis > 0 else "normal")

    if abs(selisih_stok) > 0:
        st.caption("💡 **Stok (FIFO)** = Stok dari file pembelian upload. **Stok (iPOS)** = Stok real di sistem (termasuk warisan lama). **Selisih** = Stok lama yang tidak terlacak.")

    st.markdown("---")

    # ── Tabs ────────────────────────────────────────
    tab1, tab2, tab2b, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        "📊 Dashboard", "💹 FIFO & Margin", "🔔 FIFO Layer Alert", "📈 ABC Analysis",
        "🧊 Dead Stock", "🛡️ Audit", "🛒 Reorder",
        "📦 Konversi", "🧠 Rekomendasi"
    ])

    # ── TAB 1: Dashboard ────────────────────────────
    with tab1:
        col1, col2 = st.columns(2)

        with col1:
            st.markdown('<p class="section-title">📊 Distribusi Kelas ABC</p>', unsafe_allow_html=True)
            if not df_abc.empty:
                abc_grp = df_abc.groupby("KELAS_ABC")["TOTAL_OMZET"].sum().reset_index()
                fig = px.pie(abc_grp, names="KELAS_ABC", values="TOTAL_OMZET",
                             color="KELAS_ABC",
                             color_discrete_map={"A": "#1a7f37", "B": "#f9a825", "C": "#e53935"},
                             hole=0.4, title="Omzet per Kelas ABC")
                fig.update_traces(textposition="inside", textinfo="percent+label")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Data penjualan kosong.")

        with col2:
            st.markdown('<p class="section-title">🏭 Top Supplier (Pembelian)</p>', unsafe_allow_html=True)
            if not df_beli.empty:
                sup_grp = df_beli.groupby("SUPPLIER")["TOTAL"].sum().nlargest(10).reset_index()
                fig2 = px.bar(sup_grp, x="TOTAL", y="SUPPLIER", orientation="h",
                              color="TOTAL", color_continuous_scale="Greens",
                              title="Nilai Pembelian per Supplier")
                fig2.update_layout(yaxis=dict(autorange="reversed"), coloraxis_showscale=False)
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.warning("Data pembelian kosong.")

        col3, col4 = st.columns(2)
        with col3:
            st.markdown('<p class="section-title">📈 Top 10 Item Omzet</p>', unsafe_allow_html=True)
            if not df_abc.empty:
                top10 = df_abc.nlargest(10, "TOTAL_OMZET")
                fig3 = px.bar(top10, x="TOTAL_OMZET", y="NAMA", orientation="h",
                              color="KELAS_ABC",
                              color_discrete_map={"A": "#1a7f37", "B": "#f9a825", "C": "#e53935"},
                              title="Top 10 Item by Omzet")
                fig3.update_layout(yaxis=dict(autorange="reversed"))
                st.plotly_chart(fig3, use_container_width=True)

        with col4:
            st.markdown('<p class="section-title">⚠️ Status Margin</p>', unsafe_allow_html=True)
            if not df_fifo.empty:
                status_grp = df_fifo.groupby("STATUS").size().reset_index(name="COUNT")
                fig4 = px.pie(status_grp, names="STATUS", values="COUNT",
                              color="STATUS",
                              color_discrete_map={"AMAN": "#1a7f37", "KRITIS": "#e53935"},
                              hole=0.4, title="Status Margin Transaksi")
                st.plotly_chart(fig4, use_container_width=True)

    # ── TAB 2: FIFO & Margin ────────────────────────
    with tab2:
        st.markdown('<p class="section-title">💹 Detail FIFO & Margin</p>', unsafe_allow_html=True)
        if not df_fifo.empty:
            show_cols = ["KODE","NAMA","QTY","TOTAL","HPP_FIFO_UNIT","HPP_FIFO_TOTAL","MARGIN_RP","MARGIN_PCT","STATUS","FIFO_WARNING"]
            show_cols = [c for c in show_cols if c in df_fifo.columns]
            disp = df_fifo[show_cols].copy()

            def color_status(v):
                return "background-color:#ffcccc" if v == "KRITIS" else "background-color:#ccffcc"

            st.dataframe(
                disp.style.applymap(color_status, subset=["STATUS"])
                    .format({c: fmt for c, fmt in {
                        "QTY":"{:,.1f}","TOTAL":"Rp {:,.0f}",
                        "HPP_FIFO_UNIT":"Rp {:,.0f}","HPP_FIFO_TOTAL":"Rp {:,.0f}",
                        "MARGIN_RP":"Rp {:,.0f}","MARGIN_PCT":"{:.1f}%"
                    }.items() if c in disp.columns}),
                use_container_width=True, height=400
            )

            st.markdown(f"**🔴 Transaksi KRITIS:** {jml_kritis} | "
                        f"**🟢 AMAN:** {len(df_fifo) - jml_kritis}")

            if not df_remaining.empty:
                st.markdown('<p class="section-title">📦 Sisa Stok FIFO (dalam PCS)</p>', unsafe_allow_html=True)
                rem_grp = df_remaining.groupby("KODE").agg(
                    QTY_SISA=("QTY_SISA","sum"), NILAI=("NILAI_LAYER","sum")
                ).reset_index()
                st.dataframe(rem_grp.style.format({"QTY_SISA":"{:,.1f}","NILAI":"Rp {:,.0f}"}),
                             use_container_width=True, height=300)
        else:
            st.warning("Data FIFO kosong. Pastikan file PEMBELIAN dan PENJUALAN sudah diupload.")

    # ── TAB 2b: FIFO Layer Alert ─────────────────────
    with tab2b:
        st.markdown('<p class="section-title">🔔 Peringatan Konflik Lapisan FIFO</p>',
                    unsafe_allow_html=True)
        st.caption(
            "Sistem mendeteksi saat pembelian baru masuk **sementara stok lama belum habis**, "
            "dan harga modal baru berpotensi merusak margin. "
            "Semua perhitungan sudah dinormalisasi ke **satuan PCS**."
        )

        if not df_layer_alert.empty and "SEVERITY" in df_layer_alert.columns:
            n_kritis  = len(df_layer_alert[df_layer_alert["SEVERITY"] == "KRITIS"])
            n_waspada = len(df_layer_alert[df_layer_alert["SEVERITY"] == "WASPADA"])
            col_k, col_w = st.columns(2)
            with col_k:
                if n_kritis > 0:
                    st.error(f"🔴 **{n_kritis} item** — Modal baru MELEBIHI harga jual (rugi pasti!)")
                else:
                    st.success("✅ Tidak ada item dengan modal > harga jual")
            with col_w:
                if n_waspada > 0:
                    st.warning(f"🟡 **{n_waspada} item** — Margin di bawah target atau modal naik signifikan")
                else:
                    st.success("✅ Semua kenaikan modal masih dalam batas aman")

            st.markdown("---")

            sev_filter = st.radio(
                "Tampilkan:", ["Semua", "🔴 KRITIS saja", "🟡 WASPADA saja"],
                horizontal=True
            )
            if sev_filter == "🔴 KRITIS saja":
                disp_alert = df_layer_alert[df_layer_alert["SEVERITY"] == "KRITIS"]
            elif sev_filter == "🟡 WASPADA saja":
                disp_alert = df_layer_alert[df_layer_alert["SEVERITY"] == "WASPADA"]
            else:
                disp_alert = df_layer_alert

            def color_severity(row):
                sev = row.get("SEVERITY", "")
                if sev == "KRITIS":
                    return ["background-color:#ffcccc"] * len(row)
                elif sev == "WASPADA":
                    return ["background-color:#fff9c4"] * len(row)
                return [""] * len(row)

            show_cols = [
                "KODE","NAMA","ALERT_TYPE","SEVERITY",
                "SATUAN_BELI","KONV_USED",
                "SUPPLIER_LAMA","TGL_LAMA","QTY_SISA_LAMA","HARGA_LAMA_PCS",
                "SUPPLIER_BARU","TGL_BARU",
                "HARGA_BELI_ASLI","HARGA_BARU_PCS",
                "HARGA_JML1","MARGIN_JML1_PCT",
                "HARGA_JML2","MARGIN_JML2_PCT",
                "HARGA_JUAL_PCS","MARGIN_BARU_PCT","KETERANGAN"
            ]
            show_cols = [c for c in show_cols if c in disp_alert.columns]

            fmt_alert = {
                "HARGA_LAMA_PCS"  : "Rp {:,.0f}",
                "HARGA_BELI_ASLI" : "Rp {:,.0f}",
                "HARGA_BARU_PCS"  : "Rp {:,.0f}",
                "HARGA_JML1"      : "Rp {:,.0f}",
                "HARGA_JML2"      : "Rp {:,.0f}",
                "HARGA_JUAL_PCS"  : "Rp {:,.0f}",
                "QTY_SISA_LAMA"   : "{:,.1f}",
                "MARGIN_BARU_PCT" : "{:.1f}%",
                "MARGIN_JML1_PCT" : "{:.1f}%",
                "MARGIN_JML2_PCT" : "{:.1f}%",
                "KONV_USED"       : "{:.0f}×",
            }

            st.caption(
                "ℹ️ **HARGA_BELI_ASLI** = harga per satuan beli (misal per SLOF/DUS, sebelum dikonversi). "
                "**HARGA_BARU_PCS** = harga setelah ÷ KONV_USED (sudah per PCS)."
            )
            st.dataframe(
                disp_alert[show_cols].style
                    .apply(color_severity, axis=1)
                    .format({k: v for k, v in fmt_alert.items() if k in disp_alert.columns}),
                use_container_width=True, height=420
            )

            # ── Grafik ──────────────────────────────
            st.markdown("---")
            col_ch1, col_ch2 = st.columns(2)

            with col_ch1:
                st.markdown("##### 📊 Modal vs Harga Jual per Item (per PCS)")
                chart_data = df_layer_alert[["KODE","HARGA_BARU_PCS","HARGA_JUAL_PCS","HARGA_LAMA_PCS"]].copy()
                chart_data = chart_data.drop_duplicates("KODE").head(15)
                if not chart_data.empty:
                    fig_bar = go.Figure()
                    fig_bar.add_bar(name="Modal Baru/PCS", x=chart_data["KODE"], y=chart_data["HARGA_BARU_PCS"], marker_color="#e53935")
                    fig_bar.add_bar(name="Harga Jual/PCS", x=chart_data["KODE"], y=chart_data["HARGA_JUAL_PCS"], marker_color="#1a7f37")
                    fig_bar.add_bar(name="Modal Lama/PCS", x=chart_data["KODE"], y=chart_data["HARGA_LAMA_PCS"], marker_color="#fb8c00")
                    fig_bar.update_layout(barmode="group", height=320, margin=dict(t=10, b=40),
                                          legend=dict(orientation="h", yanchor="bottom", y=1.02))
                    st.plotly_chart(fig_bar, use_container_width=True)

            with col_ch2:
                st.markdown("##### 🥧 Komposisi Alert per Tipe")
                type_counts = df_layer_alert["ALERT_TYPE"].value_counts().reset_index()
                type_counts.columns = ["TIPE", "JUMLAH"]
                fig_pie = px.pie(type_counts, names="TIPE", values="JUMLAH",
                                 color_discrete_sequence=["#e53935","#fb8c00","#f9a825"], hole=0.4)
                fig_pie.update_traces(textposition="inside", textinfo="percent+label")
                fig_pie.update_layout(height=320, margin=dict(t=10))
                st.plotly_chart(fig_pie, use_container_width=True)

            # ── Detail expandable ───────────────────
            st.markdown("---")
            st.markdown("##### 🔍 Detail Per Item")
            for kode_item, grp_item in df_layer_alert.groupby("KODE"):
                nama_item = grp_item["NAMA"].iloc[0] if "NAMA" in grp_item.columns else kode_item
                severity_max = "KRITIS" if "KRITIS" in grp_item["SEVERITY"].values else "WASPADA"
                icon = "🔴" if severity_max == "KRITIS" else "🟡"
                with st.expander(f"{icon} {kode_item} — {nama_item} ({len(grp_item)} alert)"):
                    for _, ar in grp_item.iterrows():
                        st.markdown(f"**{ar['ALERT_TYPE']}**")
                        # Baris 1: Modal Lama, Modal Baru, Konversi
                        cols = st.columns(3)
                        cols[0].metric("Modal Lama/PCS", f"Rp {ar['HARGA_LAMA_PCS']:,.0f}")
                        cols[1].metric("Modal Baru/PCS", f"Rp {ar['HARGA_BARU_PCS']:,.0f}",
                                       delta=f"Rp {ar['HARGA_BARU_PCS']-ar['HARGA_LAMA_PCS']:+,.0f}",
                                       delta_color="inverse")
                        konv_txt = str(ar['KONV_USED']) if 'KONV_USED' in ar.index and pd.notna(ar['KONV_USED']) else "?"
                        sat_txt  = str(ar['SATUAN_BELI']) if 'SATUAN_BELI' in ar.index and pd.notna(ar['SATUAN_BELI']) else ""
                        cols[2].metric("Konversi Satuan", f"{konv_txt}× {sat_txt}")

                        # Baris 2: Harga Jual 1 (Eceran) & Harga Jual 2 (Grosir) vs Modal Baru
                        hm  = float(ar.get("HARGA_BARU_PCS", 0) or 0)
                        h1  = float(ar.get("HARGA_JML1", 0) or 0)
                        h2  = float(ar.get("HARGA_JML2", 0) or 0)
                        m1  = float(ar.get("MARGIN_JML1_PCT", 0) or 0)
                        m2  = float(ar.get("MARGIN_JML2_PCT", 0) or 0)
                        st.markdown("**📊 Harga Jual vs Modal Baru:**")
                        col_hj1, col_hj2 = st.columns(2)
                        with col_hj1:
                            status1 = "🟢" if h1 > hm else "🔴"
                            label1  = "✅ Di atas modal" if h1 > hm else "❌ DI BAWAH MODAL — RUGI!"
                            if h1 > 0:
                                col_hj1.metric(
                                    f"{status1} Harga Jual 1 (Eceran)",
                                    f"Rp {h1:,.0f}",
                                    f"Margin {m1:+.1f}% vs modal baru",
                                    delta_color="normal" if m1 > 0 else "inverse"
                                )
                                st.caption(label1)
                            else:
                                col_hj1.warning("HJ1 tidak ada di master")
                        with col_hj2:
                            status2 = "🟢" if h2 > hm else "🔴"
                            label2  = "✅ Di atas modal" if h2 > hm else "❌ DI BAWAH MODAL — RUGI!"
                            if h2 > 0:
                                col_hj2.metric(
                                    f"{status2} Harga Jual 2 (Grosir)",
                                    f"Rp {h2:,.0f}",
                                    f"Margin {m2:+.1f}% vs modal baru",
                                    delta_color="normal" if m2 > 0 else "inverse"
                                )
                                st.caption(label2)
                            else:
                                col_hj2.warning("HJ2 tidak ada di master")

                        # Rekomendasi harga jual jika di bawah modal
                        rec_lines = []
                        if h1 > 0 and h1 <= hm and hm > 0:
                            rec_hj1 = round(hm * (1 + float(margin_min) / 100))
                            rec_lines.append(f"💡 **Rekomendasi HJ1:** Naikkan ke **Rp {rec_hj1:,.0f}** (modal+{margin_min}%)")
                        if h2 > 0 and h2 <= hm and hm > 0:
                            rec_hj2 = round(hm * (1 + float(margin_min) * 0.7 / 100))
                            rec_lines.append(f"💡 **Rekomendasi HJ2:** Naikkan ke **Rp {rec_hj2:,.0f}** (modal+{margin_min*0.7:.0f}% grosir)")
                        if rec_lines:
                            st.warning("\n\n".join(rec_lines))

                        st.caption(f"ℹ️ {ar['KETERANGAN']}")
                        st.markdown("---")

            # Panduan tindakan
            st.markdown("---")
            st.markdown("#### 💡 Panduan Tindakan")
            st.markdown(f"""
| Kondisi | Tindakan yang Disarankan |
|---------|--------------------------|
| 🔴 Modal > Harga Jual | Tunda pembelian atau naikkan harga jual sebelum terima barang baru |
| 🟡 Margin < {margin_min:.0f}% | Negosiasi harga ke supplier, atau review harga jual |
| 🟠 Modal naik >15% | Cek apakah kenaikan permanen / temporer, pertimbangkan supplier alternatif |
""")

            # Export peringatan
            if st.button("📥 Export Peringatan FIFO ke Excel", use_container_width=True):
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                    df_layer_alert.to_excel(writer, sheet_name="FIFO Layer Alert", index=False)
                st.download_button(
                    "⬇️ Download FIFO_Layer_Alert.xlsx",
                    data=buf.getvalue(),
                    file_name=f"FIFO_Layer_Alert_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        else:
            st.success(
                "✅ Tidak ada konflik lapisan FIFO terdeteksi. "
                "Semua pembelian baru masuk saat stok lama sudah habis, "
                "atau margin masih di atas target."
            )

    # ── TAB 3: ABC Analysis ─────────────────────────
    with tab3:
        st.markdown('<p class="section-title">📈 ABC Analysis</p>', unsafe_allow_html=True)
        if not df_abc.empty:
            col_a, col_b, col_c = st.columns(3)
            for kls, col in zip(["A","B","C"], [col_a, col_b, col_c]):
                sub = df_abc[df_abc["KELAS_ABC"] == kls]
                col.metric(f"Kelas {kls}", f"{len(sub)} item", f"Rp {sub['TOTAL_OMZET'].sum():,.0f}")

            st.dataframe(
                df_abc.style.format({
                    "TOTAL_QTY":"{:,.1f}","TOTAL_OMZET":"Rp {:,.0f}","OMZET_CUM_PCT":"{:.1f}%"
                }).applymap(
                    lambda v: "background-color:#ccffcc" if v=="A" else
                              ("background-color:#fffde7" if v=="B" else "background-color:#ffcccc"),
                    subset=["KELAS_ABC"]
                ),
                use_container_width=True, height=400
            )
        else:
            st.warning("Data penjualan kosong.")

    # ── TAB 4: Dead Stock ───────────────────────────
    with tab4:
        st.markdown('<p class="section-title">🧊 Dead Stock</p>', unsafe_allow_html=True)
        if not df_dead.empty:
            st.error(f"⚠️ Ditemukan **{len(df_dead)} item** dead stock "
                     f"(tidak terjual ≥ {dead_days} hari)")
            st.dataframe(
                df_dead.style.format({
                    "STOK_SISA":"{:,.1f}","NILAI_SISA":"Rp {:,.0f}"
                }),
                use_container_width=True, height=350
            )            

            top_dead = df_dead.nlargest(15,"NILAI_SISA").copy()
            # Buat label singkat: NAMA (maks 30 char)
            top_dead["LABEL"] = top_dead.apply(
                lambda r: (r["NAMA"][:28] + "…") if len(str(r["NAMA"])) > 30 else r["NAMA"],
                axis=1
            )
            fig_dead = px.bar(top_dead, x="NILAI_SISA", y="LABEL", orientation="h",
                              title="Top 15 Dead Stock by Nilai",
                              labels={"NILAI_SISA": "Nilai Sisa (Rp)", "LABEL": "Nama Item"},
                              color_discrete_sequence=["#e53935"])
            fig_dead.update_layout(yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig_dead, use_container_width=True)
        else:
            st.success("✅ Tidak ada dead stock terdeteksi.")

    # ── TAB 5: Audit ────────────────────────────────
    with tab5:
        st.markdown('<p class="section-title">🛡️ Audit Error Stok</p>', unsafe_allow_html=True)
        if not df_audit.empty:
            st.error(f"❌ Ditemukan **{len(df_audit)} item** dengan selisih stok!")
            st.dataframe(df_audit.style.format({
                "AWAL_QTY":"{:,.1f}","MASUK_QTY":"{:,.1f}",
                "KELUAR_QTY":"{:,.1f}","AKHIR_QTY":"{:,.1f}",
                "CALC_AKHIR":"{:,.1f}","SELISIH":"{:,.2f}"
            }), use_container_width=True)
        else:
            st.success("✅ Tidak ada error stok terdeteksi.")

        st.markdown('<p class="section-title">💣 Anomaly Harga Beli</p>', unsafe_allow_html=True)
        if not df_anomaly.empty:
            st.warning(f"⚠️ Ditemukan **{len(df_anomaly)} transaksi** dengan harga anomali (>{anomaly_pct}% dari rata-rata)!")
            st.dataframe(df_anomaly.style.format({
                "HARGA_NET":"Rp {:,.0f}","HARGA_MEAN":"Rp {:,.0f}","HARGA_STD":"Rp {:,.0f}"
            }), use_container_width=True)
        else:
            st.success("✅ Tidak ada anomali harga terdeteksi.")

    # ── TAB 6: Reorder ──────────────────────────────
    with tab6:
        st.markdown('<p class="section-title">🛒 Smart Reorder Recommendation</p>', unsafe_allow_html=True)
        if not df_reorder.empty:
            st.warning(f"🚨 **{len(df_reorder)} item** perlu di-reorder segera!")

            # Toggle global: sembunyikan/tampilkan harga modal di template WA
            col_tog1, col_tog2 = st.columns([1, 2])
            with col_tog1:
                show_modal_wa = st.toggle("💰 Tampilkan Harga Modal di WA", value=False,
                                           help="ON = harga modal muncul di pesan WA | OFF = harga disembunyikan")
            with col_tog2:
                st.caption("📱 Gunakan input nomor HP supplier untuk kirim langsung ke WhatsApp")

            st.markdown("---")

            for sup, grp in df_reorder.groupby("SUPPLIER"):
                with st.expander(f"📦 Supplier: {sup} ({len(grp)} item)", expanded=True):
                    cols_show = [c for c in
                        ["KODE","NAMA","STOK_SISA","AVG_DAILY","SAFETY_STOCK","ROP","QTY_ORDER","HARGA_MODAL"]
                        if c in grp.columns]
                    show = grp[cols_show].copy()
                    fmt = {
                        "STOK_SISA"    : "{:,.1f}",
                        "AVG_DAILY"    : "{:,.1f}",
                        "SAFETY_STOCK" : "{:,.1f}",
                        "ROP"          : "{:,.1f}",
                        "QTY_ORDER"    : "{:,}",
                        "HARGA_MODAL"  : "Rp {:,.0f}",
                    }
                    st.dataframe(
                        show.style.format({k: v for k, v in fmt.items() if k in show.columns}),
                        use_container_width=True
                    )

                    # ── Input nomor WA supplier ──────────────────────
                    wa_col1, wa_col2 = st.columns([2, 1])
                    with wa_col1:
                        sup_clean = "".join(c for c in sup if c.isalnum())[:12]
                        default_no = "628"
                        wa_number  = st.text_input(
                            f"📞 No. HP Supplier *{sup}* (format: 628xxx tanpa +)",
                            value=default_no,
                            key=f"wa_no_{sup_clean}",
                            placeholder="Contoh: 6281234567890"
                        )

                    # ── Template WA ──────────────────────────────────
                    tgl_str = datetime.now().strftime("%d/%m/%Y")
                    wa_lines = [
                        f"*Order Request — {sup}*",
                        f"_Tanggal: {tgl_str}_",
                        "",
                        "*Detail Item:*",
                    ]
                    for _, r in grp.iterrows():
                        nama  = r.get("NAMA", r["KODE"])
                        kode  = r["KODE"]
                        qty   = int(r["QTY_ORDER"])
                        modal = r.get("HARGA_MODAL", 0) or 0
                        if show_modal_wa and modal > 0:
                            wa_lines.append(
                                f"• {nama} ({kode}) — {qty} PCS | Modal: Rp {modal:,.0f}"
                            )
                        else:
                            wa_lines.append(f"• {nama} ({kode}) — {qty} PCS")
                    wa_lines += ["", "_Mohon konfirmasi ketersediaan & harga. Terima kasih 🙏_"]
                    wa_text = "\n".join(wa_lines)

                    st.text_area(
                        "📱 Template WA (copy & kirim ke supplier):",
                        wa_text, height=200, key=f"wa_{sup_clean}"
                    )

                    # ── Tombol Kirim ke WA ───────────────────────────
                    with wa_col2:
                        st.markdown("<br>", unsafe_allow_html=True)
                        wa_encoded = urllib.parse.quote(wa_text)
                        wa_no_clean = wa_number.replace("+","").replace(" ","").replace("-","")
                        if wa_no_clean and len(wa_no_clean) >= 9:
                            wa_link = f"https://wa.me/{wa_no_clean}?text={wa_encoded}"
                            st.markdown(
                                f'<a href="{wa_link}" target="_blank">'
                                f'<button style="background:#25D366;color:white;border:none;'
                                f'padding:10px 20px;border-radius:8px;font-size:1rem;'
                                f'cursor:pointer;width:100%;font-weight:bold;">'
                                f'📲 Kirim ke WhatsApp</button></a>',
                                unsafe_allow_html=True
                            )
                        else:
                            st.warning("⚠️ Isi nomor HP supplier dulu")
        else:
            st.success("✅ Semua item stok aman, tidak ada reorder diperlukan.")

        # ── TAB 7: Konversi ─────────────────────────────
    with tab7:
        st.markdown('<p class="section-title">📦 Validasi Konversi & Harga Master</p>', unsafe_allow_html=True)
        
        if not df_master.empty:
            # ★ TAMBAHAN: Deteksi harga ngaco (di luar rentang wajar)
            harga_cols = [c for c in ["HARGA_POKOK","HARGA_JML1","HARGA_JML2"] if c in df_master.columns]
            
            # Flagging: True kalau harga < 100 atau > 10.000.000 (kemungkinan typo/barcode)
            df_master_display = df_master.copy()
            for hc in harga_cols:
                if hc in df_master_display.columns:
                    df_master_display[f"{hc}_STATUS"] = df_master_display[hc].apply(
                        lambda x: "❌ NGACO" if (pd.notna(x) and (x < 100 or x > 10000000)) else "✅ OK"
                    )
            
            # Hitung jumlah ngaco
            status_cols = [f"{hc}_STATUS" for hc in harga_cols]
            df_master_display["TOTAL_NGACO"] = df_master_display[status_cols].apply(
                lambda row: sum(1 for val in row if val == "❌ NGACO"), axis=1
            )
            
            jml_ngaco = len(df_master_display[df_master_display["TOTAL_NGACO"] > 0])
            
            if jml_ngaco > 0:
                st.error(f"⚠️ Ditemukan **{jml_ngaco} item** dengan harga tidak wajar (kemungkinan typo atau barcode nyangkut)!")
                st.caption("Rentang wajar: Rp 100 - Rp 10.000.000. Diluar itu ditandai ❌ NGACO.")
            
            # Filter tampilan: pilih mau lihat yang OK saja atau yang NGACO saja
            filter_view = st.radio("Tampilkan Master:", ["Semua", "❌ Yang Ngaco Saja", "✅ Yang OK Saja"], horizontal=True)
            
            if filter_view == "❌ Yang Ngaco Saja":
                df_view = df_master_display[df_master_display["TOTAL_NGACO"] > 0]
            elif filter_view == "✅ Yang OK Saja":
                df_view = df_master_display[df_master_display["TOTAL_NGACO"] == 0]
            else:
                df_view = df_master_display
                
            show_master = [c for c in ["KODE","NAMA","SATUAN","SAT_BESAR","KONV_BESAR","HARGA_POKOK","HARGA_JML1","HARGA_JML2"] if c in df_view.columns]
            
            # Tambahkan kolom status ke tampilan
            status_show = [f"{hc}_STATUS" for hc in harga_cols if f"{hc}_STATUS" in df_view.columns]
            final_show = show_master + status_show
            
            st.dataframe(df_view[final_show].head(200).style.format(
                {c: "Rp {:,.0f}" for c in ["HARGA_POKOK","HARGA_JML1","HARGA_JML2"] if c in df_view.columns}
            ), use_container_width=True, height=400)

        # Validasi FIFO vs Grosir (hanya untuk yang harganya wajar/OK)
        if not df_konversi.empty:
            st.markdown("---")
            st.error(f"⚠️ {len(df_konversi)} item dengan Harga Grosir ≤ HPP FIFO (Setelah filter data wajar)")
            st.dataframe(df_konversi.style.format({
                "HARGA_JML1":"Rp {:,.0f}","HPP_FIFO":"Rp {:,.0f}","HARGA_POKOK":"Rp {:,.0f}"
            }), use_container_width=True)
        elif not df_master.empty and jml_ngaco == 0:
            st.success("✅ Semua harga master valid dan wajar (> HPP FIFO).")
    # ── TAB 8: Rekomendasi ──────────────────────────
    with tab8:
        st.markdown('<p class="section-title">🧠 Smart Recommendations</p>', unsafe_allow_html=True)
        if recs:
            for r in recs:
                icon = "🔴" if r["PRIORITAS"] == "TINGGI" else ("🟡" if r["PRIORITAS"] == "SEDANG" else "🟢")
                with st.expander(f"{icon} {r['TYPE']} — {r.get('NAMA','')[:45]}", expanded=r["PRIORITAS"]=="TINGGI"):
                    st.markdown(f"**Kode:** `{r['KODE']}`")
                    st.markdown(f"**Rekomendasi:** {r['REKOMENDASI']}")

                    h1   = r.get("HARGA_JML1", 0) or 0
                    h2   = r.get("HARGA_JML2", 0) or 0
                    m1   = r.get("MARGIN_JML1_PCT", 0) or 0
                    m2   = r.get("MARGIN_JML2_PCT", 0) or 0
                    hrec1= r.get("HARGA_JML1_REC", 0) or 0
                    hrec2= r.get("HARGA_JML2_REC", 0) or 0
                    hmod = r.get("HARGA_MODAL", 0) or 0

                    if h1 > 0 or h2 > 0:
                        st.markdown("**📊 Harga Jual Saat Ini vs Modal:**")
                        col_h1, col_h2 = st.columns(2)
                        with col_h1:
                            clr1 = "🟢" if m1 > 0 else "🔴"
                            st.metric(f"{clr1} Harga Jual 1 (Eceran)",
                                      f"Rp {h1:,.0f}" if h1 > 0 else "-",
                                      f"Margin {m1:.1f}%" if h1 > 0 else None,
                                      delta_color="normal" if m1 > 0 else "inverse")
                        with col_h2:
                            clr2 = "🟢" if m2 > 0 else "🔴"
                            st.metric(f"{clr2} Harga Jual 2 (Grosir)",
                                      f"Rp {h2:,.0f}" if h2 > 0 else "-",
                                      f"Margin {m2:.1f}%" if h2 > 0 else None,
                                      delta_color="normal" if m2 > 0 else "inverse")

                    # Tampilkan harga rekomendasi jika ada
                    if hrec1 > 0 or hrec2 > 0:
                        st.markdown("---")
                        st.error("⚠️ **Harga jual di bawah modal FIFO — Perlu diperbaiki segera!**")
                        st.markdown("**💡 Harga Jual Rekomendasi (min. margin aman):**")
                        col_r1, col_r2 = st.columns(2)
                        with col_r1:
                            if hrec1 > 0:
                                selisih1 = hrec1 - h1
                                st.metric(
                                    "🎯 Rekomendasi HJ1 (Eceran)",
                                    f"Rp {hrec1:,.0f}",
                                    f"Naik Rp {selisih1:,.0f} dari Rp {h1:,.0f}",
                                    delta_color="normal"
                                )
                                st.caption(f"Margin ≈ {margin_min:.0f}% di atas modal Rp {hmod:,.0f}/PCS")
                        with col_r2:
                            if hrec2 > 0:
                                selisih2 = hrec2 - h2
                                st.metric(
                                    "🎯 Rekomendasi HJ2 (Grosir)",
                                    f"Rp {hrec2:,.0f}",
                                    f"Naik Rp {selisih2:,.0f} dari Rp {h2:,.0f}",
                                    delta_color="normal"
                                )
                                st.caption(f"Margin ≈ {margin_min*0.7:.0f}% di atas modal Rp {hmod:,.0f}/PCS")

                    st.caption(f"**Prioritas:** {icon} {r['PRIORITAS']}")
        else:
            st.success("✅ Tidak ada rekomendasi kritis.")

    # ── Export Section ──────────────────────────────
    st.markdown("---")
    st.markdown('<p class="section-title">📤 Export Laporan</p>', unsafe_allow_html=True)
    col_exp1, col_exp2 = st.columns(2)

    with col_exp1:
        if st.button("📊 Export Excel (Multi-Sheet)", use_container_width=True):
            with st.spinner("Membuat file Excel..."):
                excel_bytes = export_excel(df_fifo, df_margin, df_dead, df_reorder, df_audit, df_abc,
                                           df_layer_alert)
            st.download_button(
                label="⬇️ Download SIIS_Report.xlsx",
                data=excel_bytes,
                file_name=f"SIIS_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    with col_exp2:
        summary = f"""SIIS REPORT v1.2.1 — {datetime.now().strftime('%d/%m/%Y %H:%M')}
======================================
Nilai Stok (FIFO) : Rp {total_stok_fifo:,.0f}
Nilai Stok (iPOS) : Rp {total_stok_ipos:,.0f}
Selisih Stok      : Rp {selisih_stok:,.0f} ({(selisih_stok/total_stok_ipos*100) if total_stok_ipos > 0 else 0:+.1f}%)
Total Omzet       : Rp {total_omzet:,.0f}
Transaksi Kritis  : {jml_kritis}
Dead Stock Items  : {jml_dead}
Potensi Kerugian  : Rp {nilai_dead:,.0f}
Audit Error       : {jml_audit_err}
Item Perlu Order  : {len(df_reorder) if not df_reorder.empty else 0}
FIFO Layer Alert  : {jml_layer_alert} ({jml_layer_kritis} kritis)
======================================
"""
        st.download_button(
            label="📄 Download Ringkasan (.txt)",
            data=summary.encode("utf-8"),
            file_name=f"SIIS_Summary_{datetime.now().strftime('%Y%m%d')}.txt",
            mime="text/plain",
            use_container_width=True,
        )


def _show_sample_ui():
    """Placeholder UI when no files uploaded."""
    st.markdown('<p class="section-title">📋 Panduan Penggunaan</p>', unsafe_allow_html=True)
    st.markdown("""
    **SIIS – iPOS 5 Engine** membutuhkan 4 file Excel dari iPOS 5:

    | File | Fungsi |
    |------|--------|
    | `PEMBELIAN_IPOS.xlsx` | Data pembelian per supplier |
    | `PENJUALAN_IPOS.xlsx` | Data penjualan periode |
    | `MUTASI_ITEM_IPOS.xlsx` | Mutasi stok (awal/masuk/keluar/akhir) |
    | `ITEM_PERJUMLAH_IPOS.xlsx` | Master item + konversi satuan |

    **Fitur utama:**
    - 🔄 **FIFO Engine** — HPP real berdasarkan antrean pembelian (konversi SLOF/DUS/LSN → PCS otomatis)
    - 🔔 **FIFO Layer Alert** — Peringatan otomatis saat beli baru masuk, stok lama belum habis & modal naik
    - 💰 **Margin Protection** — Flag transaksi margin < threshold (dapat diatur)
    - 📊 **ABC Analysis** — Klasifikasi item A/B/C berdasarkan omzet
    - 🧊 **Dead Stock** — Deteksi barang tidak terjual
    - 🛡️ **Audit Stok** — Validasi awal + masuk - keluar = akhir
    - 💣 **Anomaly Harga** — Deteksi harga beli janggal
    - 🛒 **Smart Reorder** — Rekomendasi order per supplier (siap WA)
    - 🧠 **AI Recommendations** — Saran actionable per item

    **v1.2.1 Fix:** Konversi satuan besar (SLOF/DUS/LSN) → PCS sekarang diterapkan dengan benar
    di seluruh pipeline FIFO (HPP per PCS, qty antrian, perbandingan modal vs harga jual).
    Error `SEVERITY` sudah diperbaiki dengan fallback kolom + pengecekan sebelum akses.
    """)

    col1, col2, col3 = st.columns(3)
    col1.markdown("""
    <div class="metric-card">
    <h3>📦 FIFO Engine</h3>
    <p>Hitung HPP akurat dengan metode FIFO berbasis antrean pembelian<br>
    <small>★ Konversi SLOF/DUS → PCS otomatis</small></p>
    </div>""", unsafe_allow_html=True)
    col2.markdown("""
    <div class="metric-warn">
    <h3>🔔 FIFO Layer Alert</h3>
    <p>Peringatan otomatis: stok lama belum habis, modal baru naik atau melewati harga jual<br>
    <small>★ Semua harga dalam satuan PCS</small></p>
    </div>""", unsafe_allow_html=True)
    col3.markdown("""
    <div class="metric-crit">
    <h3>🛡️ Audit Trail</h3>
    <p>Deteksi human error stok & anomali harga beli secara otomatis</p>
    </div>""", unsafe_allow_html=True)


if __name__ == "__main__":
    main()