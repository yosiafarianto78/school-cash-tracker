"""Aplikasi Kas Sekolah (Streamlit + Supabase)."""
import io
from datetime import date, datetime

import pandas as pd
import streamlit as st
from supabase import create_client

st.set_page_config(page_title="Kas Sekolah", page_icon="📒", layout="wide")

BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
         "Agustus", "September", "Oktober", "November", "Desember"]
TAHUN = list(range(date.today().year, date.today().year - 6, -1))
HIJAU, MERAH = "#2E7D5B", "#B4473A"
ss = st.session_state

# ---------- Tampilan (font + kartu angka) ----------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
.stApp, .stApp p, .stApp label, .stApp h1, .stApp h2, .stApp h3,
.stApp input, .stApp textarea, .stApp button { font-family: 'Plus Jakarta Sans', sans-serif; }
h1 { font-size: 1.7rem !important; font-weight: 700 !important; letter-spacing: -0.01em; }
section[data-testid="stSidebar"] { border-right: 1px solid #E3E6EC; }
.kartu { background:#fff; border:1px solid #E3E6EC; border-radius:12px; padding:16px 18px; }
.kartu span { display:block; color:#5B6675; font-size:.85rem; margin-bottom:4px; }
.kartu b { font-size:1.45rem; font-weight:700; }
</style>
""", unsafe_allow_html=True)


# ---------- Fungsi bantu ----------
def rp(v):
    """Format angka ke Rupiah: 1250000 -> Rp 1.250.000"""
    return "Rp " + f"{int(v):,}".replace(",", ".")


def kartu(label, nilai, warna="#1B2B3A"):
    st.markdown(f'<div class="kartu"><span>{label}</span>'
                f'<b style="color:{warna}">{rp(nilai)}</b></div>', unsafe_allow_html=True)


def client():
    # Satu client per sesi pengguna (jangan dibagi antar pengguna)
    if "sb" not in ss:
        ss.sb = create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])
    return ss.sb


def load_master():
    if "kat" not in ss:
        sb = client()
        ss.kat = pd.DataFrame(sb.table("kategori").select("*").eq("aktif", True).order("id").execute().data)
        ss.pos = pd.DataFrame(sb.table("pos").select("*").order("id").execute().data)


COLS = ["id", "tanggal", "jumlah", "keterangan", "kategori_id",
        "kategori", "kelompok", "jenis", "pos_id", "pos", "dibatalkan", "alasan_batal"]


def load_year(y):
    """Ambil transaksi 1 tahun (disimpan di sesi supaya tidak query berulang)."""
    key = f"tx_{y}"
    if key not in ss:
        rows, start = [], 0
        while True:  # Supabase membatasi 1000 baris per request
            r = (client().table("v_transaksi").select("*")
                 .gte("tanggal", f"{y}-01-01").lte("tanggal", f"{y}-12-31")
                 .order("tanggal", desc=True).order("id", desc=True)
                 .range(start, start + 999).execute().data)
            rows += r
            if len(r) < 1000:
                break
            start += 1000
        df = pd.DataFrame(rows, columns=COLS)
        df["tanggal"] = pd.to_datetime(df["tanggal"])
        df["jumlah"] = df["jumlah"].astype(float)
        df["dibatalkan"] = df["dibatalkan"].fillna(False).astype(bool)
        ss[key] = df
    return ss[key]


def valid(df):
    """Hanya transaksi yang tidak dibatalkan (dipakai untuk semua hitungan)."""
    return df[~df["dibatalkan"]]


def invalidate():
    """Hapus cache setelah ada tambah/ubah/hapus data."""
    for k in [k for k in ss if k.startswith("tx_")]:
        del ss[k]


def ringkas(df):
    """Pemasukan, pengeluaran, saldo per pos."""
    p = valid(df).pivot_table(index="pos", columns="jenis", values="jumlah", aggfunc="sum", fill_value=0)
    for c in ("pemasukan", "pengeluaran"):
        if c not in p:
            p[c] = 0
    p["saldo"] = p["pemasukan"] - p["pengeluaran"]
    return p.reindex(ss.pos["nama"], fill_value=0)[["pemasukan", "pengeluaran", "saldo"]]


def tabel_tahunan(df, nilai):
    """Tabel bulan x pos untuk 'pemasukan', 'pengeluaran', atau 'saldo'."""
    d = valid(df)
    d = d.assign(bulan=d["tanggal"].dt.month)
    if nilai == "saldo":
        d = d.assign(v=d["jumlah"].where(d["jenis"] == "pemasukan", -d["jumlah"]))
    else:
        d = d[d["jenis"] == nilai].assign(v=lambda x: x["jumlah"])
    p = d.pivot_table(index="bulan", columns="pos", values="v", aggfunc="sum", fill_value=0)
    p = p.reindex(index=range(1, 13), columns=list(ss.pos["nama"]), fill_value=0)
    p.index = BULAN
    p["Total"] = p.sum(axis=1)
    p.loc["Total tahun"] = p.sum()
    return p


def excel(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for n in ("pemasukan", "pengeluaran", "saldo"):
            tabel_tahunan(df, n).to_excel(w, sheet_name=n.capitalize())
        (df.drop(columns=["id", "kategori_id", "pos_id", "dibatalkan"])
           .assign(tanggal=df["tanggal"].dt.date,
                   status=df["dibatalkan"].map({True: "Dibatalkan", False: "Aktif"}))
           .to_excel(w, sheet_name="Transaksi", index=False))
    return buf.getvalue()


def pilih_periode(key):
    c1, c2, _ = st.columns([1, 1, 2])
    thn = c1.selectbox("Tahun", TAHUN, key=f"thn_{key}")
    bln = c2.selectbox("Bulan", range(1, 13), index=date.today().month - 1,
                       format_func=lambda m: BULAN[m - 1], key=f"bln_{key}")
    return thn, bln


# ---------- Halaman: Login ----------
def page_login():
    _, mid, _ = st.columns([1, 1.2, 1])
    with mid:
        st.title("Kas Sekolah")
        st.caption("Masuk untuk mengelola kas SD, TK, antar-jemput, kantin, dan sumbangan.")
        with st.form("login"):
            email = st.text_input("Email")
            sandi = st.text_input("Kata sandi", type="password")
            ok = st.form_submit_button("Masuk", type="primary", width="stretch")
        if ok:
            try:
                client().auth.sign_in_with_password({"email": email, "password": sandi})
                ss.user = email
                st.rerun()
            except Exception:
                st.error("Gagal masuk. Periksa email dan kata sandi. Jika sudah benar")


# ---------- Halaman: Ringkasan ----------
def page_ringkasan():
    st.title("Ringkasan kas")
    thn, bln = pilih_periode("ring")
    df = load_year(thn)
    dm = df[df["tanggal"].dt.month == bln]
    tbl = ringkas(dm)
    tot = tbl.sum()

    c1, c2, c3 = st.columns(3)
    with c1: kartu("Pemasukan", tot["pemasukan"], HIJAU)
    with c2: kartu("Pengeluaran", tot["pengeluaran"], MERAH)
    with c3: kartu("Saldo bulan ini", tot["saldo"])
    st.write("")

    if valid(dm).empty:
        st.info(f"Belum ada transaksi di {BULAN[bln - 1]} {thn}. Mulai dari menu Catat transaksi.")
    else:
        st.bar_chart(tbl[["pemasukan", "pengeluaran"]], color=[HIJAU, MERAH], height=280)

    st.subheader("Per pos kas")
    tampil = tbl.copy()
    tampil.loc["Total"] = tot
    st.dataframe(tampil.map(rp).rename(columns=str.capitalize).rename_axis("Pos kas"), width="stretch")


# ---------- Halaman: Catat transaksi ----------
def page_catat():
    st.title("Catat transaksi")
    a, b = st.columns(2)
    nama_pos = a.selectbox("Pos kas", ss.pos["nama"])
    pos_id = int(ss.pos.loc[ss.pos["nama"] == nama_pos, "id"].iloc[0])
    jenis = b.radio("Jenis", ["pemasukan", "pengeluaran"], horizontal=True, format_func=str.capitalize)

    sub = ss.kat[(ss.kat["pos_id"] == pos_id) & (ss.kat["jenis"] == jenis)]
    if sub.empty:
        st.info(f"{nama_pos} belum punya kategori {jenis}. Pilih jenis lain atau tambahkan kategori di database.")
        return
    kel = st.selectbox("Kelompok", list(dict.fromkeys(sub["kelompok"])))
    sub = sub[sub["kelompok"] == kel]
    nama_kat = st.selectbox("Kategori", sub["nama"])
    kid = int(sub.loc[sub["nama"] == nama_kat, "id"].iloc[0])

    with st.form("form_tx", clear_on_submit=True):
        c1, c2 = st.columns(2)
        tgl = c1.date_input("Tanggal", value=date.today(), format="DD/MM/YYYY")
        jml = c2.number_input("Jumlah (Rp)", min_value=0, step=1000, value=0)
        ket = st.text_area("Keterangan (opsional)", height=80, placeholder="Contoh: SPP Oktober kelas 3A")
        ok = st.form_submit_button("Simpan transaksi", type="primary")
    if ok:
        if jml <= 0:
            st.warning("Isi jumlah lebih dari Rp 0.")
        else:
            client().table("transaksi").insert({
                "tanggal": tgl.isoformat(), "kategori_id": kid,
                "jumlah": int(jml), "keterangan": ket.strip() or None}).execute()
            invalidate()
            st.toast(f"Tersimpan: {rp(jml)} di {nama_kat}", icon="✅")


# ---------- Halaman: Riwayat (lihat, ubah, hapus) ----------
def page_riwayat():
    st.title("Riwayat transaksi")
    c1, c2, c3 = st.columns(3)
    thn = c1.selectbox("Tahun", TAHUN, key="thn_riw")
    bln = c2.selectbox("Bulan", range(0, 13), format_func=lambda m: "Semua bulan" if m == 0 else BULAN[m - 1])
    pos = c3.selectbox("Pos kas", ["Semua pos"] + list(ss.pos["nama"]))
    q = st.text_input("Cari", placeholder="Cari kategori atau keterangan")
    sembunyi = st.checkbox("Sembunyikan transaksi yang dibatalkan")

    d = load_year(thn)
    if bln: d = d[d["tanggal"].dt.month == bln]
    if pos != "Semua pos": d = d[d["pos"] == pos]
    if q:
        d = d[d["keterangan"].fillna("").str.contains(q, case=False, regex=False)
              | d["kategori"].str.contains(q, case=False, regex=False)]
    if sembunyi:
        d = valid(d)
    d = d.reset_index(drop=True)

    if d.empty:
        st.info("Tidak ada transaksi untuk filter ini.")
        return
    view = pd.DataFrame({
        "Tanggal": d["tanggal"].dt.strftime("%d/%m/%Y"), "Pos": d["pos"],
        "Jenis": d["jenis"].str.capitalize(), "Kelompok": d["kelompok"],
        "Kategori": d["kategori"], "Jumlah": d["jumlah"].map(rp),
        "Keterangan": d["keterangan"].fillna(""),
        "Status": d["dibatalkan"].map({True: "Dibatalkan", False: "Aktif"}),
        "Alasan batal": d["alasan_batal"].fillna("")})

    def gaya(row):  # baris dibatalkan: abu-abu dan dicoret
        if row["Status"] == "Dibatalkan":
            return ["color:#9AA3AF;text-decoration:line-through;background-color:#F1F2F4"] * len(row)
        return [""] * len(row)

    st.caption("Klik satu baris untuk mengubah atau membatalkannya.")
    ev = st.dataframe(view.style.apply(gaya, axis=1), hide_index=True, width="stretch",
                      on_select="rerun", selection_mode="single-row")
    if not ev.selection.rows:
        return

    r = d.iloc[ev.selection.rows[0]]
    tabel = client().table("transaksi")
    with st.container(border=True):
        if r["dibatalkan"]:
            st.subheader("Transaksi dibatalkan")
            st.caption(f"{r['pos']}, {r['kategori']}, {rp(r['jumlah'])}")
            st.info(f"Tidak masuk hitungan. Alasan: {r['alasan_batal'] or '-'}")
            if st.button("Pulihkan transaksi", type="primary"):
                tabel.update({"dibatalkan": False, "alasan_batal": None,
                              "dibatalkan_pada": None}).eq("id", int(r["id"])).execute()
                invalidate(); st.toast("Transaksi dipulihkan", icon="↩️"); st.rerun()
            return

        st.subheader("Ubah transaksi")
        st.caption(f"{r['pos']}, {r['kategori']}")
        with st.form(f"edit_{r['id']}"):
            e1, e2 = st.columns(2)
            tgl = e1.date_input("Tanggal", value=r["tanggal"].date(), format="DD/MM/YYYY")
            jml = e2.number_input("Jumlah (Rp)", min_value=1, step=1000, value=int(r["jumlah"]))
            ket = st.text_area("Keterangan", value=r["keterangan"] if isinstance(r["keterangan"], str) else "")
            alasan = st.text_input("Alasan pembatalan (wajib jika membatalkan)",
                                   placeholder="Contoh: salah input jumlah")
            s1, s2 = st.columns(2)
            simpan = s1.form_submit_button("Simpan perubahan", type="primary")
            batal = s2.form_submit_button("Batalkan transaksi")
        if simpan:
            tabel.update({"tanggal": tgl.isoformat(), "jumlah": int(jml),
                          "keterangan": ket.strip() or None}).eq("id", int(r["id"])).execute()
            invalidate(); st.toast("Perubahan tersimpan", icon="✅"); st.rerun()
        if batal:
            if alasan.strip():
                tabel.update({"dibatalkan": True, "alasan_batal": alasan.strip(),
                              "dibatalkan_pada": datetime.now().isoformat()}).eq("id", int(r["id"])).execute()
                invalidate(); st.toast("Transaksi dibatalkan", icon="🚫"); st.rerun()
            else:
                st.warning("Isi alasan pembatalan dulu.")


# ---------- Halaman: Rekap ----------
def page_rekap():
    st.title("Rekap kas")
    tab_bln, tab_thn = st.tabs(["Bulanan", "Tahunan"])

    with tab_bln:
        thn, bln = pilih_periode("rek")
        dm = load_year(thn)
        dm = valid(dm[dm["tanggal"].dt.month == bln])
        for _, p in ss.pos.iterrows():
            dp = dm[dm["pos_id"] == p["id"]]
            masuk = dp.loc[dp["jenis"] == "pemasukan", "jumlah"].sum()
            keluar = dp.loc[dp["jenis"] == "pengeluaran", "jumlah"].sum()
            with st.expander(p["nama"], expanded=not dp.empty):
                if dp.empty:
                    st.caption("Belum ada transaksi di pos ini.")
                    continue
                k1, k2, k3 = st.columns(3)
                with k1: kartu("Pemasukan", masuk, HIJAU)
                with k2: kartu("Pengeluaran", keluar, MERAH)
                with k3: kartu("Saldo", masuk - keluar)
                g = dp.groupby(["jenis", "kelompok", "kategori"], as_index=False)["jumlah"].sum()
                g["jumlah"] = g["jumlah"].map(rp)
                g.columns = ["Jenis", "Kelompok", "Kategori", "Jumlah"]
                g["Jenis"] = g["Jenis"].str.capitalize()
                st.dataframe(g, hide_index=True, width="stretch")

    with tab_thn:
        thn = st.selectbox("Tahun", TAHUN, key="thn_rek_thn")
        df = load_year(thn)
        nilai = st.radio("Tampilkan", ["pemasukan", "pengeluaran", "saldo"],
                         horizontal=True, format_func=str.capitalize)
        st.dataframe(tabel_tahunan(df, nilai).map(rp), width="stretch", height=520)
        st.download_button("Unduh Excel", excel(df), file_name=f"rekap_kas_{thn}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------- Menu utama ----------
if "user" not in ss:
    page_login()
    st.stop()

load_master()
HALAMAN = {"Ringkasan": page_ringkasan, "Catat transaksi": page_catat,
           "Riwayat": page_riwayat, "Rekap": page_rekap}
with st.sidebar:
    st.markdown("### Kas Sekolah")
    menu = st.radio("Menu", list(HALAMAN), label_visibility="collapsed")
    st.divider()
    st.caption(ss.user)
    if st.button("Keluar"):
        client().auth.sign_out()
        ss.clear()
        st.rerun()
HALAMAN[menu]()