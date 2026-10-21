"""
StatDesk - Aplikasi Interaktif Statistika Deskriptif
UTS Statistika - S-1 Matematika (Python + Streamlit)

Jalankan:  streamlit run app.py
"""
import io
import math

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------------
# Konfigurasi halaman
# ----------------------------------------------------------------------------
st.set_page_config(page_title="StatDesk - Statistika Deskriptif", page_icon="📊", layout="wide")

SEMUA_UKURAN = [
    "N", "Minimum", "Maksimum", "Mean", "Median", "Modus",
    "Varians", "Standar Deviasi", "Q1", "Q2", "Q3",
]

# ----------------------------------------------------------------------------
# Fungsi perhitungan
# ----------------------------------------------------------------------------
def hitung_modus(x: pd.Series) -> str:
    """Modus (bisa lebih dari satu). Jika semua nilai unik -> 'Tidak ada'."""
    if x.empty:
        return "-"
    hitung = x.value_counts()
    maks = hitung.max()
    if maks == 1:
        return "Tidak ada"
    modus = sorted(hitung[hitung == maks].index.tolist())
    return ", ".join(f"{m:g}" for m in modus)


def hitung_statistik(x: pd.Series, ukuran: list, jenis_varians: str) -> dict:
    """Hitung ukuran terpilih untuk satu variabel (x sudah bebas NaN)."""
    ddof = 1 if jenis_varians == "Sampel (n-1)" else 0
    n = len(x)
    hasil = {}
    for u in ukuran:
        if n == 0:
            hasil[u] = np.nan
            continue
        if u == "N":
            hasil[u] = n
        elif u == "Minimum":
            hasil[u] = x.min()
        elif u == "Maksimum":
            hasil[u] = x.max()
        elif u == "Mean":
            hasil[u] = x.mean()
        elif u == "Median":
            hasil[u] = x.median()
        elif u == "Modus":
            hasil[u] = hitung_modus(x)
        elif u == "Varians":
            hasil[u] = x.var(ddof=ddof) if n > ddof else np.nan
        elif u == "Standar Deviasi":
            hasil[u] = x.std(ddof=ddof) if n > ddof else np.nan
        elif u == "Q1":
            hasil[u] = x.quantile(0.25)   # metode linear = QUARTILE.INC Excel
        elif u == "Q2":
            hasil[u] = x.quantile(0.50)
        elif u == "Q3":
            hasil[u] = x.quantile(0.75)
    return hasil


def tabel_frekuensi(x: pd.Series) -> pd.DataFrame:
    """Tabel distribusi frekuensi berkelompok (aturan Sturges)."""
    n = len(x)
    if n < 2 or x.max() == x.min():
        return pd.DataFrame()
    k = int(round(1 + 3.322 * math.log10(n)))
    rentang = x.max() - x.min()
    panjang = math.ceil(rentang / k * 100) / 100  # bulatkan ke atas 2 desimal
    batas = [x.min() + i * panjang for i in range(k + 1)]
    kelas = pd.cut(x, bins=batas, include_lowest=True, right=True)
    frek = kelas.value_counts(sort=False)
    df = pd.DataFrame({
        "Interval Kelas": [f"{iv.left:.2f} - {iv.right:.2f}" for iv in frek.index],
        "Titik Tengah": [round((iv.left + iv.right) / 2, 2) for iv in frek.index],
        "Frekuensi": frek.values,
    })
    df["Frek. Relatif (%)"] = (df["Frekuensi"] / n * 100).round(2)
    df["Frek. Kumulatif"] = df["Frekuensi"].cumsum()
    return df


@st.cache_data(show_spinner=False)
def baca_file(isi: bytes, nama: str, sheet=None):
    """Baca .xlsx/.xls/.csv. Mengembalikan (DataFrame, daftar_sheet)."""
    nama = nama.lower()
    if nama.endswith(".csv"):
        try:
            return pd.read_csv(io.BytesIO(isi)), []
        except Exception:
            return pd.read_csv(io.BytesIO(isi), sep=";"), []
    engine = "xlrd" if nama.endswith(".xls") else "openpyxl"
    xl = pd.ExcelFile(io.BytesIO(isi), engine=engine)
    daftar = xl.sheet_names
    return xl.parse(sheet or daftar[0]), daftar


def ke_excel(sheets: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nama, df in sheets.items():
            df.to_excel(w, sheet_name=nama[:31], index=True if nama == "Ringkasan" else False)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# Sidebar: pengaturan
# ----------------------------------------------------------------------------
st.sidebar.title("📊 StatDesk")
st.sidebar.caption("Statistika Deskriptif Otomatis")

halaman = st.sidebar.radio(
    "Navigasi",
    ["1. Data", "2. Statistik Deskriptif", "3. Visualisasi", "4. Unduh Hasil", "ℹ️ Petunjuk"],
)

st.sidebar.divider()
st.sidebar.subheader("Pengaturan")
jenis_varians = st.sidebar.radio("Varians & Standar Deviasi", ["Sampel (n-1)", "Populasi (n)"])
metode_missing = st.sidebar.selectbox(
    "Penanganan data kosong (missing)",
    ["Abaikan (hapus nilai kosong)", "Isi dengan Mean", "Isi dengan Median", "Isi dengan 0"],
)
warna = st.sidebar.color_picker("Warna grafik", "#2E75B6")

# ----------------------------------------------------------------------------
# Unggah data (disimpan di session_state agar tetap ada di semua halaman)
# ----------------------------------------------------------------------------
st.title("Aplikasi Statistika Deskriptif")

with st.expander("📂 Unggah dataset", expanded=("df" not in st.session_state)):
    file = st.file_uploader("Pilih file Excel (.xlsx / .xls) atau CSV", type=["xlsx", "xls", "csv"])
    if file is not None:
        try:
            isi = file.getvalue()
            _, daftar_sheet = baca_file(isi, file.name)
            sheet = None
            if len(daftar_sheet) > 1:
                sheet = st.selectbox("Pilih sheet", daftar_sheet)
            df, _ = baca_file(isi, file.name, sheet)
            st.session_state["df"] = df
            st.session_state["nama_file"] = file.name
            st.success(f"Berhasil membaca **{file.name}** ({df.shape[0]} baris, {df.shape[1]} kolom).")
        except Exception as e:
            st.error(f"Gagal membaca file: {e}")

if "df" not in st.session_state:
    st.info("Silakan unggah dataset terlebih dahulu. Contoh: `dataset_contoh.xlsx`.")
    st.stop()

df: pd.DataFrame = st.session_state["df"]
kolom_numerik = df.select_dtypes(include="number").columns.tolist()


def siapkan(seri: pd.Series) -> pd.Series:
    """Terapkan penanganan missing value sesuai pilihan pengguna."""
    if metode_missing.startswith("Abaikan"):
        return seri.dropna()
    if metode_missing.endswith("Mean"):
        return seri.fillna(seri.mean())
    if metode_missing.endswith("Median"):
        return seri.fillna(seri.median())
    return seri.fillna(0)


# ----------------------------------------------------------------------------
# Halaman 1: Data
# ----------------------------------------------------------------------------
if halaman == "1. Data":
    st.header("Tampilan Dataset")
    c1, c2, c3 = st.columns(3)
    c1.metric("Jumlah baris", df.shape[0])
    c2.metric("Jumlah kolom", df.shape[1])
    c3.metric("Kolom numerik", len(kolom_numerik))
    st.dataframe(df, width='stretch')

    st.subheader("Data kosong (missing values) per kolom")
    mv = df.isna().sum().rename("Jumlah kosong").to_frame()
    mv["Persen (%)"] = (mv["Jumlah kosong"] / len(df) * 100).round(2)
    st.dataframe(mv, width='stretch')
    if mv["Jumlah kosong"].sum() > 0:
        st.warning("Terdapat data kosong. Atur penanganannya di sidebar (bagian Pengaturan).")

# ----------------------------------------------------------------------------
# Halaman 2: Statistik Deskriptif
# ----------------------------------------------------------------------------
elif halaman == "2. Statistik Deskriptif":
    st.header("Statistik Deskriptif")
    if not kolom_numerik:
        st.error("Dataset tidak memiliki kolom numerik.")
        st.stop()

    var_pilih = st.multiselect("Pilih variabel numerik", kolom_numerik, default=kolom_numerik[:1])
    ukuran_pilih = st.multiselect("Pilih ukuran statistik", SEMUA_UKURAN, default=SEMUA_UKURAN)
    # urutan ukuran tetap mengikuti urutan baku
    ukuran_pilih = [u for u in SEMUA_UKURAN if u in ukuran_pilih]

    if not var_pilih or not ukuran_pilih:
        st.info("Pilih minimal satu variabel dan satu ukuran statistik.")
        st.stop()

    baris = {v: hitung_statistik(siapkan(df[v]), ukuran_pilih, jenis_varians) for v in var_pilih}
    ringkasan = pd.DataFrame(baris).T[ukuran_pilih]
    ringkasan.index.name = "Variabel"
    st.session_state["ringkasan"] = ringkasan

    tampil = ringkasan.copy()
    for k in tampil.columns:
        if k != "Modus":
            tampil[k] = pd.to_numeric(tampil[k], errors="coerce")
    fmt = {k: ("{:.0f}" if k == "N" else "{:.4f}") for k in tampil.columns if k != "Modus"}
    st.dataframe(tampil.style.format(fmt, na_rep="-"), width='stretch')

    kosong = {v: int(df[v].isna().sum()) for v in var_pilih if df[v].isna().sum() > 0}
    if kosong:
        st.caption(f"Data kosong ditangani dengan: **{metode_missing}**. Jumlah kosong: {kosong}")
    st.caption(
        f"Varians & standar deviasi: **{jenis_varians}**. "
        "Kuartil memakai metode interpolasi linear (setara QUARTILE.INC di Excel)."
    )

    st.subheader("Tabel Distribusi Frekuensi")
    v_frek = st.selectbox("Variabel untuk tabel frekuensi", var_pilih)
    frek = tabel_frekuensi(siapkan(df[v_frek]))
    st.session_state["frekuensi"] = frek
    if frek.empty:
        st.info("Data tidak cukup untuk membuat tabel frekuensi.")
    else:
        st.dataframe(frek, width='stretch', hide_index=True)
        st.caption("Jumlah kelas memakai aturan Sturges: k = 1 + 3,322 log n.")

# ----------------------------------------------------------------------------
# Halaman 3: Visualisasi
# ----------------------------------------------------------------------------
elif halaman == "3. Visualisasi":
    st.header("Visualisasi Data")
    if not kolom_numerik:
        st.error("Dataset tidak memiliki kolom numerik.")
        st.stop()

    jenis = st.selectbox("Jenis grafik", ["Histogram", "Boxplot", "Diagram Batang (kategori)", "Diagram Lingkaran (kategori)"])

    if jenis in ("Histogram", "Boxplot"):
        vs = st.multiselect("Variabel numerik", kolom_numerik, default=kolom_numerik[:1])
        if not vs:
            st.info("Pilih minimal satu variabel.")
            st.stop()
        fig, ax = plt.subplots(figsize=(8, 4.5))
        if jenis == "Histogram":
            bins = st.slider("Jumlah bin", 3, 30, 8)
            for v in vs:
                ax.hist(siapkan(df[v]), bins=bins, alpha=0.7 if len(vs) > 1 else 1,
                        color=warna if len(vs) == 1 else None, edgecolor="white", label=v)
            ax.set_xlabel("Nilai"); ax.set_ylabel("Frekuensi"); ax.set_title("Histogram")
            if len(vs) > 1:
                ax.legend()
        else:
            ax.boxplot([siapkan(df[v]) for v in vs], tick_labels=vs, patch_artist=True,
                       boxprops=dict(facecolor=warna, alpha=0.7),
                       medianprops=dict(color="black"))
            ax.set_title("Boxplot"); ax.set_ylabel("Nilai")
        ax.grid(alpha=0.3)
        st.pyplot(fig)
    else:
        kolom_kat = [c for c in df.columns if df[c].nunique(dropna=True) <= 20]
        if not kolom_kat:
            st.info("Tidak ada kolom kategorik (maks. 20 kategori unik).")
            st.stop()
        v = st.selectbox("Variabel kategori", kolom_kat)
        hitung = df[v].value_counts(dropna=True)
        fig, ax = plt.subplots(figsize=(7, 4.5))
        if jenis.startswith("Diagram Batang"):
            ax.bar(hitung.index.astype(str), hitung.values, color=warna, edgecolor="white")
            ax.set_ylabel("Frekuensi"); ax.set_title(f"Diagram Batang: {v}")
            ax.grid(axis="y", alpha=0.3)
            plt.xticks(rotation=30, ha="right")
        else:
            ax.pie(hitung.values, labels=hitung.index.astype(str), autopct="%1.1f%%", startangle=90)
            ax.set_title(f"Diagram Lingkaran: {v}")
        st.pyplot(fig)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
    st.download_button("⬇️ Unduh grafik (PNG)", buf.getvalue(), "grafik.png", "image/png")

# ----------------------------------------------------------------------------
# Halaman 4: Unduh hasil
# ----------------------------------------------------------------------------
elif halaman == "4. Unduh Hasil":
    st.header("Unduh Ringkasan Hasil")
    if "ringkasan" not in st.session_state:
        st.info("Hitung statistik terlebih dahulu di halaman **2. Statistik Deskriptif**.")
        st.stop()
    ringkasan = st.session_state["ringkasan"]
    st.dataframe(ringkasan, width='stretch')

    sheets = {"Ringkasan": ringkasan}
    if isinstance(st.session_state.get("frekuensi"), pd.DataFrame) and not st.session_state["frekuensi"].empty:
        sheets["Distribusi Frekuensi"] = st.session_state["frekuensi"]

    c1, c2 = st.columns(2)
    c1.download_button("⬇️ Unduh Excel (.xlsx)", ke_excel(sheets), "hasil_statistik.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    c2.download_button("⬇️ Unduh CSV", ringkasan.to_csv().encode("utf-8-sig"), "hasil_statistik.csv", "text/csv")

# ----------------------------------------------------------------------------
# Petunjuk
# ----------------------------------------------------------------------------
else:
    st.header("Petunjuk Penggunaan")
    st.markdown(
        """
1. **Unggah data** pada panel *Unggah dataset* (Excel `.xlsx`/`.xls`, CSV juga boleh).
2. Buka **1. Data** untuk melihat isi tabel dan jumlah data kosong.
3. Buka **2. Statistik Deskriptif**, pilih satu/beberapa variabel numerik dan ukuran statistik
   (N, min, maks, mean, median, modus, varians, standar deviasi, Q1, Q2, Q3).
4. Buka **3. Visualisasi** untuk histogram, boxplot, diagram batang, atau diagram lingkaran.
5. Buka **4. Unduh Hasil** untuk menyimpan ringkasan ke Excel/CSV.

Pengaturan di sidebar: jenis varians (sampel/populasi), penanganan data kosong, dan warna grafik.
        """
    )
