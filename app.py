"""
Aplikasi Naive Bayes - Prediksi Investasi (versi Streamlit)
Jalankan lokal:  streamlit run app.py
"""
import io
import itertools

import pandas as pd
import streamlit as st

# ============================ DATA ============================
FITUR = ["Umur", "Status", "Penghasilan", "Hutang Konsumtif"]
OPSI = {
    "Umur": ["Muda", "Paruh Baya", "Tua"],
    "Status": ["Lajang", "Menikah"],
    "Penghasilan": ["Rendah", "Sedang", "Tinggi"],
    "Hutang Konsumtif": ["Ya", "Tidak"],
}
LATIH = [
    ("Muda", "Lajang", "Sedang", "Ya", "Tidak"), ("Muda", "Lajang", "Rendah", "Ya", "Tidak"),
    ("Muda", "Lajang", "Rendah", "Tidak", "Ya"), ("Muda", "Menikah", "Sedang", "Ya", "Ya"),
    ("Muda", "Menikah", "Rendah", "Ya", "Tidak"), ("Paruh Baya", "Lajang", "Tinggi", "Ya", "Tidak"),
    ("Paruh Baya", "Lajang", "Sedang", "Ya", "Tidak"), ("Paruh Baya", "Lajang", "Rendah", "Tidak", "Ya"),
    ("Paruh Baya", "Menikah", "Tinggi", "Ya", "Ya"), ("Paruh Baya", "Menikah", "Sedang", "Ya", "Ya"),
    ("Paruh Baya", "Menikah", "Sedang", "Tidak", "Ya"), ("Paruh Baya", "Menikah", "Rendah", "Ya", "Ya"),
    ("Tua", "Lajang", "Tinggi", "Ya", "Tidak"), ("Tua", "Lajang", "Tinggi", "Tidak", "Ya"),
    ("Tua", "Lajang", "Rendah", "Tidak", "Tidak"), ("Tua", "Menikah", "Tinggi", "Ya", "Ya"),
    ("Tua", "Menikah", "Sedang", "Ya", "Tidak"), ("Tua", "Menikah", "Sedang", "Tidak", "Ya"),
    ("Tua", "Menikah", "Rendah", "Ya", "Tidak"),
]
_dikenal = {r[:4] for r in LATIH}
UJI = [dict(zip(FITUR, c)) for c in itertools.product(*OPSI.values()) if c not in _dikenal]


# ======================== NAIVE BAYES =========================
def hitung(tes, laplace=False):
    n = len(LATIH)
    kelas = {}
    for k in ("Ya", "Tidak"):
        sub = [r for r in LATIH if r[4] == k]
        nk = len(sub)
        skor, rinci = nk / n, []
        for i, f in enumerate(FITUR):
            cnt = sum(1 for r in sub if r[i] == tes[f])
            p = (cnt + 1) / (nk + len(OPSI[f])) if laplace else cnt / nk
            skor *= p
            rinci.append({"Fitur": f, "Nilai": tes[f], "Hitung": f"{cnt}/{nk}", "P": p})
        kelas[k] = {"n": nk, "prior": nk / n, "rinci": rinci, "skor": skor}
    total = sum(v["skor"] for v in kelas.values())
    for v in kelas.values():
        v["post"] = v["skor"] / total if total else 0.0
    pred = max(kelas, key=lambda k: kelas[k]["skor"]) if total else "-"
    return {"kelas": kelas, "prediksi": pred, "total": total}


# ===================== IMPOR EXCEL / CSV =====================
def baca_file(nama, data):
    bio = io.BytesIO(data)
    raw = pd.read_csv(bio) if nama.lower().endswith(".csv") else pd.read_excel(bio)
    kunci = {str(c).strip().lower().replace(" ", ""): c for c in raw.columns}
    alias = {"Umur": ["umur"], "Status": ["status"], "Penghasilan": ["penghasilan"],
             "Hutang Konsumtif": ["hutangkonsumtif", "hutang"]}
    kol, hilang = {}, []
    for f, al in alias.items():
        asli = next((kunci[a] for a in al if a in kunci), None)
        if asli is None:
            hilang.append(f)
        else:
            kol[f] = raw[asli].astype(str).str.strip().str.title().tolist()
    if hilang:
        return {"error": "Kolom tidak ditemukan: " + ", ".join(hilang)}
    akt = None
    if "investasi" in kunci:
        akt = raw[kunci["investasi"]].astype(str).str.strip().str.title().tolist()
    rows, dibuang = [], 0
    for i in range(len(raw)):
        r = {f: kol[f][i] for f in FITUR}
        if all(r[f] in OPSI[f] for f in FITUR):
            if akt and akt[i] in ("Ya", "Tidak"):
                r["Aktual"] = akt[i]
            rows.append(r)
        else:
            dibuang += 1
    return {"rows": rows, "dibuang": dibuang}


def template_xlsx():
    t = pd.DataFrame(UJI)
    t["Investasi"] = "???"
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        t.to_excel(w, index=False, sheet_name="Data Uji")
    return buf.getvalue()


# ============================ UI ==============================
st.set_page_config(page_title="NB Investasi", page_icon="📈", layout="wide")

ss = st.session_state
ss.setdefault("rows", [dict(r) for r in UJI])
ss.setdefault("ver", 0)          # dipakai untuk mereset tabel editor
ss.setdefault("last_file", None)
ss.setdefault("tes", {"Umur": "Paruh Baya", "Status": "Menikah",
                      "Penghasilan": "Rendah", "Hutang Konsumtif": "Tidak"})

st.title("📈 Klasifikasi Naive Bayes")
st.caption("Prediksi keputusan **Investasi** dari Umur, Status, Penghasilan & Hutang Konsumtif.")

with st.sidebar:
    st.header("📈 NB Investasi")
    halaman = st.radio("Menu", ["🎯 Prediksi Tunggal", "🧪 Data Uji (Batch)",
                                "🧮 Langkah Perhitungan", "🗂️ Dataset Latih"])
    laplace = st.checkbox("Laplace smoothing", value=False)

c1, c2, c3 = st.columns(3)
c1.metric("Data latih", len(LATIH))
c2.metric("Data uji", len(ss.rows))
c3.metric("Mode", "Laplace ON" if laplace else "Tanpa Laplace")
st.divider()


def input_tunggal():
    st.subheader("🧪 Input Data Uji Tunggal")
    cols = st.columns(4)
    for col, f in zip(cols, FITUR):
        ss.tes[f] = col.selectbox(f, OPSI[f], index=OPSI[f].index(ss.tes[f]), key=f"tes_{f}_{halaman}")
    return dict(ss.tes)


# ---------------------- 1. PREDIKSI TUNGGAL ----------------------
if halaman == "🎯 Prediksi Tunggal":
    tes = input_tunggal()
    r = hitung(tes, laplace)
    if not r["total"]:
        st.error("Semua skor 0. Aktifkan Laplace smoothing di sidebar.")
    else:
        y, t, p = r["kelas"]["Ya"], r["kelas"]["Tidak"], r["prediksi"]
        a, b, c = st.columns(3)
        a.metric("Hasil Prediksi Investasi", f"{'💰' if p == 'Ya' else '🚫'} {p}")
        b.metric("Probabilitas Ya", f"{y['post']*100:.2f}%")
        c.metric("Probabilitas Tidak", f"{t['post']*100:.2f}%")
        st.write("**Investasi = Ya**")
        st.progress(y["post"])
        st.write("**Investasi = Tidak**")
        st.progress(t["post"])
        st.success(
            f"📌 Berdasarkan {len(LATIH)} data latih, profil **{', '.join(tes.values())}** "
            f"diprediksi **{'berinvestasi' if p == 'Ya' else 'tidak berinvestasi'}** "
            f"(Investasi = {p}) dengan keyakinan {r['kelas'][p]['post']*100:.2f}%."
        )

# ---------------------- 2. DATA UJI (BATCH) ----------------------
elif halaman == "🧪 Data Uji (Batch)":
    st.subheader("📥 Sumber data uji")
    up = st.file_uploader("Impor Excel / CSV", type=["xlsx", "xls", "csv"])
    if up is not None and ss.last_file != (up.name, up.size):
        hasil = baca_file(up.name, up.getvalue())
        if "error" in hasil:
            st.error(hasil["error"])
        else:
            ss.rows, ss.last_file = hasil["rows"], (up.name, up.size)
            ss.ver += 1
            st.success(f"✅ {len(hasil['rows'])} baris diimpor dari {up.name}."
                       + (f" {hasil['dibuang']} baris dilewati (nilai tidak dikenali)." if hasil["dibuang"] else ""))

    b1, b2 = st.columns(2)
    b1.download_button("📄 Unduh template Excel", template_xlsx(), "template_data_uji.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if b2.button("↺ Pakai 17 data soal"):
        ss.rows, ss.last_file = [dict(r) for r in UJI], None
        ss.ver += 1
        st.rerun()
    st.caption("Kolom wajib: Umur, Status, Penghasilan, Hutang Konsumtif. Kolom Investasi (isi ??? / kosong) "
               "diabaikan; bila berisi Ya/Tidak dipakai untuk menghitung akurasi.")

    st.write("**Edit / tambah / hapus baris** (klik tabel; baris baru lewat baris kosong di bawah):")
    df_in = pd.DataFrame(ss.rows)
    ada_aktual = "Aktual" in df_in.columns
    kolom = FITUR + (["Aktual"] if ada_aktual else [])
    cfg = {f: st.column_config.SelectboxColumn(f, options=OPSI[f], required=True) for f in FITUR}
    edit = st.data_editor(df_in[kolom] if len(df_in) else pd.DataFrame(columns=kolom),
                          column_config=cfg, num_rows="dynamic", use_container_width=True,
                          key=f"editor_{ss.ver}")

    valid = edit.dropna(subset=FITUR)
    valid = valid[valid.apply(lambda r: all(r[f] in OPSI[f] for f in FITUR), axis=1)]
    if valid.empty:
        st.error("Belum ada data uji.")
    else:
        out = []
        for _, row in valid.iterrows():
            r = hitung({f: row[f] for f in FITUR}, laplace)
            rec = {f: row[f] for f in FITUR}
            rec["Prediksi"] = r["prediksi"]
            rec["Keyakinan %"] = round(max(r["kelas"]["Ya"]["post"], r["kelas"]["Tidak"]["post"]) * 100, 2)
            rec["P(Ya) %"] = round(r["kelas"]["Ya"]["post"] * 100, 2)
            rec["P(Tidak) %"] = round(r["kelas"]["Tidak"]["post"] * 100, 2)
            if ada_aktual:
                akt = row.get("Aktual")
                rec["Aktual"] = akt if akt in ("Ya", "Tidak") else "-"
                rec["Cocok"] = "-" if rec["Aktual"] == "-" else ("✅" if akt == r["prediksi"] else "❌")
            out.append(rec)
        hasil_df = pd.DataFrame(out)
        hasil_df.index = range(1, len(hasil_df) + 1)

        ny = int((hasil_df["Prediksi"] == "Ya").sum())
        nt = int((hasil_df["Prediksi"] == "Tidak").sum())
        a, b, c = st.columns(3)
        a.metric("Total data uji", len(hasil_df))
        b.metric("Prediksi Ya", ny)
        c.metric("Prediksi Tidak", nt)
        st.dataframe(hasil_df, use_container_width=True)

        teks = f"📌 Dari {len(hasil_df)} data uji, **{ny}** diprediksi berinvestasi (Ya) dan **{nt}** tidak (Tidak)."
        if ada_aktual:
            cek = hasil_df[hasil_df["Cocok"] != "-"]
            if len(cek):
                ok = int((cek["Cocok"] == "✅").sum())
                teks += f" Dibandingkan kolom Investasi pada file Anda, {ok} dari {len(cek)} cocok (akurasi {ok/len(cek)*100:.1f}%)."
        st.success(teks)
        st.download_button("⬇️ Unduh hasil (CSV)", hasil_df.to_csv(index=False).encode("utf-8-sig"),
                           "hasil_prediksi.csv", "text/csv")

# ------------------ 3. LANGKAH PERHITUNGAN -----------------------
elif halaman == "🧮 Langkah Perhitungan":
    tes = input_tunggal()
    r = hitung(tes, laplace)
    K = ["Ya", "Tidak"]

    st.markdown("### 1️⃣ Probabilitas Prior")
    cols = st.columns(2)
    for col, k in zip(cols, K):
        col.metric(f"P(Investasi = {k})", f"{r['kelas'][k]['n']}/{len(LATIH)}",
                   f"= {r['kelas'][k]['prior']:.4f}", delta_color="off")

    st.markdown("### 2️⃣ Probabilitas Likelihood")
    cols = st.columns(2)
    for col, k in zip(cols, K):
        col.write(f"**Kelas Investasi = {k}**")
        d = pd.DataFrame(r["kelas"][k]["rinci"]).rename(columns={"P": f"P(x|{k})"})
        d[f"P(x|{k})"] = d[f"P(x|{k})"].map(lambda x: f"{x:.4f}")
        col.dataframe(d, hide_index=True, use_container_width=True)

    st.markdown("### 3️⃣ Skor Akhir")
    for k in K:
        v = r["kelas"][k]
        faktor = " × ".join(f"{x:.4f}" for x in [v["prior"]] + [i["P"] for i in v["rinci"]])
        st.code(f"P({k}) × ΠP(x|{k}) = {faktor} = {v['skor']:.6f}", language=None)

    st.markdown("### 4️⃣ Normalisasi & Keputusan")
    if r["total"]:
        for k in K:
            st.code(f"P({k} | X) = {r['kelas'][k]['skor']:.6f} / {r['total']:.6f} = {r['kelas'][k]['post']*100:.2f}%",
                    language=None)
        st.success(f"📌 Skor tertinggi ada pada kelas **{r['prediksi']}**, sehingga prediksi Investasi = **{r['prediksi']}**.")
    else:
        st.error("Semua skor 0. Aktifkan Laplace smoothing di sidebar.")

# ---------------------- 4. DATASET LATIH -------------------------
else:
    ya = sum(1 for r in LATIH if r[4] == "Ya")
    a, b, c = st.columns(3)
    a.metric("Total data latih", len(LATIH))
    b.metric("Investasi Ya", ya)
    c.metric("Investasi Tidak", len(LATIH) - ya)
    d = pd.DataFrame(LATIH, columns=FITUR + ["Investasi"])
    d.index = range(1, len(d) + 1)
    st.dataframe(d, use_container_width=True)

st.caption("Naive Bayes Kategorikal • Python • Data Mining")
