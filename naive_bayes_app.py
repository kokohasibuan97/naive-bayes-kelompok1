"""
Aplikasi Naive Bayes - Prediksi Investasi (TANPA Streamlit)
Jalankan:   python naive_bayes_app.py
Browser terbuka otomatis. Untuk impor Excel butuh: pip install pandas openpyxl
"""
import io
import itertools
import json
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

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
UJI = [dict(zip(FITUR, c)) for c in itertools.product(*OPSI.values()) if c not in _dikenal]  # 17 baris "???"


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
            rinci.append({"f": f, "v": tes[f], "cnt": cnt, "nk": nk, "p": p})
        kelas[k] = {"n": nk, "prior": nk / n, "rinci": rinci, "skor": skor}
    total = sum(v["skor"] for v in kelas.values())
    for v in kelas.values():
        v["post"] = v["skor"] / total if total else 0.0
    pred = max(kelas, key=lambda k: kelas[k]["skor"]) if total else "-"
    return {"kelas": kelas, "prediksi": pred, "total": total}


# ===================== IMPOR EXCEL / CSV =====================
def baca_file(nama, data):
    import pandas as pd
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
    import pandas as pd
    t = pd.DataFrame(UJI)
    t["Investasi"] = "???"
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        t.to_excel(w, index=False, sheet_name="Data Uji")
    return buf.getvalue()


# ============================ SERVER ==========================
class H(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def kirim(self, body, ctype="application/json", extra=None):
        if not isinstance(body, bytes):
            body = (body if isinstance(body, str) else json.dumps(body)).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", ctype + ("; charset=utf-8" if ctype.startswith("text") or "json" in ctype else ""))
        self.send_header("Content-Length", str(len(body)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = urlparse(self.path).path
        try:
            if path == "/":
                self.kirim(PAGE, "text/html")
            elif path == "/api/meta":
                self.kirim({"fitur": FITUR, "opsi": OPSI, "latih": LATIH, "uji": UJI})
            elif path == "/api/template":
                self.kirim(template_xlsx(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           {"Content-Disposition": 'attachment; filename="template_data_uji.xlsx"'})
            else:
                self.send_error(404)
        except Exception as e:  # noqa
            self.kirim({"error": f"{e} (pasang: pip install pandas openpyxl)"})

    def do_POST(self):
        u = urlparse(self.path)
        raw = self.rfile.read(int(self.headers.get("Content-Length", 0)))
        try:
            if u.path == "/api/predict":
                d = json.loads(raw)
                self.kirim(hitung(d["tes"], d["laplace"]))
            elif u.path == "/api/batch":
                d = json.loads(raw)
                out = []
                for r in d["rows"]:
                    x = hitung(r, d["laplace"])
                    out.append({"pred": x["prediksi"], "ya": x["kelas"]["Ya"]["post"],
                                "tidak": x["kelas"]["Tidak"]["post"]})
                self.kirim({"hasil": out})
            elif u.path == "/api/upload":
                nama = parse_qs(u.query).get("name", ["data.xlsx"])[0]
                self.kirim(baca_file(nama, raw))
            else:
                self.send_error(404)
        except Exception as e:  # noqa
            self.kirim({"error": f"Gagal memproses file: {e}. Pastikan sudah: pip install pandas openpyxl"})


# ============================ TAMPILAN ========================
PAGE = r"""<!DOCTYPE html><html lang="id"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>Naive Bayes Investasi</title>
<style>
:root{--bg:#f3f5fb;--card:#fff;--tx:#0f172a;--mut:#64748b;--bd:#e2e8f0;--soft:#eef0ff;--r:20px}
@media(prefers-color-scheme:dark){:root{--bg:#0b1020;--card:#141a2e;--tx:#e5e7eb;--mut:#94a3b8;--bd:#263049;--soft:#1c2442}}
*{box-sizing:border-box}body{margin:0;font-family:Inter,Segoe UI,system-ui,sans-serif;background:var(--bg);color:var(--tx)}
.app{display:grid;grid-template-columns:270px 1fr;min-height:100vh}
aside{background:linear-gradient(180deg,#111936,#0b1020);color:#e5e7eb;padding:26px 18px;position:sticky;top:0;height:100vh;overflow:auto}
.brand{display:flex;gap:12px;align-items:center;margin-bottom:26px}
.logo{width:44px;height:44px;border-radius:14px;display:grid;place-items:center;font-size:1.4rem;background:linear-gradient(135deg,#6366f1,#06b6d4)}
.brand b{display:block;font-size:1.05rem}.brand small{opacity:.6}
.sec{font-size:.7rem;letter-spacing:.12em;text-transform:uppercase;opacity:.5;font-weight:700;margin:22px 6px 10px}
.mi{display:block;width:100%;text-align:left;padding:12px 14px;margin-bottom:6px;border:0;border-radius:14px;background:rgba(255,255,255,.05);
 color:#e5e7eb;font-weight:600;font-size:.94rem;font-family:inherit;cursor:pointer;transition:.15s}
.mi:hover{background:rgba(255,255,255,.12)}.mi.on{background:linear-gradient(135deg,#6366f1,#06b6d4);box-shadow:0 10px 22px -12px #6366f1}
.sw{display:flex;align-items:center;justify-content:space-between;padding:12px 14px;border-radius:14px;background:rgba(255,255,255,.05);font-weight:600;font-size:.92rem;color:#e5e7eb}
.sw input{appearance:none;width:44px;height:24px;border-radius:99px;background:#475569;position:relative;cursor:pointer;transition:.2s}
.sw input:after{content:"";position:absolute;top:3px;left:3px;width:18px;height:18px;border-radius:50%;background:#fff;transition:.2s}
.sw input:checked{background:#10b981}.sw input:checked:after{left:23px}
main{padding:28px 34px 20px;max-width:1150px;width:100%}
.hero{padding:32px 36px;border-radius:24px;color:#fff;margin-bottom:22px;overflow:hidden;isolation:isolate;
 background:radial-gradient(circle at 85% 10%,rgba(236,72,153,.55),transparent 45%),linear-gradient(135deg,#4338ca,#6d28d9 50%,#0891b2);
 box-shadow:0 18px 40px -18px rgba(79,70,229,.7)}
.hero h1{margin:0;font-size:2rem;letter-spacing:-.5px}.hero p{margin:8px 0 0;opacity:.92}
.chip{display:inline-block;padding:5px 13px;margin:14px 8px 0 0;border-radius:99px;background:rgba(255,255,255,.2);font-size:.8rem;font-weight:600}
.card{background:var(--card);border:1px solid var(--bd);border-radius:var(--r);padding:20px 24px;margin-bottom:18px}
.card h3{margin:0 0 14px;font-size:1.05rem}
.g3{display:grid;grid-template-columns:1.2fr 1fr 1fr;gap:18px}.g3 .card{margin:0}.g4{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}
.l{font-size:.74rem;text-transform:uppercase;letter-spacing:.08em;color:var(--mut);font-weight:700}.v{font-size:2.2rem;font-weight:800;margin-top:4px}
.big.ya{background:linear-gradient(135deg,#059669,#10b981);color:#fff;border:0;box-shadow:0 14px 30px -14px #10b981}
.big.tidak{background:linear-gradient(135deg,#e11d48,#f43f5e);color:#fff;border:0;box-shadow:0 14px 30px -14px #f43f5e}.big .l{color:#ffffffcc}
label{font-size:.82rem;font-weight:600;color:var(--mut)}
select,button.b{width:100%;margin-top:6px;padding:11px 12px;border-radius:14px;border:1px solid var(--bd);background:var(--soft);color:var(--tx);font:inherit}
button.b{width:auto;margin:0;font-weight:700;cursor:pointer;background:linear-gradient(135deg,#6366f1,#06b6d4);color:#fff;border:0;padding:11px 18px;text-decoration:none;display:inline-block}
button.b.alt{background:var(--soft);color:var(--tx);border:1px solid var(--bd)}
.row{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-bottom:14px}.row a{text-decoration:none}
.bw{margin:10px 0}.bh{display:flex;justify-content:space-between;font-weight:600;font-size:.92rem;margin-bottom:6px}
.bar{height:16px;border-radius:99px;background:var(--bd);overflow:hidden}.bar>div{height:100%;border-radius:99px}
.g{background:linear-gradient(90deg,#10b981,#34d399)}.r{background:linear-gradient(90deg,#f43f5e,#fb7185)}
.badge{padding:3px 12px;border-radius:99px;font-weight:700;font-size:.8rem}.badge.ya{background:#10b98126;color:#10b981}.badge.tidak{background:#f43f5e26;color:#f43f5e}
.tw{overflow-x:auto;border:1px solid var(--bd);border-radius:var(--r);background:var(--card)}
table{width:100%;border-collapse:collapse;font-size:.9rem}th{text-align:left;padding:12px 14px;background:var(--soft);font-size:.74rem;text-transform:uppercase;letter-spacing:.06em}
td{padding:10px 14px;border-top:1px solid var(--bd)}tr:hover td{background:var(--soft)}
.mini{display:inline-block;width:80px;height:8px;border-radius:99px;background:var(--bd);margin-right:8px;overflow:hidden;vertical-align:middle}.mini>i{display:block;height:100%}
.x{border:0;background:none;cursor:pointer;color:#f43f5e;font-size:1rem}
.concl{margin-top:20px;padding:22px 26px;border-radius:var(--r);color:#fff;background:linear-gradient(135deg,#4338ca,#0891b2)}
.concl h4{margin:0 0 6px}.concl p{margin:0;line-height:1.55}
.note{padding:13px 16px;border-radius:14px;margin-bottom:14px;font-size:.92rem}.ok{background:#10b98120;color:#059669}.err{background:#f43f5e20;color:#e11d48}
.step{display:flex;gap:12px;align-items:center;margin:24px 0 10px;font-weight:700;font-size:1.1rem}
.step i{font-style:normal;width:34px;height:34px;border-radius:12px;display:grid;place-items:center;color:#fff;background:linear-gradient(135deg,#6366f1,#06b6d4)}
.fm{font-family:Consolas,monospace;padding:14px 16px;border-radius:14px;background:var(--soft);border-left:4px solid #6366f1;margin:8px 0;font-size:.88rem}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:18px}.foot{text-align:center;color:var(--mut);font-size:.8rem;margin:36px 0 8px;padding-top:16px;border-top:1px solid var(--bd)}
@media(max-width:850px){.app{grid-template-columns:1fr}aside{position:static;height:auto}.g3,.g4,.g2{grid-template-columns:1fr}}
</style></head><body><div class="app">
<aside><div class="brand"><div class="logo">📈</div><div><b>NB Investasi</b><small>Klasifikasi Naive Bayes</small></div></div>
<div class="sec">Menu</div><div id="menu"></div>
<div class="sec">Pengaturan</div>
<label class="sw">Laplace smoothing<input type="checkbox" onchange="S.laplace=this.checked;show()"></label></aside>
<main><div class="hero"><h1>📈 Klasifikasi Naive Bayes</h1>
<p>Prediksi keputusan <b>Investasi</b> dari Umur, Status, Penghasilan &amp; Hutang Konsumtif.</p><div id="chips"></div></div>
<div id="page"></div>
<div class="foot">Naive Bayes Kategorikal • Python • Data Mining</div></main></div>
<script>
const $=s=>document.querySelector(s);
const PAGES=['🎯 Prediksi Tunggal','🧪 Data Uji (Batch)','🧮 Langkah Perhitungan','🗂️ Dataset Latih'];
let M,S={page:0,laplace:false,tes:{Umur:'Paruh Baya',Status:'Menikah',Penghasilan:'Rendah','Hutang Konsumtif':'Tidak'},rows:[],nr:{},msg:null,out:[]};
const pct=x=>(x*100).toFixed(2)+'%';
const badge=k=>k&&k!='-'?`<span class="badge ${k.toLowerCase()}">${k}</span>`:'-';
const bar=(l,p,c)=>`<div class="bw"><div class="bh"><span>${l}</span><span>${pct(p)}</span></div><div class="bar"><div class="${c}" style="width:${p*100}%"></div></div></div>`;
const concl=t=>`<div class="concl"><h4>📌 Kesimpulan</h4><p>${t}</p></div>`;
const api=async(p,b)=>(await fetch(p,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(b)})).json();
const opts=(f,v)=>M.opsi[f].map(o=>`<option ${o==v?'selected':''}>${o}</option>`).join('');
async function init(){M=await(await fetch('/api/meta')).json();S.rows=M.uji.map(r=>({...r}));
 M.fitur.forEach(f=>S.nr[f]=M.opsi[f][0]);menu();show()}
function menu(){$('#menu').innerHTML=PAGES.map((p,i)=>`<button class="mi ${i==S.page?'on':''}" onclick="S.page=${i};S.msg=null;menu();show()">${p}</button>`).join('')}
async function show(){
 $('#chips').innerHTML=`<span class="chip">🗂️ ${M.latih.length} data latih</span><span class="chip">🧪 ${S.rows.length} data uji</span><span class="chip">${S.laplace?'✅ Laplace ON':'🔢 Tanpa Laplace'}</span>`;
 $('#page').innerHTML=await[pPred,pBatch,pSteps,pData][S.page]();}
function inputCard(){return `<div class="card"><h3>🧪 Input Data Uji Tunggal</h3><div class="g4">${M.fitur.map(f=>
 `<div><label>${f}</label><select onchange="S.tes['${f}']=this.value;show()">${opts(f,S.tes[f])}</select></div>`).join('')}</div></div>`}
async function pPred(){
 const r=await api('/api/predict',{tes:S.tes,laplace:S.laplace});let h=inputCard();
 if(!r.total)return h+'<div class="note err">Semua skor 0. Aktifkan Laplace smoothing di sidebar.</div>';
 const y=r.kelas.Ya,t=r.kelas.Tidak,p=r.prediksi;
 h+=`<div class="g3" style="margin-bottom:18px"><div class="card big ${p.toLowerCase()}"><div class="l">Hasil Prediksi Investasi</div><div class="v">${p=='Ya'?'💰':'🚫'} ${p}</div></div>
 <div class="card"><div class="l">Probabilitas Ya</div><div class="v">${pct(y.post)}</div></div><div class="card"><div class="l">Probabilitas Tidak</div><div class="v">${pct(t.post)}</div></div></div>
 <div class="card">${bar('Investasi = Ya',y.post,'g')}${bar('Investasi = Tidak',t.post,'r')}</div>`;
 return h+concl(`Berdasarkan ${M.latih.length} data latih, profil <b>${Object.values(S.tes).join(', ')}</b> diprediksi <b>${p=='Ya'?'berinvestasi':'tidak berinvestasi'}</b> (Investasi = ${p}) dengan keyakinan ${pct(r.kelas[p].post)}.`)}
async function upload(inp){const f=inp.files[0];if(!f)return;
 const j=await(await fetch('/api/upload?name='+encodeURIComponent(f.name),{method:'POST',body:await f.arrayBuffer()})).json();
 if(j.error)S.msg={c:'err',m:j.error};else{S.rows=j.rows;S.msg={c:'ok',m:`✅ ${j.rows.length} baris data uji berhasil diimpor dari <b>${f.name}</b>.`+(j.dibuang?` ${j.dibuang} baris dilewati (nilai tidak dikenali).`:'')}}
 inp.value='';show()}
function csv(){const h=[...M.fitur,'Prediksi','P(Ya) %','P(Tidak) %'];
 const d=S.rows.map((r,i)=>[...M.fitur.map(f=>r[f]),S.out[i].pred,(S.out[i].ya*100).toFixed(2),(S.out[i].tidak*100).toFixed(2)]);
 const a=document.createElement('a');a.href=URL.createObjectURL(new Blob(['\ufeff'+[h,...d].map(x=>x.join(',')).join('\n')],{type:'text/csv'}));a.download='hasil_prediksi.csv';a.click()}
async function pBatch(){
 const r=await api('/api/batch',{rows:S.rows,laplace:S.laplace});S.out=r.hasil;
 const lab=S.rows.some(x=>x.Aktual),ny=S.out.filter(o=>o.pred=='Ya').length,nt=S.out.filter(o=>o.pred=='Tidak').length;
 let h=`<div class="card"><h3>📥 Sumber data uji</h3>${S.msg?`<div class="note ${S.msg.c}">${S.msg.m}</div>`:''}
 <div class="row"><input id="fi" type="file" hidden accept=".xlsx,.xls,.csv" onchange="upload(this)">
 <button class="b" onclick="$('#fi').click()">📥 Impor Excel / CSV</button><a href="/api/template"><button class="b alt">📄 Unduh template Excel</button></a>
 <button class="b alt" onclick="S.rows=M.uji.map(r=>({...r}));S.msg=null;show()">↺ Pakai 17 data soal</button></div>
 <small style="color:var(--mut)">Kolom wajib: Umur, Status, Penghasilan, Hutang Konsumtif. Kolom Investasi (isi ??? / kosong) diabaikan; bila berisi Ya/Tidak dipakai untuk menghitung akurasi.</small>
 <div class="g4" style="margin-top:14px;align-items:end">${M.fitur.map(f=>`<div><label>${f}</label><select onchange="S.nr['${f}']=this.value">${opts(f,S.nr[f])}</select></div>`).join('')}</div>
 <div class="row" style="margin:12px 0 0"><button class="b alt" onclick="S.rows.push({...S.nr});show()">➕ Tambah baris</button></div></div>`;
 if(!S.rows.length)return h+'<div class="note err">Belum ada data uji.</div>';
 h+=`<div class="g3" style="margin-bottom:18px"><div class="card"><div class="l">Total data uji</div><div class="v">${S.rows.length}</div></div>
 <div class="card big ya"><div class="l">Prediksi Ya</div><div class="v">${ny}</div></div><div class="card big tidak"><div class="l">Prediksi Tidak</div><div class="v">${nt}</div></div></div>`;
 let ok=0,tot=0;
 const body=S.rows.map((x,i)=>{const o=S.out[i],c=Math.max(o.ya,o.tidak)*100,col=o.pred=='Ya'?'#10b981':'#f43f5e';let ex='';
  if(lab){const m=x.Aktual?x.Aktual==o.pred:null;if(m!==null){tot++;if(m)ok++}ex=`<td>${badge(x.Aktual)}</td><td>${m===null?'-':m?'✅':'❌'}</td>`}
  return `<tr><td>${i+1}</td>${M.fitur.map(f=>`<td>${x[f]}</td>`).join('')}<td>${badge(o.pred)}</td><td><span class="mini"><i style="width:${c}%;background:${col}"></i></span>${c.toFixed(1)}%</td><td>${pct(o.ya)}</td><td>${pct(o.tidak)}</td>${ex}<td><button class="x" onclick="S.rows.splice(${i},1);show()">✕</button></td></tr>`}).join('');
 h+=`<div class="tw"><table><tr><th>#</th>${M.fitur.map(f=>`<th>${f}</th>`).join('')}<th>Prediksi</th><th>Keyakinan</th><th>P(Ya)</th><th>P(Tidak)</th>${lab?'<th>Aktual</th><th>Cocok</th>':''}<th></th></tr>${body}</table></div>`;
 h+=concl(`Dari ${S.rows.length} data uji, <b>${ny}</b> diprediksi berinvestasi (Ya) dan <b>${nt}</b> diprediksi tidak berinvestasi (Tidak).`+(tot?` Dibandingkan kolom Investasi pada file Anda, ${ok} dari ${tot} prediksi cocok (akurasi ${(ok/tot*100).toFixed(1)}%).`:''));
 return h+'<div class="row" style="margin-top:18px"><button class="b" onclick="csv()">⬇️ Unduh hasil (CSV)</button></div>'}
async function pSteps(){
 const r=await api('/api/predict',{tes:S.tes,laplace:S.laplace});const K=['Ya','Tidak'],st=(n,t)=>`<div class="step"><i>${n}</i>${t}</div>`;
 let h=`<div class="note ok">Perhitungan untuk data uji: <b>${Object.values(S.tes).join(', ')}</b> (ubah di menu Prediksi Tunggal)</div>`;
 h+=st(1,'Probabilitas Prior')+`<div class="g2">${K.map(k=>`<div class="card"><div class="l">P(Investasi = ${k})</div><div class="v">${r.kelas[k].n}/${M.latih.length}</div><div style="color:var(--mut)">= ${r.kelas[k].prior.toFixed(4)}</div></div>`).join('')}</div>`;
 h+=st(2,'Probabilitas Likelihood')+`<div class="g2">${K.map(k=>`<div><b>Kelas Investasi = ${k}</b><div class="tw" style="margin-top:8px"><table><tr><th>Fitur</th><th>Nilai</th><th>Hitung</th><th>P(x|${k})</th></tr>${r.kelas[k].rinci.map(x=>`<tr><td>${x.f}</td><td>${x.v}</td><td>${x.cnt}/${x.nk}</td><td><b>${x.p.toFixed(4)}</b></td></tr>`).join('')}</table></div></div>`).join('')}</div>`;
 h+=st(3,'Skor Akhir')+K.map(k=>`<div class="fm">P(${k}) × ΠP(x|${k}) = ${[r.kelas[k].prior,...r.kelas[k].rinci.map(x=>x.p)].map(x=>x.toFixed(4)).join(' × ')} = <b>${r.kelas[k].skor.toFixed(6)}</b></div>`).join('');
 h+=st(4,'Normalisasi & Keputusan');
 if(r.total)h+=K.map(k=>`<div class="fm">P(${k} | X) = ${r.kelas[k].skor.toFixed(6)} / ${r.total.toFixed(6)} = <b>${pct(r.kelas[k].post)}</b></div>`).join('')+concl(`Skor tertinggi ada pada kelas <b>${r.prediksi}</b>, sehingga prediksi Investasi = <b>${r.prediksi}</b>.`);
 return h}
async function pData(){const ya=M.latih.filter(r=>r[4]=='Ya').length;
 return `<div class="g3" style="margin-bottom:18px"><div class="card"><div class="l">Total data latih</div><div class="v">${M.latih.length}</div></div>
 <div class="card big ya"><div class="l">Investasi Ya</div><div class="v">${ya}</div></div><div class="card big tidak"><div class="l">Investasi Tidak</div><div class="v">${M.latih.length-ya}</div></div></div>
 <div class="tw"><table><tr><th>#</th>${M.fitur.map(f=>`<th>${f}</th>`).join('')}<th>Investasi</th></tr>${M.latih.map((r,i)=>`<tr><td>${i+1}</td>${r.slice(0,4).map(x=>`<td>${x}</td>`).join('')}<td>${badge(r[4])}</td></tr>`).join('')}</table></div>`}
init();
</script></body></html>"""


def main():
    srv = ThreadingHTTPServer(("127.0.0.1", 0), H)
    url = f"http://127.0.0.1:{srv.server_address[1]}"
    print(f"Aplikasi berjalan di {url}  (tekan Ctrl+C untuk berhenti)")
    threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\nAplikasi dihentikan.")


if __name__ == "__main__":
    main()
