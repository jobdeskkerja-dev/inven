# =====================================================================
#  INVENTARIS KELAS (REVISI) - Flask + SQLite, siap jalan di Google Colab
#  Cara pakai: tempel SELURUH kode ini ke 1 sel Colab, lalu Run.
#  Fitur: foto kerusakan, notifikasi langsung (pop-up), tampilan baru.
#  Tidak perlu izin Google Drive (data sementara). Untuk menyimpan permanen, ubah SIMPAN_DI_DRIVE = True.
#  Klik tautan yang muncul di bawah sel untuk membuka aplikasinya.
#  Akun awal:  admin / admin123  |  guru / guru123  |  siswa / siswa123
#  (segera ganti password setelah login pertama)
# =====================================================================
import os, io, re, csv, time, hmac, secrets, sqlite3, threading
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo
from functools import wraps
from flask import (Flask, render_template, request, redirect, url_for,
                   session, flash, Response, send_from_directory, jsonify, abort)
from jinja2 import DictLoader
from markupsafe import Markup
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.serving import make_server

def _pastikan(modul, paket):                      # pasang library Excel/PDF bila belum ada
    try: __import__(modul)
    except ImportError:
        import subprocess, sys
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", paket])
_pastikan("openpyxl", "openpyxl"); _pastikan("reportlab", "reportlab"); _pastikan("PIL", "pillow")
from PIL import Image, ImageOps

# ---- PENGATURAN: ubah sesuai kebutuhanmu --------------------------------
NAMA_APP = "InventClass"        # <- ganti dengan nama aplikasi pilihanmu
NAMA_SEKOLAH = "NAMA SEKOLAH"        # dipakai pada ekspor Excel/PDF resmi
NAMA_KELAS = "KELAS"
TTD_KIRI, TTD_KANAN = "Wali Kelas", "Ketua Kelas"
SIMPAN_DI_DRIVE = False              # False = tanpa izin Drive (data sementara, cocok untuk demo/dikumpulkan). True = simpan di Google Drive
# ---------------------------------------------------------------------------

CATS = ["Furnitur", "Alat Tulis", "Alat Kebersihan", "Elektronik", "Dekorasi", "Lainnya"]
KOND = {"baik": "Baik (layak dipakai)", "ringan": "Rusak Ringan (perlu perbaikan)",
        "berat": "Rusak Berat (perlu pengadaan)"}
KELAS = ["10", "11", "12"]
ROLES = {"admin": "Admin", "guru": "Guru", "anggota": "Siswa"}
PW_BAWAAN = {"admin": "admin123", "guru": "guru123", "siswa": "siswa123"}

# ---------------------------------------------------------------- waktu (WIB)
WIB = ZoneInfo("Asia/Jakarta")
def now(): return datetime.now(WIB)
def today(): return now().date()

# ---------------------------------------------------------------- lokasi data
def _tentukan_folder():
    folder = os.getcwd()
    if SIMPAN_DI_DRIVE:
        try:
            from google.colab import drive
            drive.mount("/content/drive")
            folder = "/content/drive/MyDrive/inventaris_kelas"
            os.makedirs(folder, exist_ok=True)
        except Exception as e:
            print("Google Drive tidak terpasang, data disimpan di folder sementara:", e)
    return folder

FOLDER = _tentukan_folder()
DB = os.path.join(FOLDER, "inventaris.db")
FOTO_DIR = os.path.join(FOLDER, "foto")          # foto kerusakan yang diunggah
os.makedirs(FOTO_DIR, exist_ok=True)

def _kunci_rahasia():                              # kunci sesi acak, disimpan agar login tidak hilang saat restart
    k = os.environ.get("SECRET_KEY")
    if k: return k
    p = os.path.join(FOLDER, "rahasia.key")
    try:
        if os.path.exists(p):
            with open(p) as f: return f.read().strip()
        k = secrets.token_hex(32)
        with open(p, "w") as f: f.write(k)
        return k
    except OSError:
        return secrets.token_hex(32)

app = Flask(__name__)
app.secret_key = _kunci_rahasia()
app.config.update(SESSION_COOKIE_SAMESITE="Lax", SESSION_COOKIE_HTTPONLY=True, MAX_CONTENT_LENGTH=8 * 1024 * 1024)

# ---------------------------------------------------------------- database
def q(sql, a=(), one=False):
    c = sqlite3.connect(DB); c.row_factory = sqlite3.Row
    r = c.execute(sql, a).fetchall(); c.close()
    return (r[0] if r else None) if one else r

def ex(sql, a=()):
    c = sqlite3.connect(DB); c.execute(sql, a); c.commit(); c.close()

def log(aksi):
    ex("INSERT INTO log(waktu,user,aksi) VALUES(?,?,?)",
       (now().strftime("%Y-%m-%d %H:%M"), session.get("u", "-"), aksi))

def num(x, default=0):
    try: return max(0, int(x))
    except (TypeError, ValueError): return default

def init_db():
    c = sqlite3.connect(DB)
    c.executescript("""
    CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, username TEXT UNIQUE, pw TEXT, role TEXT);
    CREATE TABLE IF NOT EXISTS items(id INTEGER PRIMARY KEY, nama TEXT, kategori TEXT, lokasi TEXT,
        baik INT DEFAULT 0, ringan INT DEFAULT 0, berat INT DEFAULT 0, tgl TEXT, catatan TEXT, aktif INT DEFAULT 1);
    CREATE TABLE IF NOT EXISTS loans(id INTEGER PRIMARY KEY, item_id INT, peminjam TEXT, jumlah INT,
        tgl_pinjam TEXT, tgl_kembali TEXT, status TEXT DEFAULT 'Dipinjam', kelas TEXT, rombel TEXT);
    CREATE TABLE IF NOT EXISTS reports(id INTEGER PRIMARY KEY, item_id INT, pelapor TEXT, deskripsi TEXT,
        tgl TEXT, status TEXT DEFAULT 'Menunggu', foto TEXT);
    CREATE TABLE IF NOT EXISTS log(id INTEGER PRIMARY KEY, waktu TEXT, user TEXT, aksi TEXT);
    CREATE TABLE IF NOT EXISTS notif(id INTEGER PRIMARY KEY, user TEXT, pesan TEXT, link TEXT, waktu TEXT, dibaca INT DEFAULT 0, kunci TEXT);
    CREATE UNIQUE INDEX IF NOT EXISTS ux_notif ON notif(user, kunci);
    CREATE TABLE IF NOT EXISTS minat(id INTEGER PRIMARY KEY, user TEXT, item_id INT, jumlah INT DEFAULT 1, dibuat TEXT, dikabari INT DEFAULT 0);
    """)
    for tabel, kolom, tipe in (("loans", "kelas", "TEXT"), ("loans", "rombel", "TEXT"),   # migrasi database lama
                               ("loans", "user", "TEXT"), ("items", "aktif", "INT DEFAULT 1"),
                               ("reports", "foto", "TEXT")):
        ada = [r[1] for r in c.execute(f"PRAGMA table_info({tabel})")]
        if kolom not in ada: c.execute(f"ALTER TABLE {tabel} ADD COLUMN {kolom} {tipe}")
    if not c.execute("SELECT 1 FROM users").fetchone():          # akun & data contoh HANYA dibuat saat database masih kosong
        c.executemany("INSERT INTO users(username,pw,role) VALUES(?,?,?)", [
            ("admin", generate_password_hash(PW_BAWAAN["admin"]), "admin"),
            ("guru", generate_password_hash(PW_BAWAAN["guru"]), "guru"),
            ("siswa", generate_password_hash(PW_BAWAAN["siswa"]), "anggota")])
        hari = today().isoformat()
        c.executemany("INSERT INTO items(nama,kategori,lokasi,baik,ringan,berat,tgl,catatan) VALUES(?,?,?,?,?,?,?,?)", [
            ("Kursi", "Furnitur", "Area siswa", 810, 28, 12, hari, ""),
            ("Meja", "Furnitur", "Area siswa", 822, 15, 13, hari, ""),
            ("Papan Tulis", "Alat Tulis", "Depan kelas", 34, 2, 0, hari, ""),
            ("Spidol Whiteboard", "Alat Tulis", "Meja guru", 25, 0, 3, hari, "Stok habis pakai"),
            ("Sapu", "Alat Kebersihan", "Pojok belakang", 2, 1, 0, hari, ""),
            ("Proyektor", "Elektronik", "Depan kelas", 1, 0, 0, hari, ""),
            ("Jam Dinding", "Dekorasi", "Dinding depan", 25, 1, 0, hari, "")])
        c.executemany("INSERT INTO loans(item_id,peminjam,jumlah,tgl_pinjam,tgl_kembali,status,kelas,rombel) VALUES(?,?,?,?,?,?,?,?)", [
            (1, "Nadia", 35, hari, hari, "Dipinjam", "10", "X-1"),
            (3, "Salsa", 2, hari, hari, "Dipinjam", "10", "X-3"),
            (5, "Raka", 1, hari, hari, "Dikembalikan", "11", "XI IPA 2"),
            (4, "Dimas", 1, hari, hari, "Dikembalikan", "12", "XII IPS 1"),
            (5, "Bagas", 1, hari, hari, "Dikembalikan", "12", "XII IPA 3")])
    c.commit(); c.close()

# ---------------------------------------------------------------- keamanan (CSRF & batas percobaan login)
def csrf():
    if "_t" not in session: session["_t"] = secrets.token_hex(16)
    return session["_t"]

def csrf_input():
    return Markup(f'<input type="hidden" name="_t" value="{csrf()}">')

_gagal = {}
def terkunci(kunci):                               # maks 5 kegagalan login per 5 menit per username
    _gagal[kunci] = [t for t in _gagal.get(kunci, []) if time.time() - t < 300]
    return len(_gagal[kunci]) >= 5

# ---------------------------------------------------------------- auth & helper
def need(admin=False, staff=False):
    def deco(f):
        @wraps(f)
        def w(*a, **k):
            if "u" not in session: return redirect("/login")
            r = session.get("role")
            if admin and r != "admin":
                flash("Fitur ini hanya untuk admin."); return redirect("/")
            if staff and r not in ("admin", "guru"):
                flash("Fitur ini hanya untuk admin dan guru."); return redirect("/")
            return f(*a, **k)
        return w
    return deco

def dipinjam(item_id):
    return q("SELECT COALESCE(SUM(jumlah),0) s FROM loans WHERE item_id=? AND status='Dipinjam'", (item_id,), one=True)["s"]

def available(item_id):
    it = q("SELECT baik FROM items WHERE id=? AND aktif=1", (item_id,), one=True)
    return (it["baik"] - dipinjam(item_id)) if it else 0

def staff_users():
    return [r["username"] for r in q("SELECT username FROM users WHERE role IN ('admin','guru')")]

Image.MAX_IMAGE_PIXELS = 40_000_000

def simpan_foto(f):                                # validasi & simpan foto kerusakan -> (nama_file, pesan_error)
    if not f or not f.filename: return None, None
    try:
        im = Image.open(f.stream)
        if im.format not in ("JPEG", "PNG", "WEBP", "GIF"): raise ValueError("format")
        im.load()
        im = ImageOps.exif_transpose(im).convert("RGB")      # putar sesuai EXIF, buang metadata
        im.thumbnail((1280, 1280))
        nama = secrets.token_hex(16) + ".jpg"
        im.save(os.path.join(FOTO_DIR, nama), "JPEG", quality=85, optimize=True)
        return nama, None
    except Exception:
        return None, "File harus berupa foto yang valid (JPG, PNG, WEBP, atau GIF)."

def minat_tambah(iid, n=1):                        # daftar "kabari saya" saat barang tersedia
    if not q("SELECT 1 FROM items WHERE id=? AND aktif=1", (iid,), one=True): return False
    m = q("SELECT id FROM minat WHERE user=? AND item_id=? AND dikabari=0", (session["u"], iid), one=True)
    if m: ex("UPDATE minat SET jumlah=? WHERE id=?", (max(1, n), m["id"]))
    else: ex("INSERT INTO minat(user,item_id,jumlah,dibuat) VALUES(?,?,?,?)", (session["u"], iid, max(1, n), now().strftime("%Y-%m-%d %H:%M")))
    return True

def cek_tersedia():                                # kabari siswa yang menunggu bila stok sudah cukup
    for m in q("SELECT m.*, i.nama FROM minat m JOIN items i ON i.id=m.item_id WHERE m.dikabari=0 AND i.aktif=1"):
        sedia = available(m["item_id"])
        if sedia >= max(1, m["jumlah"] or 1):
            notif(m["user"], f"Barang tersedia: {m['nama']} sekarang ada {sedia} unit. Silakan pinjam.", "/pinjam", f"ada-{m['id']}", False)
            ex("UPDATE minat SET dikabari=1 WHERE id=?", (m["id"],))

def notif(users, pesan, link="/notif", kunci=None, self_skip=True):
    users = [users] if isinstance(users, str) else users
    c = sqlite3.connect(DB)
    for u in users:
        if self_skip and u == session.get("u"): continue
        c.execute("INSERT OR IGNORE INTO notif(user,pesan,link,waktu,dibaca,kunci) VALUES(?,?,?,?,0,?)",
                  (u, pesan, link, now().strftime("%Y-%m-%d %H:%M"), kunci))
    c.commit(); c.close()

_sync = [0]
def sinkron_notif():                               # notifikasi otomatis: terlambat & pengingat
    if time.time() - _sync[0] < 20: return
    _sync[0] = time.time()
    hari = today().isoformat(); besok = (today() + timedelta(days=1)).isoformat(); stf = staff_users()
    for l in q("SELECT l.*, i.nama FROM loans l JOIN items i ON i.id=l.item_id WHERE l.status='Dipinjam'"):
        if l["tgl_kembali"] < hari:
            notif(stf, f"Terlambat: {l['nama']} dipinjam {l['peminjam']} (kelas {l['kelas'] or '-'}), batas {l['tgl_kembali']}.", "/pinjam", f"late-{l['id']}", False)
            if l["user"]: notif(l["user"], f"Peminjamanmu ({l['nama']}) melewati batas {l['tgl_kembali']}. Segera kembalikan.", "/pinjam", f"telat-{l['id']}", False)
        elif l["user"] and l["tgl_kembali"] in (hari, besok):
            kapan = "hari ini" if l["tgl_kembali"] == hari else "besok"
            notif(l["user"], f"Pengingat: {l['nama']} harus dikembalikan {kapan}.", "/pinjam", f"ingat-{l['id']}-{kapan}", False)
    cek_tersedia()

@app.before_request
def _sebelum():
    if request.method == "POST":                   # tolak formulir tanpa token CSRF yang benar
        t = request.form.get("_t", "").encode()
        if not t or not hmac.compare_digest(t, session.get("_t", "").encode()):
            flash("Sesi formulir kedaluwarsa, silakan ulangi."); return redirect("/")
    if "u" in session and request.method == "GET": sinkron_notif()

@app.context_processor
def _konteks():
    u, role = session.get("u"), session.get("role")
    unread = q("SELECT COUNT(*) c FROM notif WHERE user=? AND dibaca=0", (u,), one=True)["c"] if u else 0
    last = q("SELECT COALESCE(MAX(id),0) m FROM notif WHERE user=?", (u,), one=True)["m"] if u else 0
    lapn = q("SELECT COUNT(*) c FROM reports WHERE status='Menunggu'", one=True)["c"] if role in ("admin", "guru") else 0
    return dict(staff=role in ("admin", "guru"), is_admin=role == "admin", unread=unread, roles=ROLES,
                last_notif=last, lapn=lapn, hari_ini=tgl_id())

@app.route("/notif")
@need()
def notif_page():
    return render_template("notif", rows=q("SELECT * FROM notif WHERE user=? ORDER BY id DESC LIMIT 100", (session["u"],)))

@app.route("/notif/<int:i>")
@need()
def notif_buka(i):
    n = q("SELECT * FROM notif WHERE id=? AND user=?", (i, session["u"]), one=True)
    if not n: return redirect("/notif")
    ex("UPDATE notif SET dibaca=1 WHERE id=?", (i,)); return redirect(n["link"] or "/notif")

@app.route("/notif/baca-semua", methods=["POST"])
@need()
def notif_baca():
    ex("UPDATE notif SET dibaca=1 WHERE user=?", (session["u"],)); return redirect("/notif")

@app.route("/notif/cek")
@need()
def notif_cek():                                    # dipanggil browser tiap 15 detik untuk pop-up notifikasi
    sejak = num(request.args.get("sejak"))
    rows = q("SELECT id,pesan FROM notif WHERE user=? AND dibaca=0 AND id>? ORDER BY id", (session["u"], sejak))
    unread = q("SELECT COUNT(*) c FROM notif WHERE user=? AND dibaca=0", (session["u"],), one=True)["c"]
    return jsonify(unread=unread, baru=[dict(id=r["id"], pesan=r["pesan"], link=f"/notif/{r['id']}") for r in rows])

@app.route("/foto/<nama>")
@need()
def foto(nama):
    if not re.fullmatch(r"[a-f0-9]{32}\.jpg", nama): abort(404)
    return send_from_directory(FOTO_DIR, nama, mimetype="image/jpeg", max_age=86400)

@app.route("/minat/<int:i>", methods=["POST"])
@need()
def minat(i):
    if minat_tambah(i): flash("Siap! Kamu akan dikabari lewat notifikasi saat barang ini tersedia.")
    return redirect("/barang")

@app.errorhandler(413)
def _terlalu_besar(e):
    flash("Foto terlalu besar (maksimal 8 MB)."); return redirect("/laporan")

@app.route("/pengguna", methods=["GET", "POST"])
@need(admin=True)
def pengguna():
    if request.method == "POST":
        u, pw, role = request.form["username"].strip(), request.form["password"], request.form["role"]
        if not re.fullmatch(r"[A-Za-z0-9_.-]{3,30}", u): flash("Username 3-30 karakter (huruf, angka, titik, strip, underscore).")
        elif len(pw) < 6: flash("Password minimal 6 karakter.")
        elif role not in ROLES: flash("Peran tidak valid.")
        elif q("SELECT 1 FROM users WHERE username=?", (u,), one=True): flash("Username sudah dipakai.")
        else:
            ex("INSERT INTO users(username,pw,role) VALUES(?,?,?)", (u, generate_password_hash(pw), role))
            log(f"Tambah akun {u} ({ROLES[role]})"); flash("Akun ditambahkan.")
        return redirect("/pengguna")
    return render_template("pengguna", users=q("SELECT * FROM users ORDER BY role, username"))

@app.route("/pengguna/<int:i>/password", methods=["POST"])
@need(admin=True)
def pengguna_pw(i):
    u, pw = q("SELECT * FROM users WHERE id=?", (i,), one=True), request.form["password"]
    if not u: return redirect("/pengguna")
    if len(pw) < 6: flash("Password minimal 6 karakter.")
    else:
        ex("UPDATE users SET pw=? WHERE id=?", (generate_password_hash(pw), i)); log(f"Reset password akun {u['username']}"); flash("Password diganti.")
    return redirect("/pengguna")

@app.route("/pengguna/<int:i>/hapus", methods=["POST"])
@need(admin=True)
def pengguna_hapus(i):
    u = q("SELECT * FROM users WHERE id=?", (i,), one=True)
    if not u or u["username"] == session["u"]: flash("Akun ini tidak bisa dihapus.")
    elif u["role"] == "admin" and q("SELECT COUNT(*) c FROM users WHERE role='admin'", one=True)["c"] < 2: flash("Admin terakhir tidak bisa dihapus.")
    else:
        ex("DELETE FROM users WHERE id=?", (i,)); ex("DELETE FROM notif WHERE user=?", (u["username"],))
        log(f"Hapus akun {u['username']}"); flash("Akun dihapus.")
    return redirect("/pengguna")

@app.route("/akun", methods=["GET", "POST"])
@need()
def akun():
    if request.method == "POST":
        u, f = q("SELECT * FROM users WHERE username=?", (session["u"],), one=True), request.form
        if not check_password_hash(u["pw"], f["lama"]): flash("Password lama salah.")
        elif len(f["baru"]) < 6: flash("Password baru minimal 6 karakter.")
        elif f["baru"] != f["ulang"]: flash("Konfirmasi password tidak sama.")
        else:
            ex("UPDATE users SET pw=? WHERE id=?", (generate_password_hash(f["baru"]), u["id"])); log("Ganti password sendiri")
            flash("Password berhasil diganti.")
        return redirect("/akun")
    return render_template("akun")

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        nama, pw = request.form["username"].strip(), request.form["password"]
        if terkunci(nama.lower()):
            flash("Terlalu banyak percobaan gagal. Coba lagi dalam 5 menit.")
            return render_template("login")
        u = q("SELECT * FROM users WHERE username=?", (nama,), one=True)
        if u and check_password_hash(u["pw"], pw):
            _gagal.pop(nama.lower(), None)
            token = session.get("_t"); session.clear(); session["_t"] = token or secrets.token_hex(16)
            session["u"], session["role"] = u["username"], u["role"]
            log("Login"); return redirect("/")
        _gagal.setdefault(nama.lower(), []).append(time.time())
        flash("Username atau password salah.")
    return render_template("login")

@app.route("/logout")
def logout():
    session.clear(); return redirect("/login")

# ---------------------------------------------------------------- dashboard
@app.route("/")
@need()
def dash():
    items = q("SELECT * FROM items WHERE aktif=1 ORDER BY nama")
    s = {k: sum(i[k] for i in items) for k in ("baik", "ringan", "berat")}
    s["total"] = s["baik"] + s["ringan"] + s["berat"]
    s["pinjam"] = q("SELECT COALESCE(SUM(jumlah),0) s FROM loans WHERE status='Dipinjam'", one=True)["s"]
    s["lap"] = q("SELECT COUNT(*) c FROM reports WHERE status='Menunggu'", one=True)["c"]
    perhatian = [i for i in items if i["ringan"] or i["berat"]]
    # diagram donat kondisi
    segs, cum = [], 0
    for k, col, lab in (("baik", "#3f9d6b", "Baik"), ("ringan", "#e0a526", "Rusak ringan"), ("berat", "#c0392b", "Rusak berat")):
        p = s[k] * 100 / s["total"] if s["total"] else 0
        segs.append(dict(p=p, off=25 - cum, col=col, lab=lab, n=s[k])); cum += p
    # diagram batang per kategori
    rows = q("SELECT kategori, SUM(baik) b, SUM(ringan) r, SUM(berat) w FROM items WHERE aktif=1 GROUP BY kategori ORDER BY kategori")
    mx = max([r["b"] + r["r"] + r["w"] for r in rows] or [1]) or 1
    kat = []
    for r in rows:
        t = r["b"] + r["r"] + r["w"]
        if t: kat.append(dict(nama=r["kategori"], tot=t, w=t * 100 / mx, bp=r["b"] * 100 / t, rp=r["r"] * 100 / t, wp=r["w"] * 100 / t))
    # diagram peminjaman per kelas (10-12)
    lk = {r["kelas"]: r for r in q("""SELECT kelas, COUNT(*) total,
          SUM(CASE WHEN status='Dipinjam' THEN jumlah ELSE 0 END) aktif FROM loans WHERE kelas IS NOT NULL GROUP BY kelas""")}
    mk = max([max(r["total"], r["aktif"]) for r in lk.values()] or [1]) or 1
    kc = []
    for k in KELAS:
        r = lk.get(k); t = r["total"] if r else 0; a = r["aktif"] if r else 0
        kc.append(dict(kelas=k, total=t, aktif=a, h1=t * 100 / mk, h2=a * 100 / mk))
    # barang paling sering dipinjam
    tp = q("SELECT i.nama, SUM(l.jumlah) n FROM loans l JOIN items i ON i.id=l.item_id GROUP BY i.id ORDER BY n DESC LIMIT 5")
    mt = max([r["n"] for r in tp] or [1]) or 1
    top = [dict(nama=r["nama"], n=r["n"], w=r["n"] * 100 / mt) for r in tp]
    late = q("SELECT COUNT(*) c FROM loans WHERE status='Dipinjam' AND tgl_kembali<?", (today().isoformat(),), one=True)["c"]
    jam = now().hour
    sapa = "Selamat pagi" if jam < 11 else "Selamat siang" if jam < 15 else "Selamat sore" if jam < 18 else "Selamat malam"
    return render_template("dash", s=s, perhatian=perhatian, segs=segs, kat=kat, kc=kc, top=top, late=late, sapa=sapa)

# ---------------------------------------------------------------- barang
@app.route("/barang")
@need()
def barang():
    qq, kat, kon = request.args.get("q", ""), request.args.get("kat", ""), request.args.get("kon", "")
    sql, a = "SELECT * FROM items WHERE aktif=1 AND (nama LIKE ? OR lokasi LIKE ?)", [f"%{qq}%", f"%{qq}%"]
    if kat: sql += " AND kategori=?"; a.append(kat)
    if kon in KOND: sql += f" AND {kon}>0"
    items = [dict(i, tersedia=available(i["id"])) for i in q(sql + " ORDER BY nama", a)]
    mine = {m["item_id"] for m in q("SELECT item_id FROM minat WHERE user=? AND dikabari=0", (session["u"],))}
    return render_template("barang", items=items, mine=mine, qq=qq, kat=kat, kon=kon)

def baca_form():
    f = request.form
    return (f["nama"].strip(), f["kategori"], f["lokasi"].strip(), num(f["baik"]), num(f["ringan"]),
            num(f["berat"]), f["tgl"] or today().isoformat(), f["catatan"].strip())

@app.route("/barang/tambah", methods=["GET", "POST"])
@need(admin=True)
def tambah():
    if request.method == "POST":
        d = baca_form()
        if not d[0]: flash("Nama barang wajib diisi.")
        else:
            ex("INSERT INTO items(nama,kategori,lokasi,baik,ringan,berat,tgl,catatan) VALUES(?,?,?,?,?,?,?,?)", d)
            log(f"Tambah barang: {d[0]} (baik {d[3]}, ringan {d[4]}, berat {d[5]})")
            flash("Barang ditambahkan."); return redirect("/barang")
    return render_template("form", it=None, today=today().isoformat())

@app.route("/barang/<int:i>/ubah", methods=["GET", "POST"])
@need(admin=True)
def ubah(i):
    it = q("SELECT * FROM items WHERE id=? AND aktif=1", (i,), one=True)
    if not it: return redirect("/barang")
    if request.method == "POST":
        d = baca_form(); out = dipinjam(i)
        if not d[0]: flash("Nama barang wajib diisi.")
        elif d[3] < out:
            flash(f"Jumlah Baik tidak boleh kurang dari unit yang sedang dipinjam ({out}). Tandai dulu pengembaliannya.")
        else:
            ex("UPDATE items SET nama=?,kategori=?,lokasi=?,baik=?,ringan=?,berat=?,tgl=?,catatan=? WHERE id=?", d + (i,))
            log(f"Ubah barang: {d[0]} (baik {it['baik']}->{d[3]}, ringan {it['ringan']}->{d[4]}, berat {it['berat']}->{d[5]})")
            if d[5] > it["berat"]: notif(staff_users(), f"Barang rusak berat bertambah: {d[0]} perlu pengadaan.", "/tindak")
            cek_tersedia(); flash("Barang diperbarui."); return redirect("/barang")
    return render_template("form", it=it, today=today().isoformat())

@app.route("/barang/<int:i>/hapus", methods=["POST"])
@need(admin=True)
def hapus(i):                                       # soft delete: barang diarsipkan, riwayat peminjaman tetap utuh
    it = q("SELECT nama FROM items WHERE id=? AND aktif=1", (i,), one=True)
    if not it: return redirect("/barang")
    if dipinjam(i) > 0: flash("Barang masih dipinjam, tidak bisa dihapus. Tandai dulu pengembaliannya.")
    else:
        ex("UPDATE items SET aktif=0 WHERE id=?", (i,))
        log(f"Hapus barang: {it['nama']} (diarsipkan)"); flash("Barang dihapus. Riwayat peminjamannya tetap tersimpan.")
    return redirect("/barang")

# ---------------------------------------------------------------- perbaikan & pengadaan
@app.route("/tindak")
@need()
def tindak():
    return render_template("tindak",
        rep=q("SELECT * FROM items WHERE aktif=1 AND ringan>0 ORDER BY nama"),
        adq=q("SELECT * FROM items WHERE aktif=1 AND berat>0 ORDER BY nama"))

@app.route("/tindak/<int:i>/<jenis>", methods=["POST"])
@need(staff=True)
def tindak_do(i, jenis):
    if jenis not in ("perbaikan", "pengadaan"): return redirect("/tindak")
    it = q("SELECT * FROM items WHERE id=? AND aktif=1", (i,), one=True)
    src = "ringan" if jenis == "perbaikan" else "berat"
    n = num(request.form.get("n"), 1)
    if not it or n < 1 or n > it[src]: flash("Jumlah tidak valid.")
    else:
        ex(f"UPDATE items SET {src}={src}-?, baik=baik+? WHERE id=?", (n, n, i))
        log(f"{n} {it['nama']} " + ("selesai diperbaiki" if jenis == "perbaikan" else "sudah diadakan/diganti"))
        sebelum, sesudah = ("Diterima: Rusak Ringan", "Selesai diperbaiki") if jenis == "perbaikan" else ("Diterima: Rusak Berat", "Selesai diadakan/diganti")
        kata = "sudah diperbaiki" if jenis == "perbaikan" else "sudah diganti/diadakan"
        for r in q("SELECT * FROM reports WHERE item_id=? AND status=? ORDER BY id LIMIT ?", (i, sebelum, n)):      # kabari siswa pelapor
            ex("UPDATE reports SET status=? WHERE id=?", (sesudah, r["id"]))
            notif(r["pelapor"], f"Kabar baik: {it['nama']} yang kamu laporkan {kata} dan siap dipakai lagi.", "/laporan", f"fix-{r['id']}", False)
        cek_tersedia()
        flash("Status diperbarui: kembali ke kondisi Baik.")
    return redirect("/tindak")

# ---------------------------------------------------------------- peminjaman
@app.route("/pinjam", methods=["GET", "POST"])
@need()
def pinjam():
    if request.method == "POST":
        iid, n = num(request.form["item_id"]), num(request.form["jumlah"])
        nama = request.form["peminjam"].strip() or session["u"]
        kelas, rombel = request.form.get("kelas", ""), request.form.get("rombel", "").strip()
        try: tk = date.fromisoformat(request.form["tgl_kembali"])
        except ValueError: tk = None
        if kelas not in KELAS: flash("Pilih kelas peminjam (10, 11, atau 12).")
        elif tk is None or tk < today(): flash("Tanggal rencana kembali tidak valid.")
        elif n < 1: flash("Jumlah minimal 1.")
        elif n > available(iid):
            if minat_tambah(iid, n): flash("Jumlah melebihi stok yang tersedia (kondisi Baik). Kamu akan dikabari lewat notifikasi saat stoknya cukup.")
            else: flash("Barang tidak ditemukan.")
        else:
            ex("INSERT INTO loans(item_id,peminjam,jumlah,tgl_pinjam,tgl_kembali,kelas,rombel,user) VALUES(?,?,?,?,?,?,?,?)",
               (iid, nama, n, today().isoformat(), tk.isoformat(), kelas, rombel, session["u"]))
            log(f"Pinjam: {nama} (kelas {kelas} {rombel}) meminjam {n} (barang #{iid})"); flash("Peminjaman dicatat.")
        return redirect("/pinjam")
    items = [dict(i, tersedia=available(i["id"])) for i in q("SELECT * FROM items WHERE aktif=1 AND baik>0 ORDER BY nama")]
    kf = request.args.get("kelas", "")
    sql, a = "SELECT l.*, i.nama FROM loans l JOIN items i ON i.id=l.item_id", []
    if kf in KELAS: sql += " WHERE l.kelas=?"; a.append(kf)
    loans = q(sql + " ORDER BY (l.status='Dipinjam') DESC, l.id DESC", a)
    return render_template("pinjam", items=items, loans=loans, kf=kf, today=today().isoformat())

@app.route("/pinjam/<int:i>/kembali", methods=["POST"])
@need(staff=True)
def kembali(i):
    ex("UPDATE loans SET status='Dikembalikan' WHERE id=?", (i,)); log(f"Pengembalian pinjaman #{i}"); cek_tersedia()
    flash("Barang dikembalikan."); return redirect("/pinjam")

# ---------------------------------------------------------------- laporan kerusakan
@app.route("/laporan", methods=["GET", "POST"])
@need()
def laporan():
    if request.method == "POST":
        nm = q("SELECT nama FROM items WHERE id=? AND aktif=1", (num(request.form["item_id"]),), one=True)
        desk = request.form["deskripsi"].strip()
        if not nm: flash("Barang tidak ditemukan.")
        elif not desk: flash("Deskripsi kerusakan wajib diisi.")
        else:
            foto, err = simpan_foto(request.files.get("foto"))
            if err: flash(err)
            else:
                ex("INSERT INTO reports(item_id,pelapor,deskripsi,tgl,foto) VALUES(?,?,?,?,?)",
                   (num(request.form["item_id"]), session["u"], desk, today().isoformat(), foto))
                notif(staff_users(), f"Laporan kerusakan baru: {nm['nama']} (oleh {session['u']})" + (", dengan foto." if foto else "."), "/laporan")
                log("Laporan kerusakan dikirim" + (" (dengan foto)" if foto else "")); flash("Laporan terkirim, menunggu penilaian admin atau guru.")
        return redirect("/laporan")
    return render_template("laporan", items=q("SELECT * FROM items WHERE aktif=1 ORDER BY nama"),
        reps=q("SELECT r.*, i.nama FROM reports r JOIN items i ON i.id=r.item_id ORDER BY (r.status='Menunggu') DESC, r.id DESC"))

@app.route("/laporan/<int:i>/<aksi>", methods=["POST"])
@need(staff=True)
def laporan_do(i, aksi):
    if aksi not in ("tolak", "ringan", "berat"): return redirect("/laporan")
    r = q("SELECT * FROM reports WHERE id=?", (i,), one=True)
    if not r or r["status"] != "Menunggu": return redirect("/laporan")
    if aksi == "tolak":
        ex("UPDATE reports SET status='Ditolak' WHERE id=?", (i,))
    else:
        it = q("SELECT * FROM items WHERE id=?", (r["item_id"],), one=True)
        if not it or available(it["id"]) < 1:
            flash("Tidak ada unit Baik yang tersedia untuk dipindahkan (mungkin sedang dipinjam)."); return redirect("/laporan")
        ex(f"UPDATE items SET baik=baik-1, {aksi}={aksi}+1 WHERE id=?", (it["id"],))
        ex("UPDATE reports SET status=? WHERE id=?", ("Diterima: Rusak " + aksi.capitalize(), i))
        log(f"Laporan #{i} diterima: 1 {it['nama']} jadi rusak {aksi}")
    r2 = q("SELECT r.status, i.nama FROM reports r JOIN items i ON i.id=r.item_id WHERE r.id=?", (i,), one=True)
    if r2:
        notif(r["pelapor"], f"Laporanmu tentang {r2['nama']} sudah dinilai: {r2['status']}.", "/laporan")
        if aksi == "berat": notif(staff_users(), f"Barang rusak berat bertambah: {r2['nama']} perlu pengadaan.", "/tindak")
    flash("Laporan diproses."); return redirect("/laporan")

# ---------------------------------------------------------------- riwayat & ekspor
@app.route("/riwayat")
@need(staff=True)
def riwayat():
    return render_template("riwayat", logs=q("SELECT * FROM log ORDER BY id DESC LIMIT 200"))

BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
JENIS = ("inventaris", "perbaikan", "pengadaan", "peminjaman")

def tgl_id(d=None):
    d = d or today(); return f"{d.day} {BULAN[d.month - 1]} {d.year}"

def dataset(jenis, kelas=""):
    hari = today().isoformat()
    if jenis == "inventaris":
        rows = [[n, i["nama"], i["kategori"], i["lokasi"] or "-", i["baik"], i["ringan"], i["berat"], i["baik"] + i["ringan"] + i["berat"], i["tgl"]]
                for n, i in enumerate(q("SELECT * FROM items WHERE aktif=1 ORDER BY kategori, nama"), 1)]
        sm = [sum(r[c] for r in rows) for c in (4, 5, 6, 7)]
        return dict(judul="Daftar Inventaris Barang Kelas", rows=rows, total=["", "JUMLAH", "", ""] + sm + [""],
            head=["No", "Nama Barang", "Kategori", "Lokasi", "Baik", "Rusak Ringan", "Rusak Berat", "Total", "Tgl Masuk"],
            ringkas=[("Jenis barang", len(rows)), ("Total unit", sm[3]), ("Layak dipakai (Baik)", sm[0]),
                     ("Perlu perbaikan (Rusak Ringan)", sm[1]), ("Perlu pengadaan (Rusak Berat)", sm[2])])
    if jenis in ("perbaikan", "pengadaan"):
        col, ket = ("ringan", "Perlu perbaikan") if jenis == "perbaikan" else ("berat", "Perlu pengadaan")
        rows = [[n, i["nama"], i["kategori"], i["lokasi"] or "-", i[col], ket] for n, i in enumerate(q(f"SELECT * FROM items WHERE aktif=1 AND {col}>0 ORDER BY nama"), 1)]
        jml = sum(r[4] for r in rows)
        return dict(judul="Daftar Barang Perlu Perbaikan" if jenis == "perbaikan" else "Daftar Kebutuhan Pengadaan Barang",
            head=["No", "Nama Barang", "Kategori", "Lokasi", "Jumlah (unit)", "Keterangan"], rows=rows, total=["", "JUMLAH", "", "", jml, ""],
            ringkas=[("Jenis barang", len(rows)), ("Total unit", jml)])
    kl = kelas if kelas in KELAS else ""
    sql = "SELECT l.*, i.nama FROM loans l JOIN items i ON i.id=l.item_id" + (" WHERE l.kelas=?" if kl else "") + " ORDER BY l.tgl_pinjam DESC, l.id DESC"
    ls = q(sql, (kl,) if kl else ())
    def st(l): return "Dikembalikan" if l["status"] != "Dipinjam" else ("Terlambat" if l["tgl_kembali"] < hari else "Dipinjam")
    rows = [[n, l["tgl_pinjam"], l["nama"], l["peminjam"], ((l["kelas"] or "-") + " " + (l["rombel"] or "")).strip(), l["jumlah"], l["tgl_kembali"], st(l)]
            for n, l in enumerate(ls, 1)]
    return dict(judul="Daftar Peminjaman Barang" + (f" - Kelas {kl}" if kl else ""), rows=rows,
        head=["No", "Tgl Pinjam", "Barang", "Peminjam", "Kelas", "Jumlah", "Rencana Kembali", "Status"],
        total=["", "JUMLAH", "", "", "", sum(r[5] for r in rows), "", ""],
        ringkas=[("Jumlah transaksi", len(rows)), ("Masih dipinjam (unit)", sum(r[5] for r in rows if r[7] != "Dikembalikan")),
                 ("Terlambat (transaksi)", sum(1 for r in rows if r[7] == "Terlambat"))])

def buat_csv(d):
    out = io.StringIO(); w = csv.writer(out)
    w.writerow(d["head"]); w.writerows(d["rows"])
    if d["total"]: w.writerow(d["total"])
    return ("\ufeff" + out.getvalue()).encode("utf-8")

def buat_xlsx(d):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.properties import PageSetupProperties
    wb = Workbook(); ws = wb.active; ws.title = "Laporan"
    head, n = d["head"], len(d["head"])
    sisi = Side(style="thin", color="9FB4D6"); bd = Border(left=sisi, right=sisi, top=sisi, bottom=sisi)
    tengah = Alignment(horizontal="center", vertical="center", wrap_text=True)
    judul = [(d["judul"].upper(), Font(bold=True, size=14)), (f"{NAMA_SEKOLAH} - {NAMA_KELAS}", Font(size=11)),
             (f"Dicetak: {tgl_id()} oleh {session.get('u', '-')}", Font(size=10, italic=True))]
    for r, (t, f) in enumerate(judul, 1):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n)
        c = ws.cell(r, 1, t); c.font = f; c.alignment = tengah
    hr = 5
    for c, h in enumerate(head, 1):
        x = ws.cell(hr, c, h); x.font = Font(bold=True, color="FFFFFF"); x.fill = PatternFill("solid", fgColor="4F81D9")
        x.alignment = tengah; x.border = bd
    ws.row_dimensions[hr].height = 26
    data = list(d["rows"]) + ([d["total"]] if d["total"] else [])
    for ri, row in enumerate(data, hr + 1):
        tot = bool(d["total"]) and ri == hr + len(data)
        for c, v in enumerate(row, 1):
            x = ws.cell(ri, c, v); x.border = bd
            x.alignment = Alignment(horizontal="center" if isinstance(v, (int, float)) else "left", vertical="center", wrap_text=True)
            if tot: x.font = Font(bold=True); x.fill = PatternFill("solid", fgColor="D9E6F9")
            elif (ri - hr) % 2 == 0: x.fill = PatternFill("solid", fgColor="EEF4FD")
    r = hr + len(data) + 2
    ws.cell(r, 1, "Ringkasan").font = Font(bold=True)
    for lab, val in d["ringkas"]:
        r += 1; ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        ws.cell(r, 1, lab); ws.cell(r, 4, val).alignment = Alignment(horizontal="left")
    r += 2; kn = max(n - 1, 3)
    for off, (a, b) in enumerate([("Mengetahui,", tgl_id()), (TTD_KIRI, TTD_KANAN)]):
        ws.cell(r + off, 2, a).alignment = tengah; ws.cell(r + off, kn, b).alignment = tengah
    ws.cell(r + 5, 2, "( ........................ )").alignment = tengah; ws.cell(r + 5, kn, "( ........................ )").alignment = tengah
    for c in range(1, n + 1):
        L = max([len(str(head[c - 1]))] + [len(str(x[c - 1])) for x in data])
        ws.column_dimensions[get_column_letter(c)].width = min(max(L + 3, 6), 38)
    ws.freeze_panes = ws.cell(hr + 1, 1); ws.print_title_rows = f"{hr}:{hr}"
    ws.page_setup.orientation = "landscape" if n > 6 else "portrait"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

def buat_pdf(d):
    from xml.sax.saxutils import escape
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepTogether
    head, rows, n = d["head"], d["rows"], len(d["head"])
    ps = landscape(A4) if n > 6 else A4
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=ps, leftMargin=1.6 * cm, rightMargin=1.6 * cm, topMargin=1.5 * cm, bottomMargin=1.8 * cm, title=d["judul"])
    ss = getSampleStyleSheet(); ink = colors.HexColor("#2d3e5e")
    sty = lambda nm, **k: ParagraphStyle(nm, parent=ss["Normal"], textColor=ink, **k)
    jud, sub = sty("jud", fontName="Helvetica-Bold", fontSize=15, alignment=1, leading=19, spaceAfter=3), sty("sub", fontSize=10, alignment=1, leading=13)
    kiri, tgh = sty("kiri", fontSize=8.5, leading=11), sty("tgh", fontSize=8.5, leading=11, alignment=1)
    hd = ParagraphStyle("hd", parent=tgh, textColor=colors.white, fontName="Helvetica-Bold")
    bold, boldc = ParagraphStyle("bd", parent=kiri, fontName="Helvetica-Bold"), ParagraphStyle("bc", parent=tgh, fontName="Helvetica-Bold")
    numcol = [bool(rows) and all(isinstance(r[c], (int, float)) for r in rows) for c in range(n)]
    P = lambda v, st: Paragraph(escape(str(v)), st)
    data = [[P(h, hd) for h in head]] + [[P(v, tgh if numcol[c] else kiri) for c, v in enumerate(r)] for r in rows]
    if d["total"]: data.append([P(v, boldc if numcol[c] else bold) for c, v in enumerate(d["total"])])
    lens = [min(max([len(str(head[c]))] + [len(str(r[c])) for r in rows]), 26) + 3 for c in range(n)]
    t = Table(data, colWidths=[doc.width * x / sum(lens) for x in lens], repeatRows=1)
    st = [("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4F81D9")), ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#9FB4D6")),
          ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
          ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#EEF4FD")])]
    if d["total"]: st.append(("BACKGROUND", (0, len(data) - 1), (-1, len(data) - 1), colors.HexColor("#D9E6F9")))
    t.setStyle(TableStyle(st))
    ringk = [Paragraph(f"<b>{escape(a)}</b>: {escape(str(b))}", kiri) for a, b in d["ringkas"]]
    ttd = Table([["Mengetahui,", "", tgl_id()], [TTD_KIRI, "", TTD_KANAN], ["", "", ""], ["( ........................ )", "", "( ........................ )"]],
                colWidths=[doc.width * .35, doc.width * .3, doc.width * .35], rowHeights=[14, 14, 46, 14])
    ttd.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"), ("FONTSIZE", (0, 0), (-1, -1), 9), ("TEXTCOLOR", (0, 0), (-1, -1), ink)]))
    story = [P(d["judul"].upper(), jud), P(f"{NAMA_SEKOLAH} - {NAMA_KELAS}", sub), P(f"Dicetak: {tgl_id()} oleh {session.get('u', '-')}", sub),
             Spacer(1, 12), t, Spacer(1, 10)] + ringk + [Spacer(1, 18), KeepTogether(ttd)]
    def foot(cv, dc):
        cv.setFont("Helvetica", 8); cv.setFillColor(ink)
        cv.drawString(1.6 * cm, 1 * cm, NAMA_APP); cv.drawRightString(ps[0] - 1.6 * cm, 1 * cm, f"Halaman {dc.page}")
    doc.build(story, onFirstPage=foot, onLaterPages=foot)
    return buf.getvalue()

@app.route("/ekspor/<jenis>/<fmt>")
@need(staff=True)
def ekspor(jenis, fmt):
    if jenis not in JENIS or fmt not in ("csv", "xlsx", "pdf"):
        flash("Format ekspor tidak dikenal."); return redirect("/")
    d = dataset(jenis, request.args.get("kelas", ""))
    body, mime = {"csv": (buat_csv, "text/csv"), "xlsx": (buat_xlsx, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                  "pdf": (buat_pdf, "application/pdf")}[fmt]
    log(f"Ekspor {jenis} ({fmt.upper()})")
    return Response(body(d), mimetype=mime, headers={"Content-Disposition": f'attachment; filename="{jenis}_{today().isoformat()}.{fmt}"'})

# ---------------------------------------------------------------- tampilan (HTML)
T = {}
T["base"] = r"""<!doctype html><html lang="id"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin><link href="https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800&family=Poppins:wght@500;600;700&display=swap" rel="stylesheet"><title>{{app_name}}</title><style>
:root{--ink:#1b2547;--mut:#6a7898;--pri:#4f7cff;--pri2:#7c5cff;--line:#dfe7fa;--ok:#12a773;--okb:#d9f7ea;--wr:#b87500;--wrb:#fff0cc;--bd:#d63a4f;--bdb:#ffe2e5;--sh:0 10px 30px rgba(60,90,200,.12)}
*{box-sizing:border-box}[hidden]{display:none!important}html{scroll-behavior:smooth}
body{margin:0;color:var(--ink);font:15px/1.55 'Nunito',system-ui,-apple-system,'Segoe UI',sans-serif;background:linear-gradient(135deg,#d9e8ff 0%,#e6e2ff 55%,#fbe4f1 100%) fixed;min-height:100vh}
body::before,body::after{content:"";position:fixed;z-index:-1;border-radius:50%;filter:blur(70px);opacity:.55;pointer-events:none}
body::before{width:420px;height:420px;background:#8fb4ff;top:-120px;right:-80px}body::after{width:380px;height:380px;background:#ffb8dd;bottom:-140px;left:22%}
h1,h2,h3{font-family:'Poppins','Nunito',system-ui,sans-serif;letter-spacing:-.2px;margin:0}
h1{font-size:24px}h2{font-size:18px;margin:26px 0 10px}.sub{margin:2px 0 0;color:var(--mut)}.mut{color:var(--mut)}
main{flex:1;min-width:0;max-width:1080px;margin:0 auto;padding:22px 26px 44px}
.ph{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:16px}
.card{background:rgba(255,255,255,.82);backdrop-filter:blur(10px);-webkit-backdrop-filter:blur(10px);border:1px solid rgba(255,255,255,.9);border-radius:20px;padding:18px;margin-bottom:16px;box-shadow:var(--sh);animation:up .45s ease both}
@keyframes up{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}
.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:16px}
.stat{display:flex;gap:13px;align-items:center;margin:0;padding:16px;transition:.2s}.stat:hover{transform:translateY(-3px)}
.stat .ic{width:48px;height:48px;border-radius:15px;display:grid;place-items:center;font-size:23px;flex:none}
.stat b{display:block;font:700 27px 'Poppins',sans-serif;line-height:1.1}.stat span{color:var(--mut);font-size:12.5px;line-height:1.3;display:block}
.c1{background:#e1eaff}.c2{background:var(--okb)}.c3{background:var(--wrb)}.c4{background:var(--bdb)}.c5{background:#ece4ff}.c6{background:#ffe3f1}
.hero{background:linear-gradient(120deg,#3f6bff,#7c5cff 60%,#c056ff);color:#fff;border-radius:24px;padding:24px 26px;margin-bottom:16px;box-shadow:0 16px 36px rgba(80,90,230,.35);position:relative;overflow:hidden}
.hero::after{content:"📦";position:absolute;right:22px;bottom:-18px;font-size:110px;opacity:.16}
.hero h1{color:#fff}.hero p{margin:4px 0 16px;opacity:.92}.hero .btn{background:rgba(255,255,255,.22);box-shadow:none}.hero .btn:hover{background:rgba(255,255,255,.34)}
.tw{overflow-x:auto;padding:8px 10px}table{width:100%;border-collapse:collapse;min-width:560px}
th{font-size:11.5px;text-transform:uppercase;letter-spacing:.6px;color:var(--mut);text-align:left;padding:10px 12px;border-bottom:2px solid var(--line);white-space:nowrap}
td{padding:11px 12px;border-bottom:1px solid #ecf1fc;vertical-align:middle}tr:last-child td{border-bottom:0}tbody tr:hover td,table tr:hover td{background:#f5f8ff}
.b{display:inline-flex;align-items:center;gap:5px;padding:3px 10px;border-radius:99px;font-size:12px;font-weight:800;margin:1px}.b::before{content:"";width:6px;height:6px;border-radius:50%;background:currentColor}
.b.baik{background:var(--okb);color:var(--ok)}.b.ringan{background:var(--wrb);color:var(--wr)}.b.berat{background:var(--bdb);color:var(--bd)}.b.n{background:#e8edf9;color:var(--mut)}
.chip{display:inline-block;padding:3px 10px;border-radius:10px;background:#eaf0ff;color:var(--pri);font-size:12.5px;font-weight:700}
input,select,textarea,button{font:inherit;padding:9px 12px;border:1px solid var(--line);border-radius:12px;background:#fff;color:var(--ink);transition:.15s}
input:focus,select:focus,textarea:focus{outline:0;border-color:var(--pri);box-shadow:0 0 0 4px rgba(79,124,255,.16)}
input[type=file]{background:#f5f8ff;border-style:dashed;padding:8px}
button,.btn{background:linear-gradient(135deg,var(--pri),var(--pri2));color:#fff;border:0;cursor:pointer;text-decoration:none;display:inline-flex;align-items:center;gap:6px;padding:9px 16px;border-radius:12px;font-weight:700;box-shadow:0 6px 16px rgba(79,124,255,.28)}
button:hover,.btn:hover{transform:translateY(-1px);filter:brightness(1.06)}
button.g,.btn.g{background:#eaf0ff;color:var(--pri);box-shadow:none}button.d{background:linear-gradient(135deg,#f0645a,#d63a4f);box-shadow:0 6px 16px rgba(214,58,79,.25)}
button.s,.btn.s{padding:5px 11px;font-size:13px;border-radius:10px}
form.i{display:inline-flex;gap:6px;align-items:center}input.sm{width:68px}
.fg{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:14px}.fg label{display:block;font-size:12.5px;font-weight:700;color:var(--mut);margin-bottom:4px}
.fg input,.fg select,.fg textarea{width:100%}
.row{display:flex;gap:9px;flex-wrap:wrap;align-items:center;margin-bottom:12px}
.fl{display:flex;gap:10px;align-items:center;background:#fff7dc;border:1px solid #f4dfa0;padding:11px 15px;border-radius:14px;margin-bottom:14px;animation:up .35s ease both}
.fl.info{background:#e8f0ff;border-color:#c5d6ff}.fl a{color:var(--pri2);font-weight:800}
.b3d{display:inline-flex;padding:8px 15px;border-radius:99px;color:#fff;font:700 13px 'Poppins',sans-serif;text-decoration:none;box-shadow:0 5px 12px rgba(0,0,0,.14);transition:.15s}.b3d:hover{transform:translateY(-2px)}
.b3d.csv{background:linear-gradient(135deg,#46c1c9,#2a93a0)}.b3d.xls{background:linear-gradient(135deg,#4cc17b,#23975a)}.b3d.pdf{background:linear-gradient(135deg,#f58a7e,#d9473a)}
.xg{display:inline-flex;gap:8px;align-items:center;flex-wrap:wrap}.xg b{font:600 13px 'Poppins',sans-serif;color:var(--mut)}
.g2{display:grid;grid-template-columns:repeat(2,1fr);gap:16px}.g2 .card{margin:0}
.dn{display:flex;gap:20px;align-items:center;flex-wrap:wrap;margin-top:10px}.lg{list-style:none;margin:0;padding:0}.lg li{margin:6px 0}
.lg i{display:inline-block;width:12px;height:12px;border-radius:4px;margin-right:8px}
.hb{display:grid;grid-template-columns:110px 1fr 32px;gap:10px;align-items:center;margin:10px 0;font-size:13px}
.tr{background:#e6ecfb;border-radius:99px;height:12px;overflow:hidden}.tr>div{height:100%;display:flex;background:linear-gradient(90deg,var(--pri),var(--pri2));border-radius:99px}.tr i{display:block;height:100%}
.vb{display:flex;gap:10px;align-items:flex-end;padding-top:24px}.vc{flex:1;text-align:center}
.cols{display:flex;gap:6px;align-items:flex-end;justify-content:center;height:140px;border-bottom:2px solid var(--line)}
.col{width:34px;min-height:3px;border-radius:8px 8px 0 0;position:relative}.col span{position:absolute;top:-19px;left:-8px;right:-8px;font-size:12px;font-weight:800}
.lb2{font-size:13px;margin-top:6px;font-weight:700}.key{font-size:12px;color:var(--mut);margin-top:10px}.key i{display:inline-block;width:10px;height:10px;border-radius:3px;margin:0 4px 0 8px}
.rl{display:grid;gap:14px}.rc{display:grid;grid-template-columns:160px 1fr;gap:18px;padding:14px;margin:0}
.pic{width:160px;height:130px;border-radius:15px;object-fit:cover;background:#eef2ff;display:grid;place-items:center;color:var(--mut);font-size:13px;text-align:center;cursor:zoom-in}
div.pic{cursor:default}.rc h3{font-size:17px}.rc .meta{color:var(--mut);font-size:13px;margin:2px 0 8px}.rc .act{margin-top:10px;display:flex;gap:8px;flex-wrap:wrap}
.nt{display:flex;justify-content:space-between;gap:14px;padding:12px 14px;color:var(--ink);text-decoration:none;border-radius:14px;margin-bottom:6px;background:#f6f8ff;transition:.15s}
.nt:hover{background:#e9efff;transform:translateX(3px)}.nt.un{background:#e3ecff;font-weight:800;border-left:5px solid var(--pri)}.nt small{color:var(--mut);white-space:nowrap;font-weight:600}
body.app{display:flex;align-items:flex-start}
.side{width:272px;flex:none;position:sticky;top:0;height:100vh;overflow-y:auto;background:linear-gradient(185deg,#1d2b78 0%,#3050dc 60%,#6a46d8 100%);color:#fff;padding:22px 14px;display:flex;flex-direction:column;gap:4px}
.brand{display:flex;gap:11px;align-items:center;font:700 17px 'Poppins',sans-serif;line-height:1.2;padding:4px 8px 18px}
.brand small{display:block;font:700 10.5px 'Nunito',sans-serif;opacity:.7;letter-spacing:.6px;text-transform:uppercase;margin-top:2px}
.logo{width:44px;height:44px;border-radius:14px;background:rgba(255,255,255,.18);display:grid;place-items:center;font-size:23px;flex:none}
.menu{display:flex;flex-direction:column;gap:3px}
.side a{display:flex;align-items:center;gap:11px;color:rgba(255,255,255,.86);text-decoration:none;padding:10px 12px;border-radius:14px;font:500 14px 'Poppins',sans-serif;transition:.15s}
.side a i{font-style:normal;width:26px;text-align:center;font-size:17px}.side a:hover{background:rgba(255,255,255,.13);color:#fff}
.side a.on{background:#fff;color:#2f4fd6;font-weight:600;box-shadow:0 8px 20px rgba(0,0,0,.2)}
.side .cnt{margin-left:auto;background:#ff5a6a;color:#fff;border-radius:99px;font:800 11px 'Nunito',sans-serif;padding:1px 8px}
.me{margin-top:auto;display:flex;gap:10px;align-items:center;padding:14px 8px 8px;border-top:1px solid rgba(255,255,255,.2)}
.av{width:40px;height:40px;border-radius:50%;background:linear-gradient(135deg,#ffd36e,#ff8fb1);color:#4a2a00;display:grid;place-items:center;font-weight:800;font-size:17px;flex:none}
.me small{display:block;opacity:.72;font-size:12px}.mel{display:flex;gap:6px;padding:0 2px}.side .mel a{font-size:12.5px;padding:7px 10px;background:rgba(255,255,255,.13);flex:1;justify-content:center}
.topbar{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;color:var(--mut);font-weight:700}
.bell{position:relative;display:inline-flex;width:44px;height:44px;align-items:center;justify-content:center;background:#fff;border-radius:50%;text-decoration:none;font-size:19px;box-shadow:var(--sh);transition:.15s}.bell:hover,.bell.on{transform:scale(1.07)}
.bdg{position:absolute;top:-4px;right:-4px;background:#ff4d5e;color:#fff;font:800 11px 'Nunito',sans-serif;border-radius:99px;padding:1px 6px;min-width:20px;text-align:center}.bdg[hidden]{display:none}
#toasts{position:fixed;top:16px;right:16px;display:flex;flex-direction:column;gap:10px;z-index:50;max-width:340px}
.toast{display:block;background:#fff;color:var(--ink);text-decoration:none;border-radius:16px;padding:12px 15px;box-shadow:0 14px 34px rgba(40,60,150,.3);border-left:5px solid var(--pri);animation:up .35s ease both;transition:.4s}
.toast.out{opacity:0;transform:translateX(30px)}.toast b{display:block;font-size:12.5px;color:var(--pri)}.toast span{font-size:14px}
#lb{position:fixed;inset:0;background:rgba(15,23,60,.8);display:none;place-items:center;z-index:60;cursor:zoom-out}#lb.on{display:grid}
#lb img{max-width:92vw;max-height:88vh;border-radius:16px;box-shadow:0 20px 60px rgba(0,0,0,.5)}
.auth{max-width:410px;margin:7vh auto 0;text-align:center}.auth .logo{width:70px;height:70px;font-size:36px;margin:0 auto 12px;background:linear-gradient(135deg,var(--pri),var(--pri2));border-radius:22px;box-shadow:0 12px 26px rgba(79,124,255,.4)}
.auth form{text-align:left;margin-top:18px}.auth .fg{grid-template-columns:1fr}.auth button.full{width:100%;justify-content:center;padding:12px}.demo{display:flex;gap:8px;justify-content:center;flex-wrap:wrap;margin-top:14px}
@media(max-width:820px){.grid{grid-template-columns:repeat(2,1fr)}.g2{grid-template-columns:1fr}body.app{display:block}.side{width:auto;height:auto;position:static;padding:12px;gap:8px}.brand{padding:0 4px 4px}
.menu{flex-direction:row;overflow-x:auto;gap:6px}.side a{white-space:nowrap;padding:8px 12px}.me{margin-top:4px;padding-top:10px}main{padding:16px 14px 36px}}
@media(max-width:560px){.rc{grid-template-columns:1fr}.pic{width:100%;height:190px}}
/* ---- unggah foto ---- */
.up{position:relative}
.up input[type=file]{position:absolute;width:1px;height:1px;padding:0;border:0;opacity:0;pointer-events:none}
.fg label.up-box{display:grid;justify-items:center;gap:7px;text-align:center;padding:24px 16px;margin:0;font-weight:400;font-size:inherit;color:var(--mut);cursor:pointer;
border:2px dashed #b6c7f6;border-radius:20px;background:linear-gradient(180deg,#f8faff,#edf2ff);transition:.2s}
.fg label.up-box:hover,.up.drag label.up-box{border-color:var(--pri);background:#e8efff;transform:translateY(-1px)}
.up input:focus-visible+label.up-box{outline:3px solid rgba(79,124,255,.35);outline-offset:2px}
.up-ic{width:56px;height:56px;border-radius:50%;display:grid;place-items:center;color:#fff;background:linear-gradient(135deg,var(--pri),var(--pri2));box-shadow:0 10px 22px rgba(79,124,255,.35)}
.up-t{font:600 15px 'Poppins','Nunito',sans-serif;color:var(--ink)}.up-s{font-size:12.5px;line-height:1.4}
.up-pv{display:flex;align-items:center;gap:14px;padding:10px 12px;border:1px solid var(--line);border-radius:18px;background:#fff;box-shadow:var(--sh)}
.up-pv img{width:104px;height:78px;object-fit:cover;border-radius:12px;flex:none}
.up-i{flex:1;min-width:0}.up-i b{display:block;font:600 14px 'Poppins','Nunito',sans-serif;color:var(--ink);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.up-i small{color:var(--mut);font-weight:600}
.up .up-x{width:36px;height:36px;padding:0;justify-content:center;border-radius:50%;background:#ffe9ec;color:var(--bd);box-shadow:none}
.up .up-x:hover{background:#ffd7dd}.up-err{color:var(--bd);font-size:12.5px;font-weight:700;margin-top:6px;min-height:0}
.pic{align-content:center;gap:6px}.pic svg{opacity:.5}.pic span{font-weight:700;font-size:12.5px}
#lb .lbx{position:absolute;top:18px;right:18px;width:44px;height:44px;padding:0;justify-content:center;border-radius:50%;background:rgba(255,255,255,.2);box-shadow:none}
#lb .lbx:hover{background:rgba(255,255,255,.34)}
</style></head><body class="{{'app' if session.u}}">
{% if session.u %}{% set P=request.path %}<aside class="side"><div class="brand"><span class="logo">📦</span><div>{{app_name}}<small>Kelola barang kelas</small></div></div>
<nav class="menu"><a href="/" class="{{'on' if P=='/'}}"><i>🏠</i>Dashboard</a><a href="/barang" class="{{'on' if P.startswith('/barang')}}"><i>📚</i>Barang</a>
<a href="/pinjam" class="{{'on' if P.startswith('/pinjam')}}"><i>🤝</i>Peminjaman</a>
<a href="/laporan" class="{{'on' if P.startswith('/laporan')}}"><i>⚠️</i>Laporan Kerusakan{% if lapn %}<span class="cnt">{{lapn}}</span>{% endif %}</a>
<a href="/tindak" class="{{'on' if P.startswith('/tindak')}}"><i>🛠</i>Perbaikan &amp; Pengadaan</a>
{% if staff %}<a href="/riwayat" class="{{'on' if P.startswith('/riwayat')}}"><i>🕘</i>Riwayat</a>{% endif %}
{% if is_admin %}<a href="/pengguna" class="{{'on' if P.startswith('/pengguna')}}"><i>👥</i>Kelola Pengguna</a>{% endif %}</nav>
<div class="me"><div class="av">{{session.u[0]|upper}}</div><div><b>{{session.u}}</b><small>{{roles[session.role]}}</small></div></div>
<div class="mel"><a href="/akun">🔑 Password</a><a href="/logout">⎋ Keluar</a></div></aside>{% endif %}
<main>{% if session.u %}<div class="topbar"><span>📅 {{hari_ini}}</span><a class="bell {{'on' if request.path.startswith('/notif')}}" href="/notif" title="Notifikasi">🔔<span class="bdg" id="bdg" {{'hidden' if not unread}}>{{unread}}</span></a></div>{% endif %}
{% for m in get_flashed_messages() %}<div class="fl">💬 <span>{{m}}</span></div>{% endfor %}{% block c %}{% endblock %}</main>
{% if session.u %}<div id="toasts"></div><div id="lb"><button type="button" class="lbx" aria-label="Tutup"><svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M6 6l12 12M18 6L6 18"/></svg></button><img alt="Foto"></div><script>
(function(){var last={{last_notif}},box=document.getElementById('toasts'),bdg=document.getElementById('bdg');
function toast(n){var a=document.createElement('a');a.className='toast';a.href=n.link;var t=document.createElement('b');t.textContent='🔔 Notifikasi baru';var s=document.createElement('span');s.textContent=n.pesan;a.appendChild(t);a.appendChild(s);box.appendChild(a);setTimeout(function(){a.classList.add('out');setTimeout(function(){a.remove()},450)},9000)}
function cek(){fetch('/notif/cek?sejak='+last,{credentials:'same-origin'}).then(function(r){return r.json()}).then(function(d){if(bdg){bdg.textContent=d.unread;bdg.hidden=!d.unread}d.baru.forEach(function(n){toast(n);if(n.id>last)last=n.id})}).catch(function(){})}
setInterval(cek,15000);
document.addEventListener('click',function(e){var z=e.target.closest('[data-zoom]');if(z){e.preventDefault();var lb=document.getElementById('lb');lb.querySelector('img').src=z.getAttribute('href');lb.classList.add('on')}});
document.getElementById('lb').addEventListener('click',function(){this.classList.remove('on')});})();
</script>{% endif %}</body></html>"""

T["login"] = r"""{% extends "base" %}{% block c %}<div class="card auth"><div class="logo">📦</div>
<h1>{{app_name}}</h1><p class="sub">Masuk untuk mengelola barang kelas dengan rapi.</p>
<form method="post" class="fg">{{ csrf_input() }}
<div><label>Username</label><input id="u" name="username" required autofocus placeholder="mis. admin"></div>
<div><label>Password</label><input id="p" name="password" type="password" required placeholder="••••••••"></div>
<button class="full">Masuk →</button></form>
<p class="sub" style="margin-top:16px;font-size:13px">Akun contoh (klik untuk mengisi otomatis)</p>
<div class="demo"><button type="button" class="g s" onclick="document.getElementById('u').value='admin';document.getElementById('p').value='admin123'">👑 Admin</button>
<button type="button" class="g s" onclick="document.getElementById('u').value='guru';document.getElementById('p').value='guru123'">🎓 Guru</button>
<button type="button" class="g s" onclick="document.getElementById('u').value='siswa';document.getElementById('p').value='siswa123'">🎒 Siswa</button></div></div>{% endblock %}"""

T["dash"] = r"""{% extends "base" %}{% block c %}
<div class="hero"><h1>{{sapa}}, {{session.u}}! 👋</h1>
<p>{{s.pinjam}} unit sedang dipinjam{% if staff %} · {{s.lap}} laporan kerusakan menunggu penilaian{% endif %}.</p>
<div class="row" style="margin:0"><a class="btn" href="/pinjam">🤝 Pinjam barang</a><a class="btn" href="/laporan">📷 Lapor kerusakan</a>{% if staff %}<a class="btn" href="/tindak">🛠 Tindak lanjut</a>{% endif %}</div></div>
{% if staff and late %}<div class="fl">⏰ <span>Ada <b>{{late}}</b> peminjaman yang terlambat dikembalikan. <a href="/pinjam">Lihat peminjaman</a></span></div>{% endif %}
{% if staff and s.lap %}<div class="fl info">⚠️ <span><b>{{s.lap}}</b> laporan kerusakan menunggu penilaian. <a href="/laporan">Nilai sekarang</a></span></div>{% endif %}
<div class="grid">
<div class="card stat"><div class="ic c1">📦</div><div><b>{{s.total}}</b><span>Total barang</span></div></div>
<div class="card stat"><div class="ic c2">✅</div><div><b style="color:var(--ok)">{{s.baik}}</b><span>Baik (layak dipakai)</span></div></div>
<div class="card stat"><div class="ic c3">🔧</div><div><b style="color:var(--wr)">{{s.ringan}}</b><span>Rusak ringan</span></div></div>
<div class="card stat"><div class="ic c4">🛑</div><div><b style="color:var(--bd)">{{s.berat}}</b><span>Rusak berat</span></div></div>
<div class="card stat"><div class="ic c5">🤝</div><div><b>{{s.pinjam}}</b><span>Sedang dipinjam</span></div></div>
<div class="card stat"><div class="ic c6">⚠️</div><div><b>{{s.lap}}</b><span>Laporan menunggu</span></div></div></div>
<h2>📊 Grafik</h2><div class="g2">
<div class="card"><b>Komposisi kondisi barang</b>{% if s.total %}<div class="dn">
<svg viewBox="0 0 36 36" width="150" height="150"><circle cx="18" cy="18" r="15.9155" fill="none" stroke="#e6ecfb" stroke-width="5"/>
{% for g in segs %}{% if g.p>0 %}<circle cx="18" cy="18" r="15.9155" fill="none" stroke="{{g.col}}" stroke-width="5" stroke-dasharray="{{g.p}} {{100-g.p}}" stroke-dashoffset="{{g.off}}"/>{% endif %}{% endfor %}
<text x="18" y="19.6" text-anchor="middle" font-size="6" font-weight="700" fill="#1b2547">{{s.total}}</text></svg>
<ul class="lg">{% for g in segs %}<li><i style="background:{{g.col}}"></i>{{g.lab}}: <b>{{g.n}}</b> ({{'%.0f'|format(g.p)}}%)</li>{% endfor %}</ul></div>
{% else %}<p class="mut">Belum ada data.</p>{% endif %}</div>
<div class="card"><b>Barang per kategori</b>{% for k in kat %}<div class="hb"><span>{{k.nama}}</span>
<div class="tr" style="height:14px"><div style="width:{{k.w}}%;background:none"><i style="width:{{k.bp}}%;background:var(--ok)"></i><i style="width:{{k.rp}}%;background:#e0a526"></i><i style="width:{{k.wp}}%;background:var(--bd)"></i></div></div><b>{{k.tot}}</b></div>
{% else %}<p class="mut">Belum ada data.</p>{% endfor %}
<div class="key"><i style="background:var(--ok)"></i>Baik<i style="background:#e0a526"></i>Rusak ringan<i style="background:var(--bd)"></i>Rusak berat</div></div>
<div class="card"><b>Peminjaman per kelas</b><div class="vb">{% for k in kc %}<div class="vc"><div class="cols">
<div class="col" style="height:{{k.h1}}%;background:linear-gradient(180deg,var(--pri),var(--pri2))"><span>{{k.total}}</span></div>
<div class="col" style="height:{{k.h2}}%;background:linear-gradient(180deg,#ffc34d,#f0a020)"><span>{{k.aktif}}</span></div></div><div class="lb2">Kelas {{k.kelas}}</div></div>{% endfor %}</div>
<div class="key"><i style="background:var(--pri)"></i>Jumlah peminjaman<i style="background:#f0a020"></i>Unit sedang dipinjam</div></div>
<div class="card"><b>Barang paling sering dipinjam</b>{% for t in top %}<div class="hb"><span>{{t.nama}}</span>
<div class="tr"><div style="width:{{t.w}}%"></div></div><b>{{t.n}}</b></div>{% else %}<p class="mut">Belum ada peminjaman.</p>{% endfor %}</div></div>
<h2>👀 Perlu perhatian</h2><div class="card">
{% if perhatian %}<div class="tw"><table><tr><th>Barang</th><th>Rusak ringan</th><th>Rusak berat</th></tr>
{% for i in perhatian %}<tr><td><b>{{i.nama}}</b></td><td><span class="b ringan">{{i.ringan}}</span></td><td><span class="b berat">{{i.berat}}</span></td></tr>{% endfor %}</table></div>
{% else %}<p>Semua barang dalam kondisi baik 🎉</p>{% endif %}</div>{% endblock %}"""

T["barang"] = r"""{% extends "base" %}{% from "macros" import exp %}{% block c %}
<div class="ph"><div><h1>📚 Data Barang</h1><p class="sub">Pantau jumlah, kondisi, dan ketersediaan barang kelas.</p></div>
<div class="row" style="margin:0">{% if is_admin %}<a class="btn" href="/barang/tambah">＋ Tambah barang</a>{% endif %}{% if staff %}{{ exp('inventaris') }}{% endif %}</div></div>
<form class="card row"><input name="q" placeholder="🔍 Cari nama / lokasi" value="{{qq}}" style="flex:1;min-width:180px">
<select name="kat"><option value="">Semua kategori</option>{% for c in cats %}<option {{'selected' if c==kat}}>{{c}}</option>{% endfor %}</select>
<select name="kon"><option value="">Semua kondisi</option>{% for k,v in kond.items() %}<option value="{{k}}" {{'selected' if k==kon}}>{{v}}</option>{% endfor %}</select>
<button>Terapkan</button></form>
<div class="card tw"><table><tr><th>Barang</th><th>Kategori</th><th>Lokasi</th><th>Kondisi (unit)</th><th>Tersedia</th><th></th></tr>
{% for i in items %}<tr><td><b>{{i.nama}}</b>{% if i.catatan %}<br><small class="mut">{{i.catatan}}</small>{% endif %}</td><td><span class="chip">{{i.kategori}}</span></td><td>{{i.lokasi}}</td>
<td><span class="b baik">Baik {{i.baik}}</span><span class="b ringan">Ringan {{i.ringan}}</span><span class="b berat">Berat {{i.berat}}</span></td>
<td>{% if i.tersedia>0 %}<span class="b baik">{{i.tersedia}} tersedia</span>{% else %}<span class="b berat">Habis</span><br>
{% if i.id in mine %}<small class="mut">🔔 kamu akan dikabari</small>{% else %}<form class="i" method="post" action="/minat/{{i.id}}">{{ csrf_input() }}<button class="g s">🔔 Kabari saya</button></form>{% endif %}{% endif %}</td>
<td>{% if is_admin %}<a class="btn g s" href="/barang/{{i.id}}/ubah">Ubah</a>
<form class="i" method="post" action="/barang/{{i.id}}/hapus" onsubmit='return confirm({{ ("Hapus " ~ i.nama ~ "? Riwayat peminjamannya tetap tersimpan.")|tojson }})'>{{ csrf_input() }}<button class="d s">Hapus</button></form>{% endif %}</td></tr>
{% else %}<tr><td colspan="6" class="mut">Tidak ada barang.</td></tr>{% endfor %}</table></div>{% endblock %}"""

T["form"] = r"""{% extends "base" %}{% block c %}<div class="ph"><div><h1>{{'✏️ Ubah' if it else '＋ Tambah'}} Barang</h1><p class="sub">Isi jumlah per kondisi dengan teliti.</p></div></div>
<form method="post" class="card">{{ csrf_input() }}<div class="fg">
<div><label>Nama barang</label><input name="nama" required value="{{it.nama if it}}"></div>
<div><label>Kategori</label><select name="kategori">{% for c in cats %}<option {{'selected' if it and it.kategori==c}}>{{c}}</option>{% endfor %}</select></div>
<div><label>Lokasi di kelas</label><input name="lokasi" value="{{it.lokasi if it}}"></div>
<div><label>✅ Jumlah Baik (layak dipakai)</label><input name="baik" type="number" min="0" value="{{it.baik if it else 0}}"></div>
<div><label>🔧 Jumlah Rusak Ringan</label><input name="ringan" type="number" min="0" value="{{it.ringan if it else 0}}"></div>
<div><label>🛑 Jumlah Rusak Berat</label><input name="berat" type="number" min="0" value="{{it.berat if it else 0}}"></div>
<div><label>Tanggal masuk</label><input name="tgl" type="date" value="{{it.tgl if it else today}}"></div>
<div><label>Catatan</label><input name="catatan" value="{{it.catatan if it}}"></div></div>
<p class="row" style="margin:16px 0 0"><button>💾 Simpan</button> <a class="btn g" href="/barang">Batal</a></p></form>{% endblock %}"""

T["tindak"] = r"""{% extends "base" %}{% from "macros" import exp %}{% block c %}
<div class="ph"><div><h1>🛠 Perbaikan &amp; Pengadaan</h1><p class="sub">Siswa pelapor otomatis dikabari saat barang selesai ditangani.</p></div>
{% if staff %}<div class="row" style="margin:0">{{ exp('perbaikan', 'Perbaikan') }}{{ exp('pengadaan', 'Pengadaan') }}</div>{% endif %}</div>
{% for judul, data, key, jenis, tombol in [('🛠 Perlu Perbaikan (Rusak Ringan)', rep, 'ringan', 'perbaikan', '✔ Selesai diperbaiki'),
('🛒 Perlu Pengadaan (Rusak Berat)', adq, 'berat', 'pengadaan', '✔ Sudah diadakan')] %}
<div class="card"><h2 style="margin-top:0">{{judul}}</h2>{% if data %}<div class="tw"><table><tr><th>Barang</th><th>Lokasi</th><th>Jumlah</th>{% if staff %}<th>Tindak lanjut</th>{% endif %}</tr>
{% for i in data %}<tr><td><b>{{i.nama}}</b></td><td>{{i.lokasi}}</td><td><span class="b {{'ringan' if key=='ringan' else 'berat'}}">{{i[key]}} unit</span></td>{% if staff %}<td>
<form class="i" method="post" action="/tindak/{{i.id}}/{{jenis}}">{{ csrf_input() }}<input class="sm" name="n" type="number" min="1" max="{{i[key]}}" value="{{i[key]}}"><button class="s">{{tombol}}</button></form></td>{% endif %}</tr>{% endfor %}
</table></div>{% else %}<p>Tidak ada. 🎉</p>{% endif %}</div>{% endfor %}{% endblock %}"""

T["pinjam"] = r"""{% extends "base" %}{% from "macros" import exp %}{% block c %}
<div class="ph"><div><h1>🤝 Peminjaman &amp; Pengembalian</h1><p class="sub">Stok kurang? Kamu otomatis dikabari saat barangnya tersedia.</p></div></div>
<form method="post" class="card">{{ csrf_input() }}<div class="fg">
<div><label>Barang (hanya kondisi Baik)</label><select name="item_id">{% for i in items %}<option value="{{i.id}}">{{i.nama}} (tersedia {{i.tersedia}})</option>{% endfor %}</select></div>
<div><label>Nama peminjam</label><input name="peminjam" value="{{session.u}}"></div>
<div><label>Kelas peminjam</label><select name="kelas" required><option value="">Pilih kelas</option>{% for k in kelas %}<option value="{{k}}">Kelas {{k}}</option>{% endfor %}</select></div>
<div><label>Rombel (opsional, mis. X-1)</label><input name="rombel" maxlength="20"></div>
<div><label>Jumlah</label><input name="jumlah" type="number" min="1" value="1"></div>
<div><label>Rencana kembali</label><input name="tgl_kembali" type="date" min="{{today}}" value="{{today}}" required></div></div>
<p style="margin:16px 0 0"><button>✔ Catat peminjaman</button></p></form>
<form class="row"><select name="kelas"><option value="">Semua kelas</option>{% for k in kelas %}<option value="{{k}}" {{'selected' if k==kf}}>Kelas {{k}}</option>{% endfor %}</select>
<button class="g">Filter</button>{% if staff %}{{ exp('peminjaman', 'Ekspor', '?kelas=' ~ kf) }}{% endif %}</form>
<div class="card tw"><table><tr><th>Barang</th><th>Peminjam</th><th>Kelas</th><th>Jml</th><th>Pinjam</th><th>Rencana kembali</th><th>Status</th><th></th></tr>
{% for l in loans %}<tr><td><b>{{l.nama}}</b></td><td>{{l.peminjam}}</td><td><span class="chip">{{l.kelas or '-'}} {{l.rombel or ''}}</span></td><td>{{l.jumlah}}</td><td>{{l.tgl_pinjam}}</td><td>{{l.tgl_kembali}}</td>
<td>{% if l.status=='Dipinjam' and l.tgl_kembali<today %}<span class="b berat">Terlambat</span>{% elif l.status=='Dipinjam' %}<span class="b ringan">Dipinjam</span>{% else %}<span class="b baik">Dikembalikan</span>{% endif %}</td>
<td>{% if staff and l.status=='Dipinjam' %}<form class="i" method="post" action="/pinjam/{{l.id}}/kembali">{{ csrf_input() }}<button class="s">Tandai kembali</button></form>{% endif %}</td></tr>
{% else %}<tr><td colspan="8" class="mut">Belum ada peminjaman.</td></tr>{% endfor %}</table></div>{% endblock %}"""

T["laporan"] = r"""{% extends "base" %}{% block c %}
<div class="ph"><div><h1>⚠️ Laporan Kerusakan</h1><p class="sub">Sertakan foto agar admin dan guru bisa menilai tingkat kerusakannya.</p></div></div>
<form method="post" enctype="multipart/form-data" class="card">{{ csrf_input() }}<div class="fg"><div><label>Barang</label><select name="item_id">{% for i in items %}<option value="{{i.id}}">{{i.nama}}</option>{% endfor %}</select></div>
<div class="up" id="up" style="grid-column:1/-1"><label for="foto">Foto kerusakan <span class="chip">Disarankan</span></label>
<input type="file" id="foto" name="foto" accept="image/*">
<label for="foto" class="up-box"><span class="up-ic"><svg viewBox="0 0 24 24" width="26" height="26" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="M12 16V5m0 0l-4 4m4-4l4 4"/><path d="M4 16v2a2 2 0 002 2h12a2 2 0 002-2v-2"/></svg></span><span class="up-t">Klik untuk memilih foto</span><span class="up-s">atau seret dan lepas di sini &middot; JPG, PNG, atau WEBP &middot; maks. 8 MB</span></label>
<div class="up-pv" hidden><img alt="Pratinjau foto"><div class="up-i"><b id="pvn"></b><small id="pvs"></small></div><button type="button" class="up-x" aria-label="Hapus foto" title="Hapus foto"><svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M6 6l12 12M18 6L6 18"/></svg></button></div>
<div class="up-err" role="alert"></div></div>
<div style="grid-column:1/-1"><label>Deskripsi kerusakan</label><textarea name="deskripsi" rows="3" required placeholder="Contoh: kaki kursi patah, tidak bisa diduduki"></textarea></div></div>
<p style="margin:16px 0 0"><button>📨 Kirim laporan</button></p></form>
<div class="rl">{% for r in reps %}{% set st=r.status %}<div class="card rc">
{% if r.foto %}<a href="/foto/{{r.foto}}" data-zoom><img class="pic" src="/foto/{{r.foto}}" alt="Foto kerusakan {{r.nama}}" loading="lazy"></a>{% else %}<div class="pic"><svg viewBox="0 0 24 24" width="34" height="34" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="4"/><circle cx="8.5" cy="8.5" r="1.6"/><path d="M21 15l-5-5L5 21"/></svg><span>Tanpa foto</span></div>{% endif %}
<div><h3>{{r.nama}} <span class="b {{'ringan' if st=='Menunggu' or st.endswith('Ringan') else 'berat' if st.endswith('Berat') else 'baik' if st.startswith('Selesai') else 'n'}}">{{st}}</span></h3>
<div class="meta">👤 {{r.pelapor}} · 📅 {{r.tgl}}</div><div>{{r.deskripsi}}</div>
{% if staff and st=='Menunggu' %}<div class="act">
<form class="i" method="post" action="/laporan/{{r.id}}/ringan">{{ csrf_input() }}<button class="s">🔧 Rusak ringan</button></form>
<form class="i" method="post" action="/laporan/{{r.id}}/berat">{{ csrf_input() }}<button class="d s">🛑 Rusak berat</button></form>
<form class="i" method="post" action="/laporan/{{r.id}}/tolak">{{ csrf_input() }}<button class="g s">Tolak</button></form></div>{% endif %}</div></div>
{% else %}<div class="card mut">Belum ada laporan. 🎉</div>{% endfor %}</div>
<script>
(function(){var up=document.getElementById('up');if(!up)return;
var inp=up.querySelector('input[type=file]'),box=up.querySelector('.up-box'),pv=up.querySelector('.up-pv'),img=pv.querySelector('img'),
nm=document.getElementById('pvn'),sz=document.getElementById('pvs'),er=up.querySelector('.up-err'),url=null;
function fmt(b){return b>=1048576?(b/1048576).toFixed(1)+' MB':Math.max(1,Math.round(b/1024))+' KB'}
function reset(){inp.value='';pv.hidden=true;box.hidden=false;if(url){URL.revokeObjectURL(url);url=null}}
function tampil(f){er.textContent='';if(!f){reset();return}
  if(!/^image\//.test(f.type)){reset();er.textContent='File harus berupa gambar (JPG, PNG, atau WEBP).';return}
  if(f.size>8*1024*1024){reset();er.textContent='Ukuran foto maksimal 8 MB.';return}
  if(url)URL.revokeObjectURL(url);url=URL.createObjectURL(f);img.src=url;nm.textContent=f.name;sz.textContent=fmt(f.size);pv.hidden=false;box.hidden=true}
inp.addEventListener('change',function(){tampil(inp.files[0])});
up.querySelector('.up-x').addEventListener('click',reset);
['dragenter','dragover'].forEach(function(ev){up.addEventListener(ev,function(e){e.preventDefault();up.classList.add('drag')})});
['dragleave','drop'].forEach(function(ev){up.addEventListener(ev,function(e){e.preventDefault();up.classList.remove('drag')})});
up.addEventListener('drop',function(e){var f=e.dataTransfer.files[0];if(!f)return;
  try{var dt=new DataTransfer();dt.items.add(f);inp.files=dt.files;tampil(f)}
  catch(x){reset();er.textContent='Browser ini belum mendukung seret-lepas. Klik kotak untuk memilih foto.'}});})();
</script>{% endblock %}"""

T["riwayat"] = r"""{% extends "base" %}{% block c %}<div class="ph"><div><h1>🕘 Riwayat Perubahan</h1><p class="sub">200 aktivitas terakhir (waktu WIB).</p></div></div><div class="card tw"><table>
<tr><th>Waktu</th><th>Pengguna</th><th>Aktivitas</th></tr>{% for l in logs %}<tr><td class="mut">{{l.waktu}}</td><td><span class="chip">{{l.user}}</span></td><td>{{l.aksi}}</td></tr>
{% else %}<tr><td colspan="3" class="mut">Belum ada riwayat.</td></tr>{% endfor %}</table></div>{% endblock %}"""

T["macros"] = r"""{% macro exp(jenis, label='Ekspor', extra='') %}<span class="xg"><b>⬇ {{label}}</b><a class="b3d csv" href="/ekspor/{{jenis}}/csv{{extra}}">CSV</a><a class="b3d xls" href="/ekspor/{{jenis}}/xlsx{{extra}}">Excel</a><a class="b3d pdf" href="/ekspor/{{jenis}}/pdf{{extra}}">PDF</a></span>{% endmacro %}"""

T["notif"] = r"""{% extends "base" %}{% block c %}<div class="ph"><div><h1>🔔 Notifikasi</h1><p class="sub">Klik notifikasi untuk membuka halaman terkait.</p></div>
<form method="post" action="/notif/baca-semua" style="margin:0">{{ csrf_input() }}<button class="g">✔ Tandai semua dibaca</button></form></div>
<div class="card">{% for n in rows %}<a class="nt {{'un' if not n.dibaca}}" href="/notif/{{n.id}}"><span>{{n.pesan}}</span><small>{{n.waktu}}</small></a>
{% else %}<p class="mut">Belum ada notifikasi.</p>{% endfor %}</div>{% endblock %}"""

T["pengguna"] = r"""{% extends "base" %}{% block c %}<div class="ph"><div><h1>👥 Kelola Pengguna</h1><p class="sub">Tambah akun admin, guru, atau siswa.</p></div></div>
<form method="post" class="card">{{ csrf_input() }}<div class="fg"><div><label>Username</label><input name="username" required></div>
<div><label>Password (min. 6 karakter)</label><input name="password" type="password" minlength="6" required></div>
<div><label>Peran</label><select name="role">{% for k,v in roles.items() %}<option value="{{k}}" {{'selected' if k=='anggota'}}>{{v}}</option>{% endfor %}</select></div></div>
<p style="margin:16px 0 0"><button>＋ Tambah akun</button></p></form>
<div class="card tw"><table><tr><th>Username</th><th>Peran</th><th>Reset password</th><th></th></tr>
{% for u in users %}<tr><td><b>{{u.username}}</b></td><td><span class="chip">{{roles[u.role]}}</span></td>
<td><form class="i" method="post" action="/pengguna/{{u.id}}/password">{{ csrf_input() }}<input name="password" type="password" placeholder="Password baru" minlength="6" required><button class="s">Simpan</button></form></td>
<td>{% if u.username!=session.u %}<form class="i" method="post" action="/pengguna/{{u.id}}/hapus" onsubmit='return confirm({{ ("Hapus akun " ~ u.username ~ "?")|tojson }})'>{{ csrf_input() }}<button class="d s">Hapus</button></form>{% endif %}</td></tr>{% endfor %}
</table></div>{% endblock %}"""

T["akun"] = r"""{% extends "base" %}{% block c %}<div class="ph"><div><h1>🔑 Ganti Password</h1><p class="sub">Gunakan password yang sulit ditebak.</p></div></div>
<form method="post" class="card" style="max-width:440px">{{ csrf_input() }}<div class="fg" style="grid-template-columns:1fr">
<div><label>Password lama</label><input name="lama" type="password" required></div><div><label>Password baru (min. 6 karakter)</label><input name="baru" type="password" minlength="6" required></div>
<div><label>Ulangi password baru</label><input name="ulang" type="password" required></div></div><p style="margin:16px 0 0"><button>💾 Simpan</button></p></form>{% endblock %}"""

app.jinja_loader = DictLoader(T)
app.jinja_env.autoescape = True        # WAJIB: nama template tanpa ekstensi .html tidak otomatis di-escape oleh Flask
app.jinja_env.globals.update(cats=CATS, kond=KOND, kelas=KELAS, app_name=NAMA_APP, csrf_input=csrf_input)

# ---------------------------------------------------------------- jalankan di Colab
init_db()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port)
