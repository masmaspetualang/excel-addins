# Deploy ke Vercel + Update Manifest

Panduan ini untuk deploy add-in **tanpa domain kustom** dulu (pakai URL bawaan Vercel: `https://nama-project.vercel.app`).

## Prasyarat

1. Akun [Vercel](https://vercel.com) (gratis)
2. Akun [GitHub](https://github.com) (disarankan — deploy dari repo)
3. File `.env` lokal sudah berisi:
   - `SUPABASE_URL`
   - `SUPABASE_ANON_KEY`

---

## Opsi A — Deploy lewat website Vercel (paling mudah)

### 1. Push project ke GitHub

```powershell
git init
git add .
git commit -m "Prepare for Vercel deploy"
git branch -M main
git remote add origin https://github.com/USERNAME/excel-quiz-addin.git
git push -u origin main
```

Pastikan **tidak** ter-commit: `.env`, `public/js/config/app.config.js`

### 2. Import di Vercel

1. Login [vercel.com](https://vercel.com) → **Add New** → **Project**
2. Import repo GitHub Anda
3. **Framework Preset:** Other
4. **Root Directory:** biarkan `.` (root project)
5. Vercel akan membaca `vercel.json` (static build + API function, **bukan** server Express)

> Jika build gagal "No entrypoint found": pastikan `package.json` memiliki script `"vercel-build"` dan `vercel.json` memakai `@vercel/static-build`.

### 3. Environment Variables (wajib — isi SEBELUM redeploy)

Di halaman project → **Settings** → **Environment Variables**, tambahkan:

| Name | Value |
|------|--------|
| `SUPABASE_URL` | URL project Supabase Anda (contoh: `https://xxxxx.supabase.co`) |
| `SUPABASE_ANON_KEY` | Anon/public key dari Supabase → Settings → API |

**Jangan** tambahkan `SUPABASE_SERVICE_KEY` di Vercel.

Centang **Production**, **Preview**, dan **Development** → klik **Save**.

> Supabase Dashboard → Project Settings → **API** → copy **Project URL** dan **anon public** key.

Setelah menambah env vars, klik **Redeploy** (deploy ulang) agar build memakai nilai baru.

### 4. Deploy

Klik **Deploy**. Tunggu sampai status **Ready**.

Catat URL production, contoh:
`https://excel-quiz-addin.vercel.app`

### 5. Tes di browser

Buka (ganti dengan URL Anda):

- `https://NAMA-PROJECT.vercel.app/app`
- `https://NAMA-PROJECT.vercel.app/login`
- `https://NAMA-PROJECT.vercel.app/admin`
- `https://NAMA-PROJECT.vercel.app/admin/login`
- `https://NAMA-PROJECT.vercel.app/js/config/app.config.js` (harus menampilkan `window.APP_CONFIG = {...}`)

### 6. Update manifest.xml (setelah punya URL Vercel)

Di folder project, jalankan:

```powershell
npm run manifest:set -- https://NAMA-PROJECT.vercel.app
```

Contoh:

```powershell
npm run manifest:set -- https://excel-quiz-addin.vercel.app
```

### 7. Sideload di Excel

1. Buka Excel → **Insert** → **My Add-ins** → **Upload My Add-in**
2. Pilih file `manifest.xml` (sudah berisi URL Vercel)
3. Jalankan add-in dari daftar **Developer Add-ins**

---

## Opsi B — Deploy lewat CLI

```powershell
npm install -g vercel
cd g:\skripsi\excel-quiz-addin\excel-quiz-addin
vercel login
vercel
```

Ikuti pertanyaan (link ke project baru). Set environment variables di dashboard Vercel seperti Opsi A.

Production deploy:

```powershell
vercel --prod
```

Setelah dapat URL, update manifest:

```powershell
npm run manifest:set -- https://URL-DARI-VERCEL.vercel.app
```

---

## Development lokal vs production

| | Lokal | Vercel |
|---|--------|--------|
| Server | `npm run dev` (HTTPS localhost) | Static hosting |
| Manifest URL | `https://localhost:3000/...` | URL Vercel (setelah `manifest:set`) |
| Config Supabase | `.env` → Express `/js/config/app.config.js` | Env Vercel → `/api/app-config` |

Untuk develop di komputer, **jangan** ubah manifest ke Vercel — pakai localhost.  
Untuk uji production / sideload ke penguji, jalankan `manifest:set` dengan URL Vercel.

---

## Troubleshooting

| Masalah | Solusi |
|---------|--------|
| Build gagal "SUPABASE_URL" | Isi env variables di Vercel |
| Halaman blank / login gagal | Cek `/js/config/app.config.js` di browser |
| Add-in tidak muncul | Upload ulang `manifest.xml` setelah `manifest:set` |
| Icon tidak tampil | Tambahkan file PNG di `public/assets/` (icon-16, 32, 64, 80) |

---

## Domain kustom (nanti)

Jika nanti punya domain (mis. `quiz.kampus.ac.id`):

1. Vercel → Settings → Domains → tambahkan domain
2. Jalankan lagi: `npm run manifest:set -- https://quiz.kampus.ac.id`
