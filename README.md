# SheetDock + StatDock

**Dua alat kecil yang bikin Excel terasa lebih lega: navigasi sheet yang rapi di sisi layar, dan statistik seleksi yang bisa langsung "ditarik" jadi rumus.**

![Platform](https://img.shields.io/badge/platform-Excel%20for%20Windows-217346?logo=microsoftexcel&logoColor=white)
![Framework](https://img.shields.io/badge/.NET%20Framework-4.8-512BD4?logo=dotnet&logoColor=white)
![Excel-DNA](https://img.shields.io/badge/Excel--DNA-1.7.0-blue)
![Build](https://img.shields.io/badge/build-npm%20run%20build-CB3837?logo=npm&logoColor=white)

---

## Apa ini, dan kenapa kamu mungkin butuh?

Kalau kamu pernah bekerja dengan workbook Excel berisi **puluhan sheet**, kamu pasti kenal dua keluhan ini:

1. **Tab sheet di bawah layar sempit dan sulit dipakai.** Nama sheet terpotong, harus scroll kiri-kanan, dan mencari satu sheet di antara 30 tab itu melelahkan.
2. **Statistik di status bar hanya bisa dilihat, tidak bisa dipakai.** Excel menampilkan Average, Count, dan Sum dari sel yang kamu pilih, tetapi untuk memakai angka itu kamu tetap harus mengetik rumusnya sendiri.

**SheetDock + StatDock** menyelesaikan keduanya dalam satu add-in:

| | Masalah di Excel biasa | Dengan add-in ini |
|---|---|---|
| **SheetDock** | Tab sheet kecil, sempit, dan harus di-scroll | Panel daftar sheet di samping layar, satu "chip" per sheet, bisa diurutkan dengan drag |
| **StatDock** | Angka statistik hanya tampil, tidak bisa dipakai ulang | Chip statistik di status bar yang bisa **ditarik ke sel** menjadi rumus |

Tidak perlu belajar apa pun. Pasang, buka Excel, dan semuanya langsung terlihat.

---

## Tampilan sekilas

Berikut gambaran kasarnya (bukan screenshot asli):

```text
┌─────────────┬──────────────────────────────────────────────────────┐
│ SheetDock   │   A      B       C       D                           │
│             │ 1  Nama   Qty    Harga   Total                       │
│ ⠿ Ringkasan │ 2  Kopi   10     15000   150000                      │
│ ⠿ Penjualan │ 3  Teh    8      12000    96000                      │
│ ⠿ Stok      │ 4  Susu   5      18000    90000                      │
│ ⠿ Laporan   │ 5                                                    │
│ ⠿ Data 2024 │                                                      │
│ ⠿ Data 2025 │                                                      │
│    ...      │                                                      │
├─────────────┴──────────────────────────────────────────────────────┤
│ Ready   Sheet 2 of 12      Average  Count  Min  Max  [ Sum ]  100% │
└────────────────────────────────────────────────────────────────────┘
   ▲ SheetDock: daftar sheet           ▲ StatDock: tarik chip ke sel
```

---

## SheetDock: navigasi sheet yang waras

Panel di sisi kiri Excel (bisa dipindah ke sisi lain) yang menampilkan **semua sheet sebagai daftar chip**.

- **Satu klik pindah sheet.** Tidak perlu scroll tab. Worksheet dan chart sheet sama-sama didukung.
- **Drag untuk mengurutkan.** Tarik chip lewat gagang di sisinya untuk memindahkan posisi sheet. Daftar bergulir otomatis saat kamu menarik ke tepi.
- **Pilih banyak sheet sekaligus** dengan Ctrl atau Shift, misalnya untuk menduplikasi beberapa sheet dalam sekali jalan.
- **Warna tab ikut tampil**, jadi sistem pewarnaan sheet yang sudah kamu buat tetap berguna.
- **Klik kanan** membuka menu sheet bawaan Excel (Rename, Delete, Move or Copy, Tab Color, Hide, Unhide, dan seterusnya), ditambah empat menu ekstra:
  - **Add Sheet**
  - **Insert From File...** (menyisipkan sheet dari file Excel lain)
  - **Duplicate**
  - **Duplicate Selection** (menduplikasi semua sheet yang sedang dipilih)
- **Pintasan keyboard:** `F2` untuk rename, `Delete` untuk menghapus sheet terpilih.
- **Sheet tersembunyi tidak mengotori daftar.**
- **Tampilan mengikuti tema Office-mu** (White, Colorful, Dark Gray, atau Black), jadi tidak terlihat seperti benda asing.
- **Tab bawaan Excel disembunyikan** supaya tidak dobel. Bisa dimunculkan lagi kapan saja lewat ribbon.

---

## StatDock: statistik yang bisa kamu tarik

Strip tipis yang menempel di **status bar Excel** (di tempat Average/Count/Sum bawaan berada) dan menampilkan statistik dari sel yang sedang kamu pilih:

`Average` · `Count` · `Numerical Count` · `Min` · `Max` · `Sum`

Yang membuatnya berbeda dari status bar biasa:

- **Tarik chip ke sel, jadilah rumus.** Pilih A1:A10, tarik chip **Sum** ke sel kosong, dan sel itu berisi `=SUM(A1:A10)`. Selama menarik, ada pratinjau kecil yang menunjukkan alamat sel tujuan.
- **Paham filter dan baris tersembunyi.** Angka yang ditampilkan hanya dari sel yang **terlihat**. Rumus yang dihasilkan menyesuaikan: untuk data terfilter dipakai alamat area yang terlihat, dan jika terlalu panjang dipakai `SUBTOTAL(109, ...)` yang otomatis melewati baris tersembunyi.
- **Melindungi dari kesalahan.** Menjatuhkan chip ke dalam range sumbernya sendiri ditolak (menghindari *circular reference*), begitu juga menjatuhkannya di luar sel.
- **Menyesuaikan lebar.** Jika ruang sempit, chip yang kurang penting disembunyikan lebih dulu. `Sum` tidak pernah disembunyikan.
- **Ringan.** Pembaruan dilakukan setelah jeda singkat (debounce) dan tidak dijalankan saat kamu sedang drag-select atau Excel sedang menghitung, jadi Excel tidak terasa berat.
- **Mengikuti tema dan tampilan status bar.** Warna menyesuaikan Excel, dan strip ikut hilang jika status bar dimatikan.

---

## Ribbon

Add-in menambahkan satu tab **SheetDock** di ribbon dengan tiga tombol:

| Tombol | Fungsi |
|---|---|
| **Custom SheetDock** | Tampilkan atau sembunyikan panel daftar sheet |
| **Built-in Excel Tabs** | Tampilkan atau sembunyikan tab sheet bawaan Excel |
| **Quick Stats** | Aktifkan atau matikan strip StatDock di status bar |

---

## Persyaratan

- **Windows** dengan **Microsoft Excel desktop** (bukan Excel untuk Mac atau Excel di browser)
- **.NET Framework 4.8** (sudah bawaan Windows 10 dan 11 yang terbaru)
- Tersedia dua versi file: **64-bit** (`SheetDock64-packed.xll`) dan **32-bit** (`SheetDock-packed.xll`). Pilih sesuai versi Excel-mu. Mayoritas Excel modern adalah 64-bit.

---

## Instalasi

### Cara cepat (disarankan)

1. Pastikan file `SheetDock64-packed.xll` dan `Installer.bat` ada di folder yang sama.
2. **Klik dua kali `Installer.bat`.** Izinkan prompt Administrator.
3. Installer akan menutup Excel (simpan pekerjaanmu dulu), membuka blokir file unduhan, dan menyalin add-in ke folder AddIns milikmu.
4. Buka Excel. Jika tab **SheetDock** belum muncul, aktifkan lewat *File → Options → Add-ins → Manage: Excel Add-ins → Go*.

### Cara manual

1. Buka Excel, lalu *File → Options → Add-ins*.
2. Pilih **Manage: Excel Add-ins**, klik **Go...**, lalu **Browse...**.
3. Pilih file `.xll` yang sesuai dan centang add-in-nya.

> **Kalau Windows memblokir file:** klik kanan file `.xll` → *Properties* → centang **Unblock** → OK. `Installer.bat` sudah melakukan ini otomatis.

---

## Membangun dari source

Untuk developer. Yang kamu butuhkan hanya **Windows** dan **Node.js 16+**. Perintah `dotnet` tidak wajib sudah terpasang: jika tidak ditemukan, script mengunduh .NET SDK versi portable ke folder `.dotnet/` di dalam project (sekali saja, tanpa admin, tidak menyentuh PATH).

```powershell
npm run build             # build Release, hasil .xll muncul di folder project
npm run rebuild           # tutup Excel, hapus build lama, lalu build ulang
npm run rebuild:install   # sama seperti rebuild, lalu langsung jalankan Installer.bat
npm run clean             # hapus bin, obj, dan file .xll hasil build
```

Tidak perlu `npm install`: script hanya memakai modul bawaan Node.

Malas membuka terminal? Klik dua kali:

- `rebuild.bat` → tutup Excel, bersihkan, build ulang
- `rebuild-install.bat` → semua di atas, lalu pasang ke Excel

> ⚠️ `rebuild` dan `rebuild:install` **menutup Excel secara paksa** supaya file `.xll` lama tidak terkunci. Simpan pekerjaanmu dulu.

Hasil build berupa dua file di root project:

```text
SheetDock64-packed.xll   Excel 64-bit
SheetDock-packed.xll     Excel 32-bit
```

File `.xll` ini sudah *packed* (DLL dan dependency tertanam di dalamnya), jadi satu file per arsitektur sudah cukup untuk dibagikan.

---

## Struktur project

```text
SheetDock/
├── App/
│   └── CombinedAddIn.cs          Titik masuk add-in, merangkai kedua modul
├── UI/
│   └── SheetDockRibbon.cs        Tab dan tombol ribbon
├── SheetDock/Navigation/         Modul SheetDock
│   ├── NavigatorPane.cs          Panel daftar sheet (UI, drag, menu, aksi sheet)
│   ├── RecentSheets.cs           Riwayat sheet yang dikunjungi
│   ├── FolderStore.cs            Penyimpanan struktur folder virtual di dalam workbook
│   ├── Commands.cs               Makro untuk menu klik kanan
│   └── Globals.cs
├── StatDock/                     Modul StatDock
│   ├── Core/
│   │   ├── StatEngine.cs         Membaca seleksi dan menghitung statistik
│   │   ├── StatHost.cs           Siklus hidup, debounce, polling, pengaman refresh
│   │   ├── Commands.cs           Memasukkan rumus ke sel
│   │   └── Globals.cs
│   ├── UI/
│   │   ├── StatBar.cs            Strip di status bar
│   │   ├── StatChip.cs           Chip statistik yang bisa ditarik
│   │   ├── DragGhost.cs          Pratinjau saat drag
│   │   └── ThemeColors.cs        Warna mengikuti tema Excel
│   └── Infrastructure/
│       └── Native.cs             Pemanggilan Win32 (posisi jendela, DPI, mouse)
├── scripts/build.js              Skrip build npm
├── Installer.bat                 Installer add-in
├── rebuild.bat / rebuild-install.bat
├── SheetDock.csproj              Project .NET Framework 4.8 + Excel-DNA
└── package.json
```

---

## Cara kerja (untuk yang penasaran)

**Arsitektur.** Add-in dibangun dengan [Excel-DNA](https://excel-dna.net/), yang menjembatani C# ke Excel. `CombinedAddIn` adalah satu-satunya titik masuk dan menjalankan SheetDock serta StatDock sebagai dua modul di dalam add-in yang sama.

**Model refresh StatDock.** Membaca seleksi lewat COM itu mahal, jadi StatDock memakai tiga timer:

1. **Debounce:** menggabungkan banyak event Excel beruntun menjadi satu pembaruan.
2. **Track:** ringan dan sering; hanya menjaga posisi dan visibilitas strip, tanpa membaca data.
3. **Poll:** mendeteksi perubahan yang tidak selalu memicu event (misalnya filter atau baris disembunyikan), dibatasi frekuensinya.

Beberapa pengaman mencegah rantai pembaruan yang saling memicu: jendela *mute* setelah tiap pembaruan, penundaan saat tombol mouse ditekan, dan penundaan saat Excel sedang sibuk atau menghitung.

**Memasukkan rumus.** Rumus dimasukkan lewat clipboard dan perintah **Ctrl+V yang dikirim ke Excel**, sehingga dikerjakan Excel sebagai aksi pengguna. Mode copy Excel dibatalkan lebih dulu agar yang tertempel benar-benar rumus. Sesaat kemudian hasilnya diperiksa; jika sel tidak berisi rumus yang valid (misalnya karena nama fungsi berbahasa lokal), rumus ditulis langsung lewat `Range.Formula` dengan nama fungsi Inggris sebagai cadangan. Setiap drop dicatat di `%TEMP%\SheetDock.log`, berguna untuk melacak masalah.

---

## Pemecahan masalah

| Gejala | Kemungkinan penyebab dan solusi |
|---|---|
| Tab **SheetDock** tidak muncul | Aktifkan lewat *File → Options → Add-ins → Go*. Pastikan versi `.xll` (32/64-bit) cocok dengan Excel-mu. |
| Excel memblokir add-in | Klik kanan `.xll` → Properties → **Unblock**, atau jalankan `Installer.bat`. |
| `npm run build` gagal karena file terkunci | Excel masih memuat `.xll`. Pakai `npm run rebuild` (menutup Excel otomatis). |
| Muncul error `EPERM` saat clean | Sama: Excel masih membuka file di `bin`. Tutup Excel atau jalankan `npm run rebuild`. |
| Chip StatDock menutupi teks status bar | Sesuaikan konstanta `KeepNativeLeftPx` dan `KeepNativeRightPx` di `StatDock/UI/StatBar.cs`. |
| Rumus dari StatDock tidak muncul | Lihat `%TEMP%\SheetDock.log`; baris terakhir menjelaskan jalur yang dipakai dan alasannya. |
| Ctrl+C / Ctrl+V tidak bereaksi setelah klik sheet di panel | Pastikan memakai versi terbaru; fokus keyboard dikembalikan ke Excel setelah klik biasa di panel. |

---

## Daftar uji manual

Setelah build, cek minimal hal-hal ini:

- [ ] Add-in termuat dan tab **SheetDock** muncul
- [ ] Panel daftar sheet tampil dan klik chip berpindah sheet
- [ ] Drag chip sheet mengubah urutan sheet
- [ ] Klik kanan menampilkan menu Excel plus empat menu ekstra
- [ ] Strip StatDock muncul di status bar dan angkanya berubah saat seleksi berubah
- [ ] Filter data atau sembunyikan baris: angka StatDock ikut menyesuaikan
- [ ] Tarik chip **Sum** ke sel kosong: rumus muncul dan tidak ada isi lain yang ikut
- [ ] Copy sel, klik sheet lain di panel, lalu Ctrl+V: tetap berfungsi normal
- [ ] Tidak ada loading cursor yang berputar terus setelah drag

---

## Catatan

- Hanya untuk **Excel desktop di Windows**.
- Lisensi belum ditentukan. Tambahkan file `LICENSE` sebelum project ini dibagikan secara publik.
