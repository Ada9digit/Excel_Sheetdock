#!/usr/bin/env node
/**
 * SheetDock build script (zero dependency, Node built-ins only).
 *
 *   npm run build      -> dotnet build Release, hasil *-packed.xll di folder project ini
 *   npm run rebuild    -> tutup Excel, hapus bin/obj/xll lama, lalu build
 *   npm run rebuild:install -> seperti rebuild, lalu jalankan Installer.bat
 *   npm run clean      -> hapus bin, obj, dan *-packed.xll di root saja
 *
 * Catatan: C# tetap dikompilasi oleh .NET SDK (dotnet). Jika dotnet belum
 * terpasang, script mengunduh SDK portable ke ./.dotnet (tanpa install).
 * Node mengorkestrasi: siapkan SDK, build, salin hasil ke folder project.
 */
'use strict';

const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..');
const CSPROJ = path.join(ROOT, 'SheetDock.csproj');
const BIN_DIR = path.join(ROOT, 'bin');
const OBJ_DIR = path.join(ROOT, 'obj');
const LOCAL_DOTNET_DIR = path.join(ROOT, '.dotnet');
const LOCAL_DOTNET_EXE = path.join(LOCAL_DOTNET_DIR, 'dotnet.exe');
const DOTNET_CHANNEL = '8.0';
const CONFIG = 'Release';
const TFM = 'net48';

const args = process.argv.slice(2);
const cleanOnly = args.includes('--clean-only');
const clean = cleanOnly || args.includes('--clean');
const killExcel = args.includes('--kill-excel');
const runInstaller = args.includes('--install');

const log = (msg) => console.log(`[sheetdock] ${msg}`);
const fail = (msg, code = 1) => {
  console.error(`\n[sheetdock] ERROR: ${msg}\n`);
  process.exit(code);
};

function rmrf(dir) {
  try {
    fs.rmSync(dir, { recursive: true, force: true, maxRetries: 3, retryDelay: 200 });
  } catch (err) {
    if (err && (err.code === 'EPERM' || err.code === 'EBUSY' || err.code === 'EACCES')) {
      fail(
        `Tidak bisa menghapus: ${dir}\n` +
        '  Biasanya file .xll di dalamnya sedang dipakai Excel.\n' +
        '  Tutup Excel lalu ulangi, atau jalankan:  npm run rebuild -- --kill-excel'
      );
    }
    throw err;
  }
}

function ensureExcelClosed() {
  if (!isExcelRunning()) return;
  if (killExcel) {
    log('Menutup Excel (--kill-excel) ...');
    spawnSync('taskkill', ['/F', '/IM', 'EXCEL.EXE'], { stdio: 'ignore' });
    // beri waktu Windows melepas lock file
    spawnSync('powershell', ['-NoProfile', '-Command', 'Start-Sleep -Seconds 2'], { stdio: 'ignore' });
    if (isExcelRunning()) fail('Excel masih berjalan setelah taskkill. Tutup manual lalu ulangi.');
    return;
  }
  fail(
    'Excel sedang berjalan, file .xll di bin kemungkinan terkunci.\n' +
    '  Tutup Excel lalu ulangi, atau jalankan:  npm run rebuild -- --kill-excel'
  );
}

function run(cmd, cmdArgs, opts = {}) {
  return spawnSync(cmd, cmdArgs, { cwd: ROOT, stdio: 'inherit', ...opts });
}

function removeRootXlls() {
  for (const f of fs.readdirSync(ROOT)) {
    if (/-packed\.xll$/i.test(f)) fs.rmSync(path.join(ROOT, f), { force: true });
  }
}

function ps(cmdArgs) {
  return run('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', ...cmdArgs]);
}

/**
 * Cari dotnet: pakai yang ada di PATH, kalau tidak ada pakai ./.dotnet,
 * kalau belum ada juga, unduh .NET SDK portable ke ./.dotnet (tanpa admin,
 * tanpa mengubah PATH/registry). Hanya perlu internet di run pertama.
 */
function resolveDotnet() {
  const sys = spawnSync('dotnet', ['--version'], { encoding: 'utf8' });
  if (!sys.error && sys.status === 0) return { exe: 'dotnet', version: sys.stdout.trim(), local: false };

  if (!fs.existsSync(LOCAL_DOTNET_EXE)) {
    log(`dotnet tidak ada di PATH. Mengunduh .NET SDK ${DOTNET_CHANNEL} portable ke ${LOCAL_DOTNET_DIR} ...`);
    log('(sekali saja, sekitar 200-300 MB, tidak menginstall apa pun ke sistem)');
    const script = path.join(ROOT, 'scripts', '.dotnet-install.ps1');
    let r = ps(['-Command',
      "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; " +
      `Invoke-WebRequest -UseBasicParsing -Uri https://dot.net/v1/dotnet-install.ps1 -OutFile '${script}'`]);
    if (r.status !== 0) fail('Gagal mengunduh dotnet-install.ps1. Cek koneksi internet / proxy.');
    r = ps(['-File', script, '-Channel', DOTNET_CHANNEL, '-InstallDir', LOCAL_DOTNET_DIR, '-NoPath']);
    fs.rmSync(script, { force: true });
    if (r.status !== 0 || !fs.existsSync(LOCAL_DOTNET_EXE)) fail('Gagal memasang .NET SDK portable.');
  }
  const loc = spawnSync(LOCAL_DOTNET_EXE, ['--version'], { encoding: 'utf8' });
  if (loc.error || loc.status !== 0) {
    fail('.NET SDK portable di .dotnet rusak. Hapus folder .dotnet lalu jalankan lagi npm run build.');
  }
  return { exe: LOCAL_DOTNET_EXE, version: loc.stdout.trim(), local: true };
}

function isExcelRunning() {
  const r = spawnSync('tasklist', ['/FI', 'IMAGENAME eq EXCEL.EXE', '/NH'], { encoding: 'utf8' });
  return r.status === 0 && /EXCEL\.EXE/i.test(r.stdout || '');
}

function findPackedXlls(startDir) {
  // Excel-DNA menaruh hasil pack di <OutDir>\publish, tapi cari rekursif
  // di bin supaya tetap jalan kalau lokasinya berubah.
  const found = [];
  (function walk(dir) {
    if (!fs.existsSync(dir)) return;
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) walk(full);
      else if (/-packed\.xll$/i.test(entry.name)) found.push(full);
    }
  })(startDir);
  return found;
}

function main() {
  if (clean || killExcel) ensureExcelClosed();

  if (clean) {
    log('Membersihkan bin, obj, dan *-packed.xll ...');
    [BIN_DIR, OBJ_DIR].forEach(rmrf);
    removeRootXlls();
    if (cleanOnly) return log('Selesai.');
  }

  if (process.platform !== 'win32') {
    fail('Project ini (net48 + WinForms + Excel-DNA) hanya bisa di-build di Windows.');
  }
  if (!fs.existsSync(CSPROJ)) fail(`SheetDock.csproj tidak ditemukan di ${ROOT}`);

  const dotnet = resolveDotnet();
  log(`dotnet ${dotnet.version}${dotnet.local ? ' (portable, .dotnet)' : ''}, node ${process.version}`);
  const env = {
    ...process.env,
    DOTNET_CLI_TELEMETRY_OPTOUT: '1',
    DOTNET_NOLOGO: '1',
    DOTNET_SKIP_FIRST_TIME_EXPERIENCE: '1',
    ...(dotnet.local ? { DOTNET_ROOT: LOCAL_DOTNET_DIR, DOTNET_MULTILEVEL_LOOKUP: '0' } : {}),
  };

  if (isExcelRunning()) {
    log('PERINGATAN: Excel sedang berjalan. Jika add-in dimuat dari folder bin, file XLL bisa terkunci.');
  }

  // Hasil build lama dihapus supaya tidak ada file basi di root.
  removeRootXlls();
  rmrf(path.join(BIN_DIR, CONFIG, TFM, 'publish'));

  log(`dotnet build -c ${CONFIG} ...`);
  const build = run(dotnet.exe, ['build', CSPROJ, '-c', CONFIG, '-nologo', '-v', 'minimal'], { env });
  if (build.status !== 0) fail('dotnet build gagal. Lihat pesan error di atas.', build.status || 1);

  const xlls = findPackedXlls(path.join(BIN_DIR, CONFIG, TFM));
  if (xlls.length === 0) {
    fail(
      'Build sukses tetapi file *-packed.xll tidak ditemukan.\n' +
      `  Periksa ${path.join(BIN_DIR, CONFIG, TFM, 'publish')} dan pastikan paket ExcelDna.AddIn ter-restore.`
    );
  }

  const copied = [];
  for (const src of xlls) {
    const name = path.basename(src);
    fs.copyFileSync(src, path.join(ROOT, name));
    copied.push(name);
  }

  log(`Selesai. Hasil ada di folder ini: ${ROOT}`);
  for (const f of copied) {
    const kb = (fs.statSync(path.join(ROOT, f)).size / 1024).toFixed(0);
    console.log(`           - ${f} (${kb} KB)`);
  }

  if (runInstaller) {
    const installer = path.join(ROOT, 'Installer.bat');
    if (!fs.existsSync(installer)) fail('Installer.bat tidak ditemukan di folder project.');
    log('Menjalankan Installer.bat (akan meminta izin Administrator) ...');
    const r = spawnSync('cmd', ['/c', 'call', installer], { cwd: ROOT, stdio: 'inherit' });
    if (r.status !== 0 && r.status !== null) log(`Installer.bat selesai dengan kode ${r.status}.`);
  }
}

main();
