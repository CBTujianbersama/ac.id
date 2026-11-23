<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>CBT-PRO | Sistem Ujian Berbasis Komputer</title>
    
    <!-- Tailwind CSS CDN (Murni tanpa bentrok CSS statis) -->
    <script src="https://cdn.tailwindcss.com"></script>
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    fontFamily: {
                        sans: ['Inter', 'sans-serif'],
                    }
                }
            }
        };
    </script>

    <!-- Google Fonts: Inter -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    
    <!-- Lucide Icons CDN -->
    <script src="https://unpkg.com/lucide@latest"></script>
    <!-- SheetJS (XLSX) CDN untuk Import/Export Excel -->
    <script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
    <!-- Mammoth.js CDN untuk membaca file Word (.docx) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js"></script>
    <!-- PDF.js CDN untuk membaca file PDF (.pdf) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
    <script>
        if (typeof pdfjsLib !== 'undefined') {
            pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        }
    </script>

    <style>
        *, *::before, *::after {
            box-sizing: border-box;
        }

        html, body {
            margin: 0;
            padding: 0;
            width: 100%;
            min-height: 100vh;
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
            background-color: #f8fafc;
            color: #0f172a;
            overflow-x: hidden;
            -webkit-font-smoothing: antialiased;
        }

        /* Mencegah elemen grid/flex melebar keluar layar hingga tumpang tindih */
        .grid > *, .flex > * {
            min-width: 0;
        }

        /* Mencegah seleksi teks saat siswa mengerjakan ujian */
        .no-select {
            -webkit-user-select: none;
            -moz-user-select: none;
            -ms-user-select: none;
            user-select: none;
        }

        /* Animasi latar belakang layar terkunci anti-curang */
        @keyframes lockPulse {
            0%, 100% { background-color: rgba(127, 29, 29, 0.98); }
            50% { background-color: rgba(185, 28, 28, 0.98); }
        }
        .lock-overlay-bg {
            animation: lockPulse 2s infinite;
        }

        /* Aturan Cetak (Print A4 2 Kolom) Khusus Kartu Login Peserta */
        @media print {
            body * {
                visibility: hidden;
            }
            #printCardsContainer, #printCardsContainer * {
                visibility: visible;
            }
            #printCardsContainer {
                position: absolute;
                left: 0;
                top: 0;
                width: 100%;
                display: grid !important;
                grid-template-columns: repeat(2, 1fr) !important;
                gap: 12px !important;
                padding: 8px !important;
                background: #ffffff !important;
            }
            .no-print {
                display: none !important;
            }
            .login-card-item {
                break-inside: avoid;
                page-break-inside: avoid;
                border: 1.5px solid #0f172a !important;
                box-shadow: none !important;
            }
        }

        /* Scrollbar halus */
        ::-webkit-scrollbar {
            width: 6px;
            height: 6px;
        }
        ::-webkit-scrollbar-track {
            background: #f1f5f9;
        }
        ::-webkit-scrollbar-thumb {
            background: #cbd5e1;
            border-radius: 4px;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: #94a3b8;
        }
    </style>
</head>
<body class="min-h-screen flex flex-col bg-slate-50 text-slate-900">

    <!-- ===================================================================== -->
    <!-- 1. SISTEM TOAST & DIALOG KONFIRMASI (PENGGANTI ALERT & CONFIRM)       -->
    <!-- ===================================================================== -->
    <div id="toastContainer" class="fixed top-4 right-4 z-[9999] flex flex-col gap-2 max-w-sm w-full pointer-events-none px-4 sm:px-0"></div>

    <div id="confirmModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[9990] items-center justify-center p-4 hidden">
        <div class="bg-white rounded-2xl max-w-md w-full p-6 shadow-2xl border border-slate-200">
            <div class="flex items-center gap-3 mb-3">
                <div class="w-10 h-10 rounded-xl bg-amber-100 text-amber-600 flex items-center justify-center shrink-0">
                    <i data-lucide="alert-triangle" class="w-5 h-5"></i>
                </div>
                <h3 id="confirmTitle" class="text-lg font-bold text-slate-900">Konfirmasi</h3>
            </div>
            <p id="confirmMessage" class="text-sm text-slate-600 mb-6 leading-relaxed">Apakah Anda yakin?</p>
            <div class="flex justify-end gap-2.5">
                <button id="confirmCancelBtn" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 bg-slate-100 hover:bg-slate-200 transition">Batal</button>
                <button id="confirmOkBtn" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-white bg-blue-600 hover:bg-blue-700 transition shadow-sm">Ya, Lanjutkan</button>
            </div>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 2. OVERLAY LAYAR TERKUNCI ANTI-CURANG + SIRINE ALARM OTOMATIS         -->
    <!-- ===================================================================== -->
    <div id="antiCheatLockOverlay" class="fixed inset-0 lock-overlay-bg z-[99999] flex-col items-center justify-center p-4 sm:p-6 text-white text-center select-none hidden">
        <div class="max-w-lg w-full bg-red-950/95 border-2 border-red-400 rounded-3xl p-6 sm:p-8 shadow-2xl">
            <div class="w-20 h-20 rounded-full bg-red-600/30 border-2 border-red-400 flex items-center justify-center mx-auto mb-4 animate-bounce">
                <i data-lucide="siren" class="w-10 h-10 text-amber-300"></i>
            </div>
            <div class="flex items-center justify-center gap-2 mb-3">
                <span class="inline-block px-3 py-1 rounded-full bg-red-600 text-white text-xs font-extrabold uppercase tracking-wider">
                    SIRINE ALARM AKTIF • PELANGGARAN TERDETEKSI
                </span>
            </div>
            <h1 class="text-2xl sm:text-3xl font-extrabold tracking-tight mb-2">LAYAR UJIAN TERKUNCI!</h1>
            <p id="lockReasonText" class="text-red-200 text-xs sm:text-sm mb-5 leading-relaxed">Layar ujian kehilangan fokus! Dilarang membuka jendela lain, split-screen, atau berpindah aplikasi.</p>

            <div class="bg-red-900/70 border border-red-700 rounded-2xl p-4 mb-5 text-left text-xs space-y-2">
                <div class="flex justify-between gap-2">
                    <span class="text-red-300 shrink-0">Nama Peserta:</span>
                    <span id="lockStudentName" class="font-bold text-white text-right truncate">-</span>
                </div>
                <div class="flex justify-between gap-2">
                    <span class="text-red-300 shrink-0">NISN / Kelas:</span>
                    <span id="lockStudentMeta" class="font-semibold text-white text-right truncate">-</span>
                </div>
                <div class="flex justify-between gap-2">
                    <span class="text-red-300 shrink-0">Mata Pelajaran Ujian:</span>
                    <span id="lockExamSubject" class="font-semibold text-amber-200 text-right truncate">-</span>
                </div>
                <div class="flex justify-between gap-2">
                    <span class="text-red-300 shrink-0">Total Pelanggaran Tercatat:</span>
                    <span id="lockViolationBadge" class="font-extrabold text-amber-300">1 Kali</span>
                </div>
                <div class="flex justify-between gap-2">
                    <span class="text-red-300 shrink-0">Status Alarm &amp; Log:</span>
                    <span class="font-semibold text-emerald-300 text-right">Sirine Menyala • Terkirim ke Pengawas &amp; Admin</span>
                </div>
            </div>

            <div class="space-y-3">
                <label class="block text-xs font-bold uppercase tracking-wider text-red-200">
                    Masukkan Kode PIN Pengawas untuk Mematikan Sirine &amp; Membuka Kunci
                </label>
                <input type="password" id="unlockPinField" placeholder="Masukkan Kode PIN Pengawas" class="w-full px-4 py-3.5 rounded-xl bg-white text-slate-900 font-extrabold text-center text-lg tracking-widest focus:outline-none focus:ring-4 focus:ring-amber-400" onkeypress="if(event.key === 'Enter') verifyPinUnlock()">
                <button onclick="verifyPinUnlock()" class="w-full py-3.5 px-6 rounded-xl bg-amber-400 hover:bg-amber-300 text-slate-950 font-extrabold text-sm transition shadow-lg flex items-center justify-center gap-2">
                    <i data-lucide="unlock" class="w-4 h-4"></i> MATIKAN SIRINE &amp; BUKA KUNCI
                </button>
            </div>
            <p class="mt-4 text-[11px] text-red-300/90">
                Sirine peringatan akan terus berbunyi sampai Pengawas Ruangan memverifikasi PIN atau membuka kunci jarak jauh.
            </p>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 3. HEADER / NAVBAR UTAMA                                              -->
    <!-- ===================================================================== -->
    <header id="mainHeader" class="bg-white/95 backdrop-blur-md border-b border-slate-200 sticky top-0 z-40 shadow-sm no-print">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div class="flex items-center justify-between min-h-[64px] py-2.5 gap-3">
                <div class="flex items-center gap-3 cursor-pointer min-w-0 flex-1" onclick="handleHeaderBrandClick()">
                    <div class="w-10 h-10 rounded-xl bg-gradient-to-br from-blue-600 to-indigo-700 text-white flex items-center justify-center shadow-md shrink-0">
                        <i data-lucide="shield-check" class="w-6 h-6"></i>
                    </div>
                    <div class="min-w-0 flex-1">
                        <div class="flex items-center gap-2">
                            <h1 id="headerSchoolTitle" class="font-extrabold text-sm sm:text-base text-slate-900 truncate">CBT-PRO | Sistem Ujian Berbasis Komputer</h1>
                            <span class="hidden md:inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-700 text-[10px] font-bold shrink-0">
                                <span class="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-ping"></span> LIVE SYNC
                            </span>
                        </div>
                        <p id="headerSchoolSubtitle" class="text-[11px] text-slate-500 truncate">Portal Evaluasi Akademik Terpadu • Tahun Pelajaran 2026/2027</p>
                    </div>
                </div>

                <!-- Tombol Aksi Header -->
                <div class="flex items-center gap-2 shrink-0">
                    <div id="headerGuestButtons" class="flex items-center gap-1.5 sm:gap-2">
                        <button onclick="openLoginModal('siswa')" class="px-3 py-2 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold transition flex items-center gap-1.5 shadow-sm">
                            <i data-lucide="graduation-cap" class="w-4 h-4"></i>
                            <span>Login Siswa</span>
                        </button>
                        <button onclick="openLoginModal('pengawas')" class="px-3 py-2 rounded-xl bg-amber-500 hover:bg-amber-600 text-white text-xs font-bold transition hidden sm:flex items-center gap-1.5">
                            <i data-lucide="user-check" class="w-4 h-4"></i>
                            <span>Pengawas</span>
                        </button>
                        <button onclick="openLoginModal('admin')" class="px-3 py-2 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition hidden sm:flex items-center gap-1.5">
                            <i data-lucide="settings" class="w-4 h-4"></i>
                            <span>Admin</span>
                        </button>
                    </div>

                    <div id="headerAuthControls" class="hidden items-center gap-1.5 sm:gap-2 flex-wrap justify-end">
                        <span id="headerUserBadge" class="hidden md:inline-block px-3 py-1.5 rounded-xl bg-slate-100 text-slate-700 text-xs font-bold max-w-[200px] truncate">
                            -
                        </span>
                        <button id="headerBackDashBtn" onclick="returnToRoleDashboard()" class="hidden px-3 py-2 rounded-xl bg-blue-50 hover:bg-blue-100 text-blue-700 text-xs font-bold transition">
                            Dashboard Saya
                        </button>
                        <button onclick="navigateTo('portal')" class="px-3 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold transition">
                            Portal
                        </button>
                        <button onclick="logoutSession()" class="px-3 py-2 rounded-xl bg-red-50 hover:bg-red-100 text-red-600 text-xs font-bold transition flex items-center gap-1">
                            <i data-lucide="log-out" class="w-3.5 h-3.5"></i> Keluar
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </header>

    <!-- ===================================================================== -->
    <!-- 4. KONTEN UTAMA (PORTAL, DASHBOARD SISWA, PENGAWAS, ADMIN, UJIAN)     -->
    <!-- ===================================================================== -->
    <main class="flex-grow w-full">

        <!-- ================================================================= -->
        <!-- VIEW A: HALAMAN PORTAL UTAMA (LANDING PORTAL)                     -->
        <!-- ================================================================= -->
        <section id="view-portal" class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-10">
            <div class="bg-gradient-to-br from-slate-900 via-blue-950 to-indigo-950 rounded-3xl p-6 sm:p-10 text-white shadow-xl relative overflow-hidden">
                <div class="absolute -top-24 -right-24 w-80 h-80 bg-blue-500/15 rounded-full blur-3xl pointer-events-none"></div>
                <div class="relative z-10 grid grid-cols-1 lg:grid-cols-12 gap-8 items-center">
                    <div class="lg:col-span-7 space-y-5">
                        <div class="inline-flex items-center gap-2 px-3 py-1.5 rounded-full bg-blue-500/20 border border-blue-400/30 text-blue-200 text-xs font-semibold">
                            <i data-lucide="shield-alert" class="w-4 h-4 text-amber-300 shrink-0"></i>
                            <span>Sistem Ujian Multi-Mapel • Token Sesi • Sirine Anti-Curang</span>
                        </div>
                        <h2 class="text-2xl sm:text-4xl font-extrabold tracking-tight leading-tight">
                            Selamat Datang di Portal <span class="text-blue-400" id="heroExamTitle">Ujian Akhir Semester (UAS) Berbasis Komputer</span>
                        </h2>
                        <p class="text-slate-300 text-sm sm:text-base leading-relaxed">
                            Silakan login ke <strong>Dashboard Siswa</strong> menggunakan <strong>NISN / Username</strong> dan <strong>Password</strong> Kartu Login untuk melihat profil, jadwal penugasan mata pelajaran, serta memasukkan <strong>Token Ujian</strong> saat soal sudah siap dikerjakan.
                        </p>

                        <!-- Tombol Login 3 Peran -->
                        <div class="flex flex-wrap gap-3 pt-1">
                            <button onclick="openLoginModal('siswa')" class="px-5 py-3 rounded-xl bg-blue-600 hover:bg-blue-500 text-white font-bold text-xs sm:text-sm transition flex items-center gap-2 shadow-lg shadow-blue-950/60">
                                <i data-lucide="graduation-cap" class="w-4 h-4"></i> Login Dashboard Siswa
                            </button>
                            <button onclick="openLoginModal('pengawas')" class="px-4 py-3 rounded-xl bg-amber-500 hover:bg-amber-400 text-slate-950 font-bold text-xs sm:text-sm transition flex items-center gap-2">
                                <i data-lucide="eye" class="w-4 h-4"></i> Portal Pengawas
                            </button>
                            <button onclick="openLoginModal('admin')" class="px-4 py-3 rounded-xl bg-white/15 hover:bg-white/25 text-white border border-white/20 font-bold text-xs sm:text-sm transition flex items-center gap-2">
                                <i data-lucide="sliders" class="w-4 h-4"></i> Administrator
                            </button>
                        </div>

                        <!-- Tombol Salin Link Akses -->
                        <div class="pt-3 border-t border-white/10 flex flex-wrap items-center gap-2 text-xs">
                            <span class="text-slate-400 font-medium mr-1">Bagikan Link Akses:</span>
                            <button onclick="copyPortalLink('siswa')" class="px-3 py-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-blue-200 font-semibold transition flex items-center gap-1.5">
                                <i data-lucide="link" class="w-3.5 h-3.5"></i> Link Portal Siswa
                            </button>
                            <button onclick="copyPortalLink('pengawas')" class="px-3 py-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-amber-200 font-semibold transition flex items-center gap-1.5">
                                <i data-lucide="link" class="w-3.5 h-3.5"></i> Link Pengawas
                            </button>
                            <button onclick="copyPortalLink('admin')" class="px-3 py-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-indigo-200 font-semibold transition flex items-center gap-1.5">
                                <i data-lucide="link" class="w-3.5 h-3.5"></i> Link Administrator
                            </button>
                        </div>
                    </div>

                    <!-- Kartu Motivasi & Statistik Live -->
                    <div class="lg:col-span-5">
                        <div class="bg-white/10 backdrop-blur-md border border-white/20 rounded-3xl p-6 space-y-4">
                            <div class="flex items-center justify-between gap-2">
                                <span class="inline-flex items-center gap-1.5 text-xs font-bold uppercase tracking-wider text-amber-300">
                                    <i data-lucide="sparkles" class="w-4 h-4 shrink-0"></i> Motivasi Belajar &amp; Kejujuran
                                </span>
                                <button onclick="nextMotivationQuote()" class="px-2.5 py-1 rounded-lg bg-white/10 hover:bg-white/20 text-xs font-semibold text-white transition flex items-center gap-1 shrink-0">
                                    <i data-lucide="refresh-cw" class="w-3.5 h-3.5"></i> Acak
                                </button>
                            </div>
                            <blockquote class="space-y-3">
                                <p id="motivationQuoteText" class="text-sm sm:text-base font-medium italic text-white leading-relaxed">
                                    "Prestasi sejati lahir dari kerja keras dan kejujuran. Nilai tinggi tanpa kejujuran hanyalah angka, sedangkan kejujuran adalah bekal kesuksesan masa depanmu."
                                </p>
                                <footer id="motivationQuoteAuthor" class="text-xs font-bold text-blue-300">
                                    — Pesan Akademik Sekolah
                                </footer>
                            </blockquote>
                            <div class="pt-3 border-t border-white/10 grid grid-cols-3 gap-2 text-center">
                                <div class="bg-slate-900/50 rounded-xl p-2.5">
                                    <div id="portalStatSoal" class="text-lg font-extrabold text-white">0</div>
                                    <div class="text-[10px] text-slate-400">Paket Ujian</div>
                                </div>
                                <div class="bg-slate-900/50 rounded-xl p-2.5">
                                    <div id="portalStatSiswa" class="text-lg font-extrabold text-white">0</div>
                                    <div class="text-[10px] text-slate-400">Peserta Aktif</div>
                                </div>
                                <div class="bg-slate-900/50 rounded-xl p-2.5">
                                    <div id="portalStatDurasi" class="text-lg font-extrabold text-amber-300">0</div>
                                    <div class="text-[10px] text-slate-400">Terjadwal</div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- BAGIAN PANDUAN SISTEM UJIAN CBT (SISWA, PENGAWAS, ADMIN) -->
            <div class="space-y-5">
                <div class="flex flex-col sm:flex-row sm:items-end justify-between gap-2">
                    <div>
                        <span class="text-xs font-bold uppercase tracking-wider text-blue-600">Tata Tertib &amp; Alur Pelaksanaan</span>
                        <h2 class="text-xl sm:text-2xl font-extrabold text-slate-900 mt-1">Alur Sistem Ujian CBT Terpadu</h2>
                    </div>
                    <p class="text-xs text-slate-500">Sinkronisasi real-time antara Admin, Pengawas, dan Siswa</p>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
                    <!-- Panduan Siswa -->
                    <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                        <div class="flex items-center gap-3">
                            <div class="w-11 h-11 rounded-2xl bg-blue-100 text-blue-600 flex items-center justify-center shrink-0">
                                <i data-lucide="user-check" class="w-6 h-6"></i>
                            </div>
                            <div>
                                <h3 class="font-bold text-slate-900 text-base">1. Alur Peserta / Siswa</h3>
                                <span class="text-xs text-blue-600 font-semibold">Dashboard &amp; Pemilihan Soal</span>
                            </div>
                        </div>
                        <ol class="space-y-2.5 text-xs text-slate-600 list-decimal pl-4 leading-relaxed">
                            <li><strong>Login ke Dashboard Siswa:</strong> Masuk dengan NISN/Username &amp; Password Kartu Login (tanpa langsung masuk soal).</li>
                            <li><strong>Cek Profil &amp; Jadwal Penugasan:</strong> Lihat mata pelajaran yang sudah di-assign oleh Admin atau masih menunggu jadwal.</li>
                            <li><strong>Masukkan Token saat Siap:</strong> Klik ujian yang berstatus <em>Siap Dikerjakan</em> dan input <strong>Token Aktif</strong> dari Pengawas.</li>
                            <li><strong>Kunci &amp; Sirine Anti-Curang:</strong> Dilarang membuka aplikasi/tab lain saat ujian berlangsung atau layar akan terkunci disertai bunyi sirine.</li>
                        </ol>
                    </div>

                    <!-- Panduan Pengawas -->
                    <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                        <div class="flex items-center gap-3">
                            <div class="w-11 h-11 rounded-2xl bg-amber-100 text-amber-600 flex items-center justify-center shrink-0">
                                <i data-lucide="shield" class="w-6 h-6"></i>
                            </div>
                            <div>
                                <h3 class="font-bold text-slate-900 text-base">2. Alur Pengawas Ruang</h3>
                                <span class="text-xs text-amber-600 font-semibold">3 Menu Kendali Pengawasan</span>
                            </div>
                        </div>
                        <ul class="space-y-2.5 text-xs text-slate-600 list-disc pl-4 leading-relaxed">
                            <li><strong>Pantau Keluar/Masuk &amp; Token:</strong> Lihat log kehadiran siswa secara live beserta Token Ujian yang sinkron otomatis dengan Admin.</li>
                            <li><strong>Monitor Pelanggaran &amp; Sirine:</strong> Deteksi siswa yang keluar aplikasi, matikan alarm &amp; buka kunci layar jarak jauh.</li>
                            <li><strong>Input &amp; Upload Soal:</strong> Input soal manual atau upload file Word (.docx) / PDF (.pdf) per mata pelajaran yang langsung tersinkron ke Bank Soal.</li>
                        </ul>
                    </div>

                    <!-- Panduan Administrator -->
                    <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                        <div class="flex items-center gap-3">
                            <div class="w-11 h-11 rounded-2xl bg-indigo-100 text-indigo-600 flex items-center justify-center shrink-0">
                                <i data-lucide="cpu" class="w-6 h-6"></i>
                            </div>
                            <div>
                                <h3 class="font-bold text-slate-900 text-base">3. Alur Administrator</h3>
                                <span class="text-xs text-indigo-600 font-semibold">Penjadwalan &amp; Audit Log</span>
                            </div>
                        </div>
                        <ul class="space-y-2.5 text-xs text-slate-600 list-disc pl-4 leading-relaxed">
                            <li><strong>Bank Soal Per Mata Pelajaran:</strong> Siapkan paket soal per mata pelajaran lewat upload Word/PDF atau input manual.</li>
                            <li><strong>Assign Jadwal Mata Pelajaran:</strong> Pilih soal yang sudah disiapkan di menu Assignment, tentukan tanggal/jam jadwal &amp; kelas target.</li>
                            <li><strong>Log Aktivitas Masuk/Keluar:</strong> Pantau riwayat login, logout, mulai ujian, dan pelanggaran seluruh Siswa &amp; Pengawas.</li>
                        </ul>
                    </div>
                </div>
            </div>

            <!-- BAGIAN GALERI FOTO KEGIATAN UJIAN -->
            <div class="space-y-5">
                <div class="flex flex-col sm:flex-row sm:items-end justify-between gap-2">
                    <div>
                        <span class="text-xs font-bold uppercase tracking-wider text-emerald-600">Dokumentasi &amp; Suasana Ujian</span>
                        <h2 class="text-xl sm:text-2xl font-extrabold text-slate-900 mt-1">Galeri Kegiatan Ujian Berbasis Komputer</h2>
                    </div>
                    <p class="text-xs text-slate-500">Dokumentasi kegiatan dapat dikelola melalui menu Setelan Administrator</p>
                </div>
                <div id="portalGalleryGrid" class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-6"></div>
            </div>

            <!-- BAGIAN DOWNLOAD & INSTALASI APLIKASI -->
            <div class="bg-white rounded-3xl border border-slate-200 p-6 sm:p-8 shadow-sm space-y-6">
                <div class="flex flex-col md:flex-row md:items-center justify-between gap-4 pb-5 border-b border-slate-100">
                    <div>
                        <span class="text-xs font-bold uppercase tracking-wider text-blue-600">Aplikasi Multi-Platform</span>
                        <h2 class="text-xl font-extrabold text-slate-900 mt-1">Download &amp; Pasang Aplikasi di HP Android, iPhone, maupun Komputer</h2>
                        <p class="text-xs sm:text-sm text-slate-500 mt-1">Pasang sebagai aplikasi mandiri layar penuh untuk kenyamanan dan keamanan ujian maksimal.</p>
                    </div>
                    <button onclick="downloadStandaloneApp()" class="px-4 py-3 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition flex items-center gap-2 self-start md:self-center shrink-0 shadow-sm">
                        <i data-lucide="download" class="w-4 h-4"></i> Unduh File Aplikasi (.HTML)
                    </button>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
                    <div class="rounded-2xl border border-emerald-200 bg-emerald-50/40 p-5 flex flex-col justify-between space-y-4">
                        <div class="space-y-3">
                            <div class="flex items-center gap-3">
                                <div class="w-11 h-11 rounded-2xl bg-emerald-600 text-white flex items-center justify-center shrink-0 shadow-sm">
                                    <i data-lucide="smartphone" class="w-6 h-6"></i>
                                </div>
                                <div>
                                    <h3 class="font-bold text-slate-900 text-sm">HP / Tablet Android</h3>
                                    <span class="text-[11px] font-semibold text-emerald-700">Android 8.0+ (Chrome / WebApp)</span>
                                </div>
                            </div>
                            <p class="text-xs text-slate-600 leading-relaxed">
                                Pasang aplikasi CBT langsung ke layar utama Android tanpa bilah alamat browser sehingga berjalan seperti aplikasi mandiri.
                            </p>
                        </div>
                        <button onclick="openInstallModal('android')" class="w-full py-2.5 px-4 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold transition flex items-center justify-center gap-2">
                            <i data-lucide="download-cloud" class="w-4 h-4"></i> Install / Panduan Android
                        </button>
                    </div>

                    <div class="rounded-2xl border border-slate-200 bg-slate-50/70 p-5 flex flex-col justify-between space-y-4">
                        <div class="space-y-3">
                            <div class="flex items-center gap-3">
                                <div class="w-11 h-11 rounded-2xl bg-slate-900 text-white flex items-center justify-center shrink-0 shadow-sm">
                                    <i data-lucide="tablet-smartphone" class="w-6 h-6"></i>
                                </div>
                                <div>
                                    <h3 class="font-bold text-slate-900 text-sm">iPhone &amp; iPad (iOS)</h3>
                                    <span class="text-[11px] font-semibold text-slate-600">iOS Safari WebClip / PWA</span>
                                </div>
                            </div>
                            <p class="text-xs text-slate-600 leading-relaxed">
                                Gunakan fitur <em>"Tambahkan ke Layar Utama"</em> di Safari iPhone/iPad untuk menjalankan CBT dalam mode layar penuh yang stabil.
                            </p>
                        </div>
                        <button onclick="openInstallModal('ios')" class="w-full py-2.5 px-4 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition flex items-center justify-center gap-2">
                            <i data-lucide="smartphone" class="w-4 h-4"></i> Install / Panduan iPhone (iOS)
                        </button>
                    </div>

                    <div class="rounded-2xl border border-blue-200 bg-blue-50/40 p-5 flex flex-col justify-between space-y-4">
                        <div class="space-y-3">
                            <div class="flex items-center gap-3">
                                <div class="w-11 h-11 rounded-2xl bg-blue-600 text-white flex items-center justify-center shrink-0 shadow-sm">
                                    <i data-lucide="monitor" class="w-6 h-6"></i>
                                </div>
                                <div>
                                    <h3 class="font-bold text-slate-900 text-sm">Komputer / Laptop (PC &amp; Mac)</h3>
                                    <span class="text-[11px] font-semibold text-blue-700">Windows / macOS / Chromebook</span>
                                </div>
                            </div>
                            <p class="text-xs text-slate-600 leading-relaxed">
                                Pasang sebagai aplikasi desktop mandiri lewat Chrome/Edge atau unduh file paket aplikasi offline untuk laboratorium komputer.
                            </p>
                        </div>
                        <button onclick="openInstallModal('desktop')" class="w-full py-2.5 px-4 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold transition flex items-center justify-center gap-2">
                            <i data-lucide="laptop" class="w-4 h-4"></i> Install / Download Komputer
                        </button>
                    </div>
                </div>
            </div>
        </section>

        <!-- ================================================================= -->
        <!-- VIEW B: DASHBOARD SISWA (PROFIL, PENUGASAN MAPEL, PANDUAN & TATIB) -->
        <!-- ================================================================= -->
        <section id="view-siswa-dashboard" class="hidden max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 space-y-6">
            <!-- Header Sapaan Siswa -->
            <div class="bg-gradient-to-r from-slate-900 via-blue-950 to-indigo-950 rounded-3xl p-6 text-white shadow-lg flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div class="flex items-center gap-4">
                    <div class="w-14 h-14 rounded-2xl bg-blue-600/30 border border-blue-400/40 flex items-center justify-center font-extrabold text-xl text-blue-200 shrink-0">
                        <i data-lucide="user-check" class="w-7 h-7"></i>
                    </div>
                    <div>
                        <span class="px-2.5 py-0.5 rounded-full bg-emerald-500/20 border border-emerald-400/30 text-emerald-300 text-[10px] font-bold uppercase tracking-wider">
                            Akun Peserta Terverifikasi
                        </span>
                        <h2 id="stuDashWelcomeName" class="text-xl sm:text-2xl font-extrabold mt-1">Selamat Datang, Peserta</h2>
                        <p id="stuDashWelcomeSub" class="text-xs text-blue-200 mt-0.5">Silakan pilih mata pelajaran yang sudah di-assign untuk memulai ujian.</p>
                    </div>
                </div>
                <div class="flex items-center gap-2 self-start md:self-center shrink-0">
                    <span class="px-3 py-1.5 rounded-xl bg-white/10 border border-white/15 text-xs font-semibold flex items-center gap-1.5">
                        <span class="w-2 h-2 rounded-full bg-emerald-400 animate-ping"></span> Sinkron Jadwal Live
                    </span>
                </div>
            </div>

            <!-- Navigasi Tab Dashboard Siswa -->
            <div class="flex flex-wrap gap-2 bg-white p-2 rounded-2xl border border-slate-200 shadow-sm">
                <button onclick="switchStudentDashTab('ujian')" id="stuTabBtn-ujian" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-bold bg-blue-600 text-white transition flex items-center gap-2">
                    <i data-lucide="book-open-check" class="w-4 h-4"></i> Penugasan &amp; Jadwal Ujian
                </button>
                <button onclick="switchStudentDashTab('profil')" id="stuTabBtn-profil" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition flex items-center gap-2">
                    <i data-lucide="id-card" class="w-4 h-4"></i> Profil &amp; Kartu Login Saya
                </button>
                <button onclick="switchStudentDashTab('panduan')" id="stuTabBtn-panduan" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition flex items-center gap-2">
                    <i data-lucide="shield-alert" class="w-4 h-4"></i> Panduan Menjawab &amp; Tata Tertib
                </button>
            </div>

            <!-- TAB 1 SISWA: PENUGASAN / UJIAN MATA PELAJARAN TERJADWAL -->
            <div id="stuDashTab-ujian" class="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
                <!-- Daftar Mata Pelajaran yang Di-Assign & Menunggu Assign -->
                <div class="lg:col-span-8 space-y-4">
                    <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                        <div>
                            <h3 class="font-extrabold text-slate-900 text-base">Daftar Ujian &amp; Penugasan Mata Pelajaran</h3>
                            <p class="text-xs text-slate-500">Pilih ujian yang berstatus <strong>Aktif / Siap Dikerjakan</strong>, lalu masukkan Token Ujian saat Anda siap.</p>
                        </div>
                        <span id="stuAssignedCountBadge" class="px-3 py-1 rounded-full bg-blue-50 text-blue-700 text-xs font-bold self-start shrink-0">0 Paket</span>
                    </div>

                    <div id="studentAssignmentsList" class="space-y-4"></div>
                </div>

                <!-- Ringkasan Profil Kartu Login & Pengingat Tata Tertib -->
                <div class="lg:col-span-4 space-y-4">
                    <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm space-y-3">
                        <div class="flex items-center justify-between border-b border-slate-100 pb-2.5">
                            <span class="text-xs font-extrabold uppercase tracking-wider text-blue-600">Ringkasan Kartu Peserta</span>
                            <button onclick="switchStudentDashTab('profil')" class="text-[11px] font-bold text-slate-500 hover:text-blue-600">Lihat Kartu &rarr;</button>
                        </div>
                        <div class="space-y-2 text-xs">
                            <div class="flex justify-between gap-2">
                                <span class="text-slate-500 shrink-0">Nama Lengkap:</span>
                                <span id="stuSideName" class="font-bold text-slate-900 text-right truncate">-</span>
                            </div>
                            <div class="flex justify-between gap-2">
                                <span class="text-slate-500 shrink-0">NISN / Username:</span>
                                <span id="stuSideNisn" class="font-mono font-bold text-slate-800">-</span>
                            </div>
                            <div class="flex justify-between gap-2">
                                <span class="text-slate-500 shrink-0">Kelas:</span>
                                <span id="stuSideClass" class="font-bold text-indigo-700">-</span>
                            </div>
                            <div class="flex justify-between gap-2">
                                <span class="text-slate-500 shrink-0">Ruangan Ujian:</span>
                                <span id="stuSideRoom" class="font-bold text-slate-900">-</span>
                            </div>
                        </div>
                    </div>

                    <div class="bg-amber-50 border border-amber-200 rounded-2xl p-5 space-y-2.5 text-xs text-amber-950">
                        <div class="font-extrabold text-amber-900 flex items-center gap-1.5 text-sm">
                            <i data-lucide="siren" class="w-4 h-4 text-red-600 shrink-0"></i> Peringatan Keamanan Ruang Ujian
                        </div>
                        <p class="leading-relaxed text-amber-900">
                            Setelah Anda menekan tombol <strong>Mulai Ujian</strong> dan memasukkan Token:
                        </p>
                        <ul class="list-disc pl-4 space-y-1 text-amber-900">
                            <li>Aplikasi akan mengunci mode layar penuh.</li>
                            <li>Dilarang berpindah tab, membuka WhatsApp/Google, atau membagi layar.</li>
                            <li>Jika keluar sistem, <strong>layar akan terkunci otomatis &amp; sirine alarm berbunyi keras</strong> serta tercatat di Pengawas &amp; Admin.</li>
                        </ul>
                        <button onclick="switchStudentDashTab('panduan')" class="mt-2 inline-flex items-center gap-1 text-xs font-bold text-red-700 hover:underline">
                            Baca Selengkapnya Tata Tertib &amp; Panduan &rarr;
                        </button>
                    </div>
                </div>
            </div>

            <!-- TAB 2 SISWA: PROFIL LENGKAP SESUAI KARTU LOGIN -->
            <div id="stuDashTab-profil" class="hidden grid grid-cols-1 lg:grid-cols-12 gap-6">
                <div class="lg:col-span-6 bg-white rounded-3xl border-2 border-slate-300 p-6 shadow-sm space-y-4">
                    <div class="flex items-center justify-between pb-3 border-b border-slate-200 gap-3">
                        <div class="flex items-center gap-3 min-w-0">
                            <div class="w-11 h-11 rounded-xl bg-slate-900 text-white flex items-center justify-center font-extrabold text-sm shrink-0">CBT</div>
                            <div class="min-w-0">
                                <div class="text-xs font-extrabold uppercase tracking-wider text-slate-900">KARTU IDENTITAS PESERTA UJIAN</div>
                                <div id="stuCardSchoolName" class="text-[11px] text-slate-500 truncate">CBT-PRO | Sistem Ujian Berbasis Komputer</div>
                            </div>
                        </div>
                        <span class="px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-700 text-[10px] font-bold shrink-0">AKTIF</span>
                    </div>

                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs pt-1">
                        <div class="p-3 rounded-xl bg-slate-50 border border-slate-200">
                            <span class="text-[10px] uppercase font-bold text-slate-400 block">Nama Lengkap Peserta</span>
                            <span id="stuProfName" class="text-sm font-extrabold text-slate-900 mt-0.5 block">-</span>
                        </div>
                        <div class="p-3 rounded-xl bg-slate-50 border border-slate-200">
                            <span class="text-[10px] uppercase font-bold text-slate-400 block">NISN</span>
                            <span id="stuProfNisn" class="text-sm font-mono font-extrabold text-blue-700 mt-0.5 block">-</span>
                        </div>
                        <div class="p-3 rounded-xl bg-slate-50 border border-slate-200">
                            <span class="text-[10px] uppercase font-bold text-slate-400 block">Kelas / Jurusan</span>
                            <span id="stuProfClass" class="text-sm font-bold text-slate-900 mt-0.5 block">-</span>
                        </div>
                        <div class="p-3 rounded-xl bg-slate-50 border border-slate-200">
                            <span class="text-[10px] uppercase font-bold text-slate-400 block">Ruangan Ujian</span>
                            <span id="stuProfRoom" class="text-sm font-bold text-indigo-700 mt-0.5 block">-</span>
                        </div>
                        <div class="p-3 rounded-xl bg-blue-50/60 border border-blue-200">
                            <span class="text-[10px] uppercase font-bold text-blue-500 block">Username Kartu Login</span>
                            <span id="stuProfUser" class="text-sm font-mono font-extrabold text-slate-900 mt-0.5 block">-</span>
                        </div>
                        <div class="p-3 rounded-xl bg-blue-50/60 border border-blue-200">
                            <span class="text-[10px] uppercase font-bold text-blue-500 block">Status Sesi Login</span>
                            <span id="stuProfLoginTime" class="text-xs font-bold text-emerald-700 mt-0.5 block">Terhubung Real-Time</span>
                        </div>
                    </div>
                </div>

                <div class="lg:col-span-6 bg-white rounded-3xl border border-slate-200 p-6 shadow-sm space-y-4">
                    <h3 class="font-extrabold text-slate-900 text-base flex items-center gap-2">
                        <i data-lucide="history" class="w-4 h-4 text-blue-600"></i> Riwayat Pengerjaan Ujian Saya
                    </h3>
                    <div id="stuMyHistoryList" class="space-y-2.5 text-xs"></div>
                </div>
            </div>

            <!-- TAB 3 SISWA: PANDUAN MENJAWAB SOAL & TATA TERTIB RUANG UJIAN -->
            <div id="stuDashTab-panduan" class="hidden grid grid-cols-1 lg:grid-cols-2 gap-6">
                <div class="bg-white rounded-3xl border border-slate-200 p-6 shadow-sm space-y-4">
                    <div class="flex items-center gap-3 pb-3 border-b border-slate-100">
                        <div class="w-10 h-10 rounded-xl bg-blue-100 text-blue-600 flex items-center justify-center shrink-0">
                            <i data-lucide="help-circle" class="w-5 h-5"></i>
                        </div>
                        <div>
                            <h3 class="font-extrabold text-slate-900 text-base">Panduan Cara Menjawab Soal CBT</h3>
                            <span class="text-xs text-slate-500">Langkah teknis pengerjaan soal pada sistem</span>
                        </div>
                    </div>
                    <ol class="space-y-3 text-xs sm:text-sm text-slate-700 list-decimal pl-5 leading-relaxed">
                        <li><strong>Pilih Mata Pelajaran Terjadwal:</strong> Pada tab <em>Penugasan &amp; Jadwal Ujian</em>, pastikan status ujian sudah <strong>Siap Dikerjakan (Assigned)</strong>, lalu klik tombol <strong>Mulai Ujian</strong>.</li>
                        <li><strong>Masukkan Token Ujian:</strong> Ketikkan kode <strong>Token Ujian</strong> yang diberikan oleh Pengawas Ruangan pada layar konfirmasi token.</li>
                        <li><strong>Memilih Jawaban:</strong> Klik langsung pada pilihan jawaban <strong>A, B, C, D, atau E</strong>. Jawaban tersimpan otomatis pada sistem.</li>
                        <li><strong>Navigasi Nomor &amp; Ragu-Ragu:</strong> Gunakan panel nomor soal di sebelah kanan untuk melompat ke nomor tertentu. Gunakan tombol <em>Tandai Ragu-Ragu</em> (warna kuning) jika ingin meninjau ulang soal tersebut.</li>
                        <li><strong>Pengaturan Ukuran Teks:</strong> Gunakan tombol <strong>A- / A / A+</strong> di bagian atas layar ujian untuk menyesuaikan ukuran huruf soal.</li>
                        <li><strong>Mengakhiri Ujian:</strong> Setelah seluruh soal terjawab dan diperiksa, klik tombol <strong>Selesai Ujian</strong> untuk mengirim hasil secara permanen.</li>
                    </ol>
                </div>

                <div class="bg-white rounded-3xl border border-red-200 p-6 shadow-sm space-y-4">
                    <div class="flex items-center gap-3 pb-3 border-b border-red-100">
                        <div class="w-10 h-10 rounded-xl bg-red-100 text-red-600 flex items-center justify-center shrink-0">
                            <i data-lucide="shield-alert" class="w-5 h-5"></i>
                        </div>
                        <div>
                            <h3 class="font-extrabold text-red-900 text-base">Tata Tertib Peserta di Ruangan Ujian</h3>
                            <span class="text-xs text-red-600 font-semibold">Wajib Dipatuhi Seluruh Peserta Ujian</span>
                        </div>
                    </div>
                    <ul class="space-y-3 text-xs sm:text-sm text-slate-700 list-disc pl-5 leading-relaxed">
                        <li>Peserta wajib hadir di ruangan ujian minimal 15 menit sebelum jadwal dimulai dengan membawa <strong>Kartu Login Peserta</strong>.</li>
                        <li>Peserta dilarang keras membuka tab browser lain, aplikasi pesan, kalkulator eksternal, catatan digital, atau melakukan <em>split-screen</em> selama ujian berlangsung.</li>
                        <li>Peserta dilarang keluar dari mode layar penuh (<em>Fullscreen</em>), menekan tombol kembali (<em>Back</em>), atau mengambil tangkapan layar (<em>Screenshot</em>).</li>
                        <li><strong>Sanksi Sistem Otomatis:</strong> Setiap upaya keluar dari halaman ujian akan langsung <strong>mengunci layar ujian, membunyikan sirine alarm peringatan</strong>, serta mencatat pelanggaran secara real-time di layar Pengawas dan Administrator.</li>
                        <li>Pembukaan kunci layar hanya dapat dilakukan oleh <strong>Pengawas Ruangan</strong> setelah memeriksa perangkat peserta.</li>
                        <li>Kerjakan soal secara mandiri, jujur, dan tenang hingga waktu berakhir.</li>
                    </ul>
                </div>
            </div>
        </section>

        <!-- ================================================================= -->
        <!-- VIEW C: DASHBOARD ADMINISTRATOR (7 MENU SIDEBAR TERPISAH)         -->
        <!-- ================================================================= -->
        <section id="view-admin" class="hidden max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
            <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
                
                <!-- SIDEBAR NAVIGASI -->
                <aside class="lg:col-span-3 bg-white rounded-2xl border border-slate-200 p-4 shadow-sm space-y-2 lg:sticky lg:top-20 no-print">
                    <div class="px-3 py-2 border-b border-slate-100 mb-2">
                        <span class="text-[10px] font-extrabold uppercase tracking-wider text-indigo-600">Panel Kendali Utama</span>
                        <h2 class="font-extrabold text-slate-900 text-base">Menu Administrator</h2>
                    </div>

                    <nav class="flex lg:flex-col overflow-x-auto lg:overflow-visible gap-1.5 pb-1 lg:pb-0">
                        <button onclick="switchAdminMenu('monitoring')" id="adminNav-monitoring" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-bold bg-indigo-600 text-white transition whitespace-nowrap">
                            <i data-lucide="activity" class="w-4 h-4 shrink-0"></i>
                            <span>1. Monitoring Siswa</span>
                        </button>
                        <button onclick="switchAdminMenu('soal')" id="adminNav-soal" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="file-text" class="w-4 h-4 shrink-0"></i>
                            <span>2. Bank Soal (Per Mapel)</span>
                        </button>
                        <button onclick="switchAdminMenu('assign')" id="adminNav-assign" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="calendar-check" class="w-4 h-4 shrink-0"></i>
                            <span>3. Assignment &amp; Jadwal</span>
                        </button>
                        <button onclick="switchAdminMenu('logs')" id="adminNav-logs" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="scroll-text" class="w-4 h-4 shrink-0"></i>
                            <span>4. Log Masuk/Keluar &amp; Pelanggaran</span>
                        </button>
                        <button onclick="switchAdminMenu('nilai')" id="adminNav-nilai" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="file-spreadsheet" class="w-4 h-4 shrink-0"></i>
                            <span>5. Data &amp; Download Nilai</span>
                        </button>
                        <button onclick="switchAdminMenu('kartu')" id="adminNav-kartu" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="id-card" class="w-4 h-4 shrink-0"></i>
                            <span>6. Akun Siswa &amp; Kartu</span>
                        </button>
                        <button onclick="switchAdminMenu('setelan')" id="adminNav-setelan" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="settings-2" class="w-4 h-4 shrink-0"></i>
                            <span>7. Setelan Keseluruhan</span>
                        </button>
                    </nav>

                    <!-- Info Token Aktif di Sidebar -->
                    <div class="hidden lg:block mt-4 pt-4 border-t border-slate-100">
                        <div class="bg-slate-900 text-white rounded-xl p-3.5">
                            <div class="flex items-center justify-between text-[10px] text-slate-400 uppercase font-bold">
                                <span>Token Ujian Aktif</span>
                                <span class="text-emerald-400">SINKRON</span>
                            </div>
                            <div class="flex items-center justify-between mt-1">
                                <span id="sidebarTokenDisplay" class="font-mono font-extrabold text-xl text-amber-400 tracking-widest">------</span>
                                <button onclick="copyTextToClipboard(document.getElementById('sidebarTokenDisplay').innerText, 'Token berhasil disalin!')" class="p-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-white" title="Salin Token">
                                    <i data-lucide="copy" class="w-3.5 h-3.5"></i>
                                </button>
                            </div>
                        </div>
                    </div>
                </aside>

                <!-- KONTEN KANAN ADMINISTRATOR -->
                <div class="lg:col-span-9 space-y-6">

                    <!-- MENU 1: MONITORING SISWA (REAL-TIME) -->
                    <div id="adminSection-monitoring" class="space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                            <div>
                                <h2 class="text-xl font-extrabold text-slate-900">Monitoring Peserta Ujian Real-Time</h2>
                                <p class="text-xs text-slate-500 mt-0.5">Pantau status siswa di Dashboard maupun ruang ujian, matikan sirine/buka kunci layar, atau reset sesi.</p>
                            </div>
                            <span class="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-emerald-100 text-emerald-700 text-xs font-bold self-start shrink-0">
                                <span class="w-2 h-2 rounded-full bg-emerald-500 animate-ping"></span> Sinkronisasi Real-Time
                            </span>
                        </div>

                        <div class="grid grid-cols-2 sm:grid-cols-4 gap-4">
                            <div class="bg-white rounded-2xl border border-slate-200 p-4 shadow-sm">
                                <span class="text-xs text-slate-500 font-medium">Total Peserta</span>
                                <div id="monStatTotal" class="text-2xl font-extrabold text-slate-900 mt-1">0</div>
                            </div>
                            <div class="bg-white rounded-2xl border border-blue-200 p-4 shadow-sm">
                                <span class="text-xs text-blue-600 font-semibold">Sedang Ujian / Online</span>
                                <div id="monStatActive" class="text-2xl font-extrabold text-blue-700 mt-1">0</div>
                            </div>
                            <div class="bg-white rounded-2xl border border-emerald-200 p-4 shadow-sm">
                                <span class="text-xs text-emerald-600 font-semibold">Selesai Ujian</span>
                                <div id="monStatDone" class="text-2xl font-extrabold text-emerald-700 mt-1">0</div>
                            </div>
                            <div class="bg-white rounded-2xl border border-red-200 p-4 shadow-sm">
                                <span class="text-xs text-red-600 font-semibold">Terkunci / Sirine</span>
                                <div id="monStatLocked" class="text-2xl font-extrabold text-red-600 mt-1">0</div>
                            </div>
                        </div>

                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                            <div class="flex flex-col sm:flex-row gap-3 justify-between">
                                <input type="text" id="monSearchInput" oninput="renderAdminMonitoring()" placeholder="Cari nama atau NISN siswa..." class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm w-full sm:w-64 focus:ring-2 focus:ring-indigo-500 focus:outline-none">
                                <select id="monRoomFilter" onchange="renderAdminMonitoring()" class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm bg-white">
                                    <option value="">Semua Kelas / Ruangan</option>
                                </select>
                            </div>

                            <div class="overflow-x-auto border border-slate-200 rounded-xl">
                                <table class="w-full text-left border-collapse text-xs sm:text-sm">
                                    <thead>
                                        <tr class="bg-slate-50 border-b border-slate-200 text-[11px] uppercase text-slate-600">
                                            <th class="py-3 px-4 font-bold">Peserta &amp; NISN</th>
                                            <th class="py-3 px-4 font-bold">Kelas / Ruang</th>
                                            <th class="py-3 px-4 font-bold">Mapel Aktif</th>
                                            <th class="py-3 px-4 font-bold">Status Live</th>
                                            <th class="py-3 px-4 font-bold">Progres</th>
                                            <th class="py-3 px-4 font-bold">Pelanggaran</th>
                                            <th class="py-3 px-4 font-bold text-right">Aksi Cepat</th>
                                        </tr>
                                    </thead>
                                    <tbody id="adminMonitoringTbody" class="divide-y divide-slate-100"></tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <!-- MENU 2: BANK SOAL PER MATA PELAJARAN / PAKET -->
                    <div id="adminSection-soal" class="hidden space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                            <div>
                                <h2 class="text-xl font-extrabold text-slate-900">Bank Soal &amp; Paket Mata Pelajaran</h2>
                                <p class="text-xs text-slate-500 mt-0.5">Kelola paket soal per mata pelajaran (Upload Word/PDF atau Input Manual), lalu tugaskan di menu <strong>Assignment &amp; Jadwal</strong>.</p>
                            </div>
                        </div>

                        <!-- Kelola / Pilih Paket Mata Pelajaran -->
                        <div class="bg-white rounded-2xl border border-indigo-200 p-5 shadow-sm space-y-4">
                            <div class="flex flex-col md:flex-row md:items-end justify-between gap-4">
                                <div class="flex-1">
                                    <label class="block text-xs font-extrabold uppercase tracking-wider text-indigo-700 mb-1.5">Pilih Paket Soal Mata Pelajaran yang Dikelola:</label>
                                    <select id="adminPackageSelector" onchange="selectAdminQuestionPackage(this.value)" class="w-full px-4 py-2.5 rounded-xl border-2 border-indigo-300 bg-indigo-50/40 text-sm font-bold text-slate-900"></select>
                                </div>
                                <form onsubmit="createNewSubjectPackage(event)" class="flex flex-wrap sm:flex-nowrap gap-2">
                                    <input type="text" id="newPkgSubjectName" required placeholder="Nama Mapel Baru (Misal: Matematika)" class="px-3.5 py-2.5 rounded-xl border border-slate-300 text-xs sm:text-sm bg-white">
                                    <button type="submit" class="px-4 py-2.5 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold whitespace-nowrap transition">
                                        + Buat Paket Mapel
                                    </button>
                                </form>
                            </div>
                        </div>

                        <div class="grid grid-cols-1 xl:grid-cols-12 gap-6">
                            <!-- Kolom Input Soal -->
                            <div class="xl:col-span-5 space-y-6">
                                <!-- Upload Word (.docx) & PDF (.pdf) -->
                                <div class="bg-white rounded-2xl border border-indigo-200 p-5 shadow-sm space-y-3">
                                    <div class="flex items-center justify-between gap-2">
                                        <h3 class="font-bold text-slate-900 text-sm flex items-center gap-2">
                                            <i data-lucide="file-up" class="w-4 h-4 text-indigo-600 shrink-0"></i> Upload Soal Word (.docx) / PDF (.pdf)
                                        </h3>
                                        <button onclick="downloadSampleQuestionFormat()" class="text-[11px] font-bold text-indigo-600 hover:underline shrink-0">
                                            Unduh Contoh Format
                                        </button>
                                    </div>
                                    <p class="text-[11px] text-slate-500 leading-relaxed">
                                        Soal akan masuk ke paket mata pelajaran yang dipilih di atas.<br>
                                        <code class="block bg-slate-50 p-2 rounded-lg border border-slate-200 text-slate-700 mt-1 overflow-x-auto">
                                            1. Pertanyaan soal di sini?<br>
                                            A. Pilihan A &nbsp; B. Pilihan B<br>
                                            C. Pilihan C &nbsp; D. Pilihan D &nbsp; E. Pilihan E<br>
                                            Kunci: B
                                        </code>
                                    </p>
                                    <label class="flex flex-col items-center justify-center w-full h-28 border-2 border-dashed border-indigo-300 rounded-2xl bg-indigo-50/40 hover:bg-indigo-50 cursor-pointer transition p-4 text-center">
                                        <i data-lucide="upload-cloud" class="w-7 h-7 text-indigo-600 mb-1"></i>
                                        <span class="text-xs font-bold text-indigo-900">Klik untuk Pilih File .DOCX, .PDF, atau .TXT</span>
                                        <span class="text-[10px] text-slate-500 mt-0.5">Otomatis mengekstrak soal, opsi A-E, dan kunci jawaban</span>
                                        <input type="file" accept=".docx,.pdf,.txt,.csv" onchange="handleDocumentQuestionUpload(event, 'admin')" class="hidden">
                                    </label>
                                    <div id="uploadParseStatus" class="hidden text-xs p-2.5 rounded-xl bg-blue-50 text-blue-700 font-medium"></div>
                                </div>

                                <!-- Form Input Soal Manual -->
                                <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
                                    <h3 class="font-bold text-slate-900 text-sm mb-3">Tambah Soal Manual ke Paket Ini</h3>
                                    <form onsubmit="handleAddManualQuestion(event, 'admin')" class="space-y-3">
                                        <div>
                                            <label class="block text-xs font-semibold text-slate-700 mb-1">Teks Pertanyaan *</label>
                                            <textarea id="mQText" required rows="2" placeholder="Tuliskan pertanyaan..." class="w-full p-2.5 rounded-xl border border-slate-300 text-xs sm:text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none"></textarea>
                                        </div>
                                        <div class="space-y-2">
                                            <div class="flex items-center gap-2">
                                                <span class="w-6 text-xs font-bold text-slate-500 shrink-0">A.</span>
                                                <input type="text" id="mQOptA" required placeholder="Opsi A" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                            </div>
                                            <div class="flex items-center gap-2">
                                                <span class="w-6 text-xs font-bold text-slate-500 shrink-0">B.</span>
                                                <input type="text" id="mQOptB" required placeholder="Opsi B" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                            </div>
                                            <div class="flex items-center gap-2">
                                                <span class="w-6 text-xs font-bold text-slate-500 shrink-0">C.</span>
                                                <input type="text" id="mQOptC" required placeholder="Opsi C" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                            </div>
                                            <div class="flex items-center gap-2">
                                                <span class="w-6 text-xs font-bold text-slate-500 shrink-0">D.</span>
                                                <input type="text" id="mQOptD" required placeholder="Opsi D" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                            </div>
                                            <div class="flex items-center gap-2">
                                                <span class="w-6 text-xs font-bold text-slate-500 shrink-0">E.</span>
                                                <input type="text" id="mQOptE" placeholder="Opsi E (Opsional)" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                            </div>
                                        </div>
                                        <div class="grid grid-cols-2 gap-3 pt-1">
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Kunci Jawaban</label>
                                                <select id="mQAnswer" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-bold text-indigo-700">
                                                    <option value="A">Opsi A</option>
                                                    <option value="B">Opsi B</option>
                                                    <option value="C">Opsi C</option>
                                                    <option value="D">Opsi D</option>
                                                    <option value="E">Opsi E</option>
                                                </select>
                                            </div>
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Bobot Poin</label>
                                                <input type="number" id="mQPoints" value="10" min="1" max="100" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-bold">
                                            </div>
                                        </div>
                                        <button type="submit" class="w-full py-2.5 px-4 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold transition shadow-sm">
                                            + Simpan Soal ke Paket
                                        </button>
                                    </form>
                                </div>
                            </div>

                            <!-- Daftar & Edit Soal -->
                            <div class="xl:col-span-7 bg-white rounded-2xl border border-slate-200 p-5 shadow-sm flex flex-col">
                                <div class="flex items-center justify-between gap-2 mb-4 pb-3 border-b border-slate-100">
                                    <div>
                                        <h3 class="font-bold text-slate-900 text-sm sm:text-base">Daftar Soal Paket Terpilih (<span id="adminBankSoalCount">0</span>)</h3>
                                        <p class="text-[11px] text-slate-500">Soal yang sudah siap di sini dapat langsung di-assign di Menu 3</p>
                                    </div>
                                    <button onclick="confirmClearPackageQuestions()" class="px-3 py-1.5 rounded-lg bg-red-50 hover:bg-red-100 text-red-600 text-xs font-bold transition shrink-0">
                                        Kosongkan Soal Paket
                                    </button>
                                </div>
                                <div id="adminSoalListContainer" class="space-y-3 overflow-y-auto max-h-[680px] pr-1"></div>
                            </div>
                        </div>
                    </div>

                    <!-- MENU 3: ASSIGNMENT MATA PELAJARAN, JADWAL & TOKEN -->
                    <div id="adminSection-assign" class="hidden space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm">
                            <h2 class="text-xl font-extrabold text-slate-900">Assignment Mata Pelajaran, Jadwal Ujian &amp; Token</h2>
                            <p class="text-xs text-slate-500 mt-0.5">Pilih soal mata pelajaran yang telah disiapkan di Bank Soal, atur jadwal pelaksanaan &amp; status Assignment (Aktif / Menunggu Assign), serta kelola Token sinkron.</p>
                        </div>

                        <div class="grid grid-cols-1 lg:grid-cols-12 gap-6">
                            <!-- Manajemen Token Ujian (Sinkron Pengawas & Siswa) -->
                            <div class="lg:col-span-5 bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-5 self-start">
                                <div>
                                    <span class="px-2.5 py-1 rounded-full bg-amber-100 text-amber-800 text-[11px] font-bold">Sinkron Otomatis ke Pengawas &amp; Siswa</span>
                                    <h3 class="font-bold text-slate-900 text-base mt-2">Token Akses Ujian Aktif</h3>
                                    <p class="text-xs text-slate-500 mt-0.5">Siswa memasukkan token ini di Dashboard Siswa saat menekan tombol Mulai Ujian pada mapel yang sudah di-assign.</p>
                                </div>

                                <div class="bg-slate-900 text-white p-5 rounded-2xl flex items-center justify-between gap-3">
                                    <div>
                                        <span class="text-[10px] uppercase text-slate-400 font-bold block">Token Saat Ini</span>
                                        <span id="assignTokenDisplay" class="text-3xl font-mono font-extrabold text-amber-400 tracking-widest">CBT26X</span>
                                    </div>
                                    <button onclick="copyTextToClipboard(document.getElementById('assignTokenDisplay').innerText, 'Token berhasil disalin!')" class="p-3 rounded-xl bg-white/10 hover:bg-white/20 text-white transition shrink-0" title="Salin Token">
                                        <i data-lucide="copy" class="w-5 h-5"></i>
                                    </button>
                                </div>

                                <div class="space-y-2 pt-2 border-t border-slate-100">
                                    <label class="block text-xs font-bold text-slate-700">Ubah Token Secara Manual</label>
                                    <div class="flex gap-2">
                                        <input type="text" id="customTokenInput" maxlength="10" placeholder="Misal: UAS2026" class="flex-1 px-3.5 py-2.5 rounded-xl border border-slate-300 font-mono font-bold uppercase text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none">
                                        <button onclick="saveManualToken()" class="px-4 py-2.5 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition shrink-0">
                                            Simpan
                                        </button>
                                    </div>
                                </div>

                                <button onclick="generateRandomToken()" class="w-full py-3 px-4 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold transition flex items-center justify-center gap-2 shadow-sm">
                                    <i data-lucide="refresh-cw" class="w-4 h-4"></i> Acak / Generate Token Baru Otomatis
                                </button>
                            </div>

                            <!-- Form Penugasan (Assign) Paket Soal Mata Pelajaran -->
                            <div class="lg:col-span-7 bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                                <h3 class="font-bold text-slate-900 text-base flex items-center gap-2">
                                    <i data-lucide="calendar-plus" class="w-4 h-4 text-indigo-600"></i> Jadwalkan &amp; Assign Paket Soal ke Siswa
                                </h3>
                                <form onsubmit="savePackageAssignment(event)" class="space-y-4">
                                    <div>
                                        <label class="block text-xs font-bold text-slate-700 mb-1">Pilih Paket Soal dari Bank Soal *</label>
                                        <select id="assignPackageSelect" onchange="loadSelectedPackageToAssignForm()" required class="w-full px-3.5 py-2.5 rounded-xl border-2 border-indigo-200 bg-indigo-50/30 text-sm font-bold text-slate-900"></select>
                                    </div>

                                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Judul Sesi Ujian *</label>
                                            <input type="text" id="assignExamTitle" required placeholder="Misal: UAS Ganjil 2026" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 text-sm">
                                        </div>
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Mata Pelajaran *</label>
                                            <input type="text" id="assignSubject" required placeholder="Misal: Akuntansi Dasar" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 text-sm">
                                        </div>
                                    </div>

                                    <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Tanggal Jadwal *</label>
                                            <input type="date" id="assignScheduleDate" required class="w-full px-3 py-2.5 rounded-xl border border-slate-300 text-xs sm:text-sm font-semibold">
                                        </div>
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Jam Mulai - Selesai *</label>
                                            <input type="text" id="assignScheduleTime" required placeholder="08:00 - 09:30 WITA" class="w-full px-3 py-2.5 rounded-xl border border-slate-300 text-xs sm:text-sm font-semibold">
                                        </div>
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Durasi (Menit) *</label>
                                            <input type="number" id="assignDuration" min="1" max="360" required value="45" class="w-full px-3 py-2.5 rounded-xl border border-slate-300 text-sm font-bold">
                                        </div>
                                    </div>

                                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Target Kelas / Ruangan</label>
                                            <select id="assignTargetGroup" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 text-sm bg-white font-semibold">
                                                <option value="ALL">Semua Kelas &amp; Ruangan</option>
                                            </select>
                                        </div>
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Status Penugasan (Assign)</label>
                                            <select id="assignStatusSelect" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 text-sm bg-white font-bold text-emerald-700">
                                                <option value="ASSIGNED">DI-ASSIGN (Siap Dikerjakan Siswa)</option>
                                                <option value="WAITING">TERJADWAL (Menunggu Assign)</option>
                                            </select>
                                        </div>
                                    </div>

                                    <div class="p-4 rounded-2xl bg-slate-50 border border-slate-200 grid grid-cols-1 sm:grid-cols-2 gap-3">
                                        <label class="flex items-center gap-2.5 cursor-pointer">
                                            <input type="checkbox" id="assignShuffleQuestions" class="w-4 h-4 rounded text-indigo-600">
                                            <span class="text-xs font-bold text-slate-800">Acak Urutan Soal</span>
                                        </label>
                                        <label class="flex items-center gap-2.5 cursor-pointer">
                                            <input type="checkbox" id="assignShuffleOptions" class="w-4 h-4 rounded text-indigo-600">
                                            <span class="text-xs font-bold text-slate-800">Acak Pilihan Jawaban</span>
                                        </label>
                                    </div>

                                    <div class="flex justify-end pt-1">
                                        <button type="submit" class="px-6 py-3 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs sm:text-sm font-bold transition shadow-sm flex items-center gap-2">
                                            <i data-lucide="check-circle" class="w-4 h-4"></i> Simpan &amp; Sinkronkan Penugasan
                                        </button>
                                    </div>
                                </form>
                            </div>
                        </div>

                        <!-- Tabel Daftar Penugasan Mata Pelajaran -->
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                            <div>
                                <h3 class="font-bold text-slate-900 text-base">Daftar Jadwal &amp; Status Assignment Mata Pelajaran</h3>
                                <p class="text-xs text-slate-500">Ubah status antara <strong>Siap Dikerjakan (Assigned)</strong> atau <strong>Menunggu Assign</strong> kapan saja sesuai jadwal.</p>
                            </div>
                            <div class="overflow-x-auto border border-slate-200 rounded-xl">
                                <table class="w-full text-left border-collapse text-xs sm:text-sm">
                                    <thead>
                                        <tr class="bg-slate-50 border-b border-slate-200 text-[11px] uppercase text-slate-600">
                                            <th class="py-3 px-4 font-bold">Mata Pelajaran &amp; Paket</th>
                                            <th class="py-3 px-4 font-bold">Jadwal Pelaksanaan</th>
                                            <th class="py-3 px-4 font-bold">Target</th>
                                            <th class="py-3 px-4 font-bold">Jumlah Soal</th>
                                            <th class="py-3 px-4 font-bold">Status</th>
                                            <th class="py-3 px-4 font-bold text-right">Aksi Assignment</th>
                                        </tr>
                                    </thead>
                                    <tbody id="adminAssignmentsTableBody" class="divide-y divide-slate-100"></tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <!-- MENU 4: LOG MASUK/KELUAR SISWA & PENGAWAS -->
                    <div id="adminSection-logs" class="hidden space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                            <div>
                                <h2 class="text-xl font-extrabold text-slate-900">Audit Log Aktivitas Masuk/Keluar &amp; Pelanggaran</h2>
                                <p class="text-xs text-slate-500 mt-0.5">Rekam jejak real-time seluruh aktivitas login, logout, mulai ujian, selesai ujian, dan pelanggaran keluar sistem oleh Siswa &amp; Pengawas.</p>
                            </div>
                            <div class="flex items-center gap-2 shrink-0">
                                <button onclick="exportActivityLogsExcel()" class="px-4 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold transition flex items-center gap-1.5">
                                    <i data-lucide="download" class="w-3.5 h-3.5"></i> Export Log (.XLSX)
                                </button>
                                <button onclick="confirmClearActivityLogs()" class="px-3.5 py-2.5 rounded-xl bg-red-50 hover:bg-red-100 text-red-600 text-xs font-bold transition">
                                    Bersihkan Log
                                </button>
                            </div>
                        </div>

                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                            <div class="flex flex-col sm:flex-row gap-3 justify-between">
                                <input type="text" id="logSearchInput" oninput="renderAdminLogsTable()" placeholder="Cari nama pengguna, NISN, atau aktivitas..." class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm w-full sm:w-72 focus:ring-2 focus:ring-indigo-500 focus:outline-none">
                                <div class="flex flex-wrap gap-2">
                                    <select id="logRoleFilter" onchange="renderAdminLogsTable()" class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm bg-white">
                                        <option value="">Semua Peran (Siswa &amp; Pengawas)</option>
                                        <option value="siswa">Khusus Siswa</option>
                                        <option value="pengawas">Khusus Pengawas</option>
                                        <option value="admin">Administrator</option>
                                    </select>
                                    <select id="logTypeFilter" onchange="renderAdminLogsTable()" class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm bg-white">
                                        <option value="">Semua Jenis Log</option>
                                        <option value="LOGIN">Login / Masuk Sistem</option>
                                        <option value="LOGOUT">Logout / Keluar Sistem</option>
                                        <option value="EXAM_START">Mulai Ujian</option>
                                        <option value="EXAM_FINISH">Selesai Ujian</option>
                                        <option value="VIOLATION">Pelanggaran / Sirine</option>
                                        <option value="UNLOCK">Buka Kunci Layar</option>
                                    </select>
                                </div>
                            </div>

                            <div class="overflow-x-auto border border-slate-200 rounded-xl max-h-[540px] overflow-y-auto">
                                <table class="w-full text-left border-collapse text-xs sm:text-sm">
                                    <thead>
                                        <tr class="bg-slate-50 border-b border-slate-200 text-[11px] uppercase text-slate-600 sticky top-0">
                                            <th class="py-3 px-4 font-bold">Waktu</th>
                                            <th class="py-3 px-4 font-bold">Peran</th>
                                            <th class="py-3 px-4 font-bold">Nama Pengguna &amp; ID</th>
                                            <th class="py-3 px-4 font-bold">Kategori</th>
                                            <th class="py-3 px-4 font-bold">Keterangan Aktivitas</th>
                                        </tr>
                                    </thead>
                                    <tbody id="adminLogsTableBody" class="divide-y divide-slate-100"></tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <!-- MENU 5: DATA SISWA & DOWNLOAD NILAI (.XLSX) -->
                    <div id="adminSection-nilai" class="hidden space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                            <div>
                                <h2 class="text-xl font-extrabold text-slate-900">Data Peserta &amp; Rekap Nilai Ujian</h2>
                                <p class="text-xs text-slate-500 mt-0.5">Nilai masuk secara real-time dan dapat diunduh langsung dalam format Microsoft Excel (.xlsx).</p>
                            </div>
                            <button onclick="downloadExcelResults()" class="px-5 py-3 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs sm:text-sm font-bold transition flex items-center gap-2 shadow-sm self-start shrink-0">
                                <i data-lucide="file-spreadsheet" class="w-4 h-4"></i> Download Nilai Excel (.XLSX)
                            </button>
                        </div>

                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
                                <input type="text" id="nilaiSearchInput" oninput="renderAdminNilaiTable()" placeholder="Cari Nama atau NISN..." class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm">
                                <select id="nilaiClassFilter" onchange="renderAdminNilaiTable()" class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm bg-white">
                                    <option value="">Semua Kelas</option>
                                </select>
                                <select id="nilaiRoomFilter" onchange="renderAdminNilaiTable()" class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm bg-white">
                                    <option value="">Semua Ruangan</option>
                                </select>
                            </div>

                            <div class="overflow-x-auto border border-slate-200 rounded-xl">
                                <table class="w-full text-left border-collapse text-xs sm:text-sm">
                                    <thead>
                                        <tr class="bg-slate-50 border-b border-slate-200 text-[11px] uppercase text-slate-600">
                                            <th class="py-3 px-3 font-bold">No</th>
                                            <th class="py-3 px-3 font-bold">Nama Peserta</th>
                                            <th class="py-3 px-3 font-bold">NISN</th>
                                            <th class="py-3 px-3 font-bold">Kelas / Ruang</th>
                                            <th class="py-3 px-3 font-bold">Mata Pelajaran</th>
                                            <th class="py-3 px-3 font-bold">Status</th>
                                            <th class="py-3 px-3 font-bold">Benar</th>
                                            <th class="py-3 px-3 font-bold">Salah</th>
                                            <th class="py-3 px-3 font-bold">Nilai Akhir</th>
                                            <th class="py-3 px-3 font-bold">Pelanggaran</th>
                                        </tr>
                                    </thead>
                                    <tbody id="adminNilaiTbody" class="divide-y divide-slate-100"></tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <!-- MENU 6: AKUN SISWA SIAP DIGUNAKAN & KARTU LOGIN -->
                    <div id="adminSection-kartu" class="hidden space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4 no-print">
                            <div>
                                <h2 class="text-xl font-extrabold text-slate-900">Manajemen Akun Siswa &amp; Kartu Login Resmi</h2>
                                <p class="text-xs text-slate-500 mt-0.5">Kelola akun siswa siap digunakan berbasis NISN, edit/reset password, import Excel, dan cetak Kartu Login A4.</p>
                            </div>
                            <div class="flex flex-wrap items-center gap-2 shrink-0">
                                <button onclick="exportStudentAccountsExcel()" class="px-4 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold transition flex items-center gap-1.5 shadow-sm">
                                    <i data-lucide="file-spreadsheet" class="w-4 h-4"></i> Export Akun (.XLSX)
                                </button>
                                <button onclick="window.print()" class="px-4 py-2.5 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition flex items-center gap-1.5 shadow-sm">
                                    <i data-lucide="printer" class="w-4 h-4"></i> Cetak Kartu Login (A4)
                                </button>
                            </div>
                        </div>

                        <div class="grid grid-cols-1 xl:grid-cols-12 gap-6 no-print">
                            <div class="xl:col-span-4 space-y-5">
                                <div class="bg-white rounded-2xl border border-emerald-200 p-5 shadow-sm space-y-3">
                                    <h3 class="font-bold text-slate-900 text-sm flex items-center gap-2">
                                        <i data-lucide="file-spreadsheet" class="w-4 h-4 text-emerald-600"></i> Import Data Siswa via Excel
                                    </h3>
                                    <p class="text-[11px] text-slate-500 leading-relaxed">
                                        Format kolom: <strong>Nama, NISN, Kelas, Ruangan</strong>. Username otomatis menggunakan NISN dan Password dibuat otomatis.
                                    </p>
                                    <button onclick="downloadStudentExcelTemplate()" class="w-full py-2.5 px-4 rounded-xl bg-emerald-50 hover:bg-emerald-100 text-emerald-700 border border-emerald-200 text-xs font-bold transition flex items-center justify-center gap-2">
                                        <i data-lucide="download" class="w-4 h-4"></i> Download Template Excel (.XLSX)
                                    </button>
                                    <label class="flex flex-col items-center justify-center w-full py-3.5 px-3 border-2 border-dashed border-emerald-300 rounded-xl bg-emerald-50/30 hover:bg-emerald-50 cursor-pointer transition text-center">
                                        <span class="text-xs font-bold text-emerald-900"><i data-lucide="upload" class="w-3.5 h-3.5 inline mr-1"></i> Upload File Excel (.XLSX / .CSV)</span>
                                        <input type="file" accept=".xlsx,.xls,.csv" onchange="handleStudentExcelUpload(event)" class="hidden">
                                    </label>
                                </div>

                                <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
                                    <h3 class="font-bold text-slate-900 text-sm mb-3">Tambah Akun Siswa Manual</h3>
                                    <form onsubmit="handleAddManualStudent(event)" class="space-y-3">
                                        <div>
                                            <label class="block text-xs font-semibold text-slate-700 mb-1">Nama Lengkap *</label>
                                            <input type="text" id="kStuName" required placeholder="Contoh: Nadia Putri" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm">
                                        </div>
                                        <div>
                                            <label class="block text-xs font-semibold text-slate-700 mb-1">NISN *</label>
                                            <input type="text" id="kStuNisn" required placeholder="Contoh: 0081234504" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm font-mono">
                                        </div>
                                        <div class="grid grid-cols-2 gap-2.5">
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Kelas *</label>
                                                <input type="text" id="kStuClass" required placeholder="XII IPA 1" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm">
                                            </div>
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Ruangan *</label>
                                                <input type="text" id="kStuRoom" required placeholder="Ruang 01" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm">
                                            </div>
                                        </div>
                                        <div class="grid grid-cols-2 gap-2.5">
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Username (Opsional)</label>
                                                <input type="text" id="kStuUser" placeholder="Otomatis NISN" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-mono">
                                            </div>
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Password (Opsional)</label>
                                                <input type="text" id="kStuPass" placeholder="Otomatis Acak" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-mono">
                                            </div>
                                        </div>
                                        <button type="submit" class="w-full py-2.5 px-4 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold transition shadow-sm">
                                            + Simpan Akun &amp; Buat Kartu
                                        </button>
                                    </form>
                                </div>
                            </div>

                            <div class="xl:col-span-8 bg-white rounded-2xl border border-slate-200 p-5 shadow-sm flex flex-col space-y-4">
                                <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                                    <div>
                                        <h3 class="font-bold text-slate-900 text-sm sm:text-base">Daftar Akun Siswa Aktif (<span id="studentAccountCount">0</span>)</h3>
                                        <p class="text-[11px] text-slate-500">Gunakan NISN/Username dan Password di bawah untuk login ke Dashboard Siswa</p>
                                    </div>
                                    <div class="flex items-center gap-2 shrink-0">
                                        <button onclick="toggleShowStudentPasswords()" id="btnTogglePassVis" class="px-3 py-1.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold transition flex items-center gap-1.5">
                                            <i data-lucide="eye" class="w-3.5 h-3.5"></i> <span>Tampilkan Password</span>
                                        </button>
                                        <button onclick="confirmClearAllStudents()" class="px-3 py-1.5 rounded-xl bg-red-50 hover:bg-red-100 text-red-600 text-xs font-bold transition">
                                            Kosongkan
                                        </button>
                                    </div>
                                </div>

                                <input type="text" id="accountSearchInput" oninput="renderAdminStudentAccountTable()" placeholder="Cari berdasarkan Nama, NISN, Kelas, atau Ruangan..." class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm w-full focus:ring-2 focus:ring-indigo-500 focus:outline-none">

                                <div class="overflow-x-auto border border-slate-200 rounded-xl max-h-[380px] overflow-y-auto">
                                    <table class="w-full text-left border-collapse text-xs">
                                        <thead>
                                            <tr class="bg-slate-50 border-b border-slate-200 text-[11px] uppercase text-slate-600 sticky top-0">
                                                <th class="py-2.5 px-3 font-bold">Nama &amp; NISN</th>
                                                <th class="py-2.5 px-3 font-bold">Kelas / Ruang</th>
                                                <th class="py-2.5 px-3 font-bold">Username</th>
                                                <th class="py-2.5 px-3 font-bold">Password</th>
                                                <th class="py-2.5 px-3 font-bold text-right">Kelola Akun</th>
                                            </tr>
                                        </thead>
                                        <tbody id="adminStudentAccountTbody" class="divide-y divide-slate-100"></tbody>
                                    </table>
                                </div>
                            </div>
                        </div>

                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                            <div class="flex items-center justify-between gap-2 no-print">
                                <div>
                                    <h3 class="font-bold text-slate-900 text-sm sm:text-base">Pratinjau Kartu Login Peserta Siap Cetak (<span id="cardTotalCount">0</span> Kartu)</h3>
                                    <p class="text-[11px] text-slate-500">Format cetak otomatis 2 kolom per halaman A4</p>
                                </div>
                                <button onclick="window.print()" class="px-4 py-2 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold transition flex items-center gap-1.5 shrink-0">
                                    <i data-lucide="printer" class="w-3.5 h-3.5"></i> Cetak Sekarang
                                </button>
                            </div>
                            <div id="printCardsContainer" class="grid grid-cols-1 sm:grid-cols-2 gap-4"></div>
                        </div>
                    </div>

                    <!-- MENU 7: SETELAN KESELURUHAN -->
                    <div id="adminSection-setelan" class="hidden space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm">
                            <h2 class="text-xl font-extrabold text-slate-900">Setelan Keseluruhan Sistem, Pengawas &amp; Keamanan</h2>
                            <p class="text-xs text-slate-500 mt-0.5">Kelola identitas sekolah, kredensial administrator, akun pengawas ruangan, sirine &amp; keamanan anti-curang, galeri portal, serta backup database.</p>
                        </div>

                        <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                            <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                                <h3 class="font-bold text-slate-900 text-base flex items-center gap-2">
                                    <i data-lucide="school" class="w-4 h-4 text-indigo-600"></i> Identitas Sekolah &amp; Akun Administrator
                                </h3>
                                <div class="space-y-3">
                                    <div>
                                        <label class="block text-xs font-bold text-slate-700 mb-1">Nama Sekolah / Lembaga</label>
                                        <input type="text" id="setSchoolName" class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-sm">
                                    </div>
                                    <div>
                                        <label class="block text-xs font-bold text-slate-700 mb-1">Subjudul / Alamat &amp; Tahun Pelajaran</label>
                                        <input type="text" id="setSchoolSubtitle" class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-sm">
                                    </div>
                                    <div class="pt-2 border-t border-slate-100 grid grid-cols-2 gap-2.5">
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Username Admin</label>
                                            <input type="text" id="setAdminUsername" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-mono font-bold">
                                        </div>
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Password Admin</label>
                                            <input type="text" id="setAdminPassword" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-mono font-bold">
                                        </div>
                                    </div>
                                    <button onclick="saveSchoolIdentity()" class="px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold transition">
                                        Simpan Identitas &amp; Akun Admin
                                    </button>
                                </div>
                            </div>

                            <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                                <h3 class="font-bold text-slate-900 text-base flex items-center gap-2">
                                    <i data-lucide="shield-alert" class="w-4 h-4 text-red-600"></i> Setelan Keamanan, Sirine &amp; Perilaku Siswa
                                </h3>
                                <div class="space-y-3">
                                    <div>
                                        <label class="block text-xs font-bold text-red-700 mb-1">Kode PIN Pembuka Kunci Layar (Default: 123)</label>
                                        <input type="text" id="setUnlockPin" class="w-full px-3.5 py-2 rounded-xl border border-red-300 bg-red-50/40 font-mono font-extrabold text-red-700 text-sm">
                                    </div>
                                    <label class="flex items-center justify-between gap-2 text-xs font-semibold text-slate-700 pt-1">
                                        <span>Wajib Mode Layar Penuh (Fullscreen)</span>
                                        <input type="checkbox" id="setRequireFullscreen" class="w-4 h-4 rounded text-indigo-600 shrink-0">
                                    </label>
                                    <label class="flex items-center justify-between gap-2 text-xs font-semibold text-slate-700">
                                        <span>Kunci Otomatis Jika Pindah Tab / Aplikasi</span>
                                        <input type="checkbox" id="setDetectTabSwitch" class="w-4 h-4 rounded text-indigo-600 shrink-0">
                                    </label>
                                    <label class="flex items-center justify-between gap-2 text-xs font-semibold text-slate-700">
                                        <span>Bunyikan Sirine Alarm Keras Saat Pelanggaran</span>
                                        <input type="checkbox" id="setEnableSiren" class="w-4 h-4 rounded text-red-600 shrink-0">
                                    </label>
                                    <label class="flex items-center justify-between gap-2 text-xs font-semibold text-slate-700">
                                        <span>Tampilkan Skor Akhir ke Siswa Setelah Selesai</span>
                                        <input type="checkbox" id="setShowScoreToStudent" class="w-4 h-4 rounded text-indigo-600 shrink-0">
                                    </label>
                                    <div class="flex flex-wrap gap-2 pt-1">
                                        <button onclick="saveSecuritySettings()" class="px-5 py-2.5 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition">
                                            Simpan Setelan Keamanan
                                        </button>
                                        <button type="button" onclick="testSirenSound()" class="px-4 py-2.5 rounded-xl bg-red-50 hover:bg-red-100 text-red-700 border border-red-200 text-xs font-bold transition flex items-center gap-1.5">
                                            <i data-lucide="volume-2" class="w-3.5 h-3.5"></i> Tes Suara Sirine (2 Detik)
                                        </button>
                                    </div>
                                </div>
                            </div>

                            <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                                <h3 class="font-bold text-slate-900 text-base flex items-center gap-2">
                                    <i data-lucide="user-cog" class="w-4 h-4 text-amber-600"></i> Manajemen Akun Pengawas Ruangan
                                </h3>
                                <form onsubmit="handleAddProctor(event)" class="grid grid-cols-1 sm:grid-cols-2 gap-2.5 bg-slate-50 p-3.5 rounded-xl border border-slate-200">
                                    <input type="text" id="procName" required placeholder="Nama Pengawas" class="px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white">
                                    <input type="text" id="procRoom" required placeholder="Ruang Tugas (Misal: Ruang 01 / Semua)" class="px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white">
                                    <input type="text" id="procUser" required placeholder="Username Login" class="px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white font-mono">
                                    <div class="flex gap-2">
                                        <input type="text" id="procPass" required placeholder="Password" class="flex-1 px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white font-mono">
                                        <button type="submit" class="px-3 py-2 rounded-lg bg-amber-500 hover:bg-amber-600 text-white text-xs font-bold shrink-0">
                                            + Tambah
                                        </button>
                                    </div>
                                </form>
                                <div id="proctorListContainer" class="space-y-2 max-h-52 overflow-y-auto"></div>
                            </div>

                            <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                                <h3 class="font-bold text-slate-900 text-base flex items-center gap-2">
                                    <i data-lucide="image" class="w-4 h-4 text-emerald-600"></i> Galeri Foto Portal &amp; Manajemen Database
                                </h3>
                                <form onsubmit="handleAddGalleryItem(event)" class="space-y-2 bg-slate-50 p-3.5 rounded-xl border border-slate-200">
                                    <input type="text" id="galTitle" required placeholder="Judul Foto Kegiatan (Misal: Pelaksanaan Ujian Sesi 1)" class="w-full px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white">
                                    <div class="flex gap-2">
                                        <input type="url" id="galUrl" placeholder="URL Gambar Eksternal (Kosongkan untuk Ilustrasi SVG)" class="flex-1 px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white">
                                        <button type="submit" class="px-3.5 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold shrink-0">
                                            + Tambah Foto
                                        </button>
                                    </div>
                                </form>

                                <div class="pt-3 border-t border-slate-100 space-y-2">
                                    <span class="text-xs font-bold text-slate-700 block">Cadangan, Pemulihan &amp; Bersihkan Data Awal</span>
                                    <div class="flex flex-wrap gap-2">
                                        <button onclick="exportDatabaseJSON()" class="px-3.5 py-2 rounded-xl bg-blue-50 hover:bg-blue-100 text-blue-700 text-xs font-bold transition flex items-center gap-1.5">
                                            <i data-lucide="download" class="w-3.5 h-3.5"></i> Backup JSON
                                        </button>
                                        <label class="px-3.5 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold transition flex items-center gap-1.5 cursor-pointer">
                                            <i data-lucide="upload" class="w-3.5 h-3.5"></i> Restore JSON
                                            <input type="file" accept=".json" onchange="importDatabaseJSON(event)" class="hidden">
                                        </label>
                                        <button onclick="confirmCleanStartEmpty()" class="px-3.5 py-2 rounded-xl bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 text-xs font-bold transition flex items-center gap-1.5">
                                            <i data-lucide="eraser" class="w-3.5 h-3.5"></i> Mulai Kosong
                                        </button>
                                        <button onclick="confirmFactoryReset()" class="px-3.5 py-2 rounded-xl bg-red-50 hover:bg-red-100 text-red-600 text-xs font-bold transition flex items-center gap-1.5">
                                            <i data-lucide="rotate-ccw" class="w-3.5 h-3.5"></i> Reset Bawaan
                                        </button>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>

                </div>
            </div>
        </section>

        <!-- ================================================================= -->
        <!-- VIEW D: DASHBOARD PENGAWAS RUANG (3 MENU TERSTRUKTUR & SINKRON)   -->
        <!-- ================================================================= -->
        <section id="view-pengawas" class="hidden max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-6">
            <!-- Header Pengawas -->
            <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div>
                    <div class="flex flex-wrap items-center gap-2">
                        <span class="px-2.5 py-0.5 rounded-md bg-amber-100 text-amber-800 text-xs font-bold">DASHBOARD PENGAWAS RUANG</span>
                        <span id="proctorAssignedRoomBadge" class="px-2.5 py-0.5 rounded-md bg-slate-100 text-slate-700 text-xs font-semibold">Semua Ruangan</span>
                    </div>
                    <h2 id="proctorDashboardName" class="text-2xl font-extrabold text-slate-900 mt-1">Pengawas Ujian</h2>
                    <p class="text-xs text-slate-500 mt-0.5">Terhubung real-time dengan Administrator &amp; Peserta Ujian di ruangan.</p>
                </div>

                <!-- Kartu Token Ruang Aktif (Sinkron dengan Admin) -->
                <div class="bg-slate-900 text-white px-5 py-3.5 rounded-2xl flex items-center gap-4 self-start md:self-center shrink-0">
                    <div>
                        <span class="text-[10px] uppercase text-emerald-400 font-bold block">Token Sinkron Admin</span>
                        <span id="proctorTokenDisplay" class="text-2xl font-mono font-extrabold text-amber-400 tracking-widest">------</span>
                    </div>
                    <button onclick="copyTextToClipboard(document.getElementById('proctorTokenDisplay').innerText, 'Token berhasil disalin!')" class="px-3 py-2 rounded-xl bg-white/10 hover:bg-white/20 text-xs font-bold transition flex items-center gap-1.5">
                        <i data-lucide="copy" class="w-4 h-4"></i> Salin
                    </button>
                </div>
            </div>

            <!-- 3 MENU UTAMA PENGAWAS -->
            <div class="flex flex-wrap gap-2 bg-white p-2 rounded-2xl border border-slate-200 shadow-sm">
                <button onclick="switchProctorMenu('kehadiran')" id="procMenuBtn-kehadiran" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-bold bg-amber-500 text-white transition flex items-center gap-2">
                    <i data-lucide="arrow-left-right" class="w-4 h-4"></i> 1. Pantau Siswa Keluar/Masuk &amp; Token
                </button>
                <button onclick="switchProctorMenu('pelanggaran')" id="procMenuBtn-pelanggaran" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition flex items-center gap-2">
                    <i data-lucide="siren" class="w-4 h-4"></i> 2. Pelanggaran &amp; Buka Kunci Sirine
                    <span id="proctorViolationCountPill" class="px-2 py-0.5 rounded-full bg-red-100 text-red-700 text-[10px] font-extrabold">0</span>
                </button>
                <button onclick="switchProctorMenu('soal')" id="procMenuBtn-soal" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition flex items-center gap-2">
                    <i data-lucide="file-plus-2" class="w-4 h-4"></i> 3. Input &amp; Upload Soal (Word / PDF)
                </button>
            </div>

            <!-- MENU 1 PENGAWAS: PANTAU SISWA KELUAR MASUK & TOKEN SINKRON -->
            <div id="proctorSection-kehadiran" class="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
                <div class="lg:col-span-8 bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                    <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                        <div>
                            <h3 class="font-bold text-slate-900 text-base">Status Kehadiran &amp; Pengerjaan Siswa Real-Time</h3>
                            <p class="text-xs text-slate-500">Memantau siswa yang masuk dashboard, sedang mengerjakan soal, atau sudah selesai.</p>
                        </div>
                        <span class="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-emerald-100 text-emerald-700 text-xs font-bold self-start shrink-0">
                            <span class="w-2 h-2 rounded-full bg-emerald-500 animate-ping"></span> Live Sync
                        </span>
                    </div>

                    <div id="proctorStudentCards" class="space-y-3"></div>
                </div>

                <!-- Panel Token & Log Keluar Masuk Siswa Ruang -->
                <div class="lg:col-span-4 space-y-4">
                    <div class="bg-slate-900 text-white rounded-2xl p-5 shadow-sm space-y-3">
                        <div class="flex items-center justify-between">
                            <span class="text-xs font-bold uppercase tracking-wider text-amber-300">Token Ujian Aktif</span>
                            <span class="px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 text-[10px] font-bold">SINKRON ADMIN</span>
                        </div>
                        <div class="text-3xl font-mono font-extrabold tracking-widest text-white" id="proctorCardTokenBig">------</div>
                        <p class="text-[11px] text-slate-300 leading-relaxed">
                            Berikan token ini kepada peserta di ruangan saat mereka siap menekan tombol <strong>Mulai Ujian</strong> pada mata pelajaran yang telah di-assign.
                        </p>
                        <div class="pt-2 border-t border-white/10 flex justify-between text-xs">
                            <span class="text-slate-400">PIN Buka Kunci Layar:</span>
                            <span id="proctorPinDisplay" class="font-mono font-extrabold text-amber-400">123</span>
                        </div>
                    </div>

                    <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm space-y-3">
                        <h4 class="font-bold text-slate-900 text-sm flex items-center gap-2">
                            <i data-lucide="clock" class="w-4 h-4 text-blue-600"></i> Log Keluar / Masuk Siswa Terbaru
                        </h4>
                        <div id="proctorInOutLogList" class="space-y-2 max-h-72 overflow-y-auto text-xs"></div>
                    </div>
                </div>
            </div>

            <!-- MENU 2 PENGAWAS: PELANGGARAN & KENDALI SIRINE -->
            <div id="proctorSection-pelanggaran" class="hidden space-y-6">
                <div class="bg-white rounded-2xl border border-red-200 p-6 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                    <div>
                        <h3 class="text-lg font-extrabold text-slate-900 flex items-center gap-2">
                            <i data-lucide="siren" class="w-5 h-5 text-red-600"></i> Monitoring Pelanggaran &amp; Layar Terkunci (Sirine)
                        </h3>
                        <p class="text-xs text-slate-500 mt-0.5">Daftar peserta yang terdeteksi keluar dari halaman ujian, berpindah tab, atau membuka aplikasi lain.</p>
                    </div>
                    <div class="bg-red-50 border border-red-200 px-4 py-2 rounded-xl text-xs text-red-900 self-start shrink-0">
                        PIN Pembuka Kunci Manual: <strong id="proctorPinDisplay2" class="font-mono text-sm font-extrabold">123</strong>
                    </div>
                </div>

                <div id="proctorActiveLockedBox" class="space-y-3"></div>

                <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-3">
                    <h4 class="font-bold text-slate-900 text-sm">Riwayat Catatan Pelanggaran Peserta</h4>
                    <div class="overflow-x-auto border border-slate-200 rounded-xl">
                        <table class="w-full text-left border-collapse text-xs sm:text-sm">
                            <thead>
                                <tr class="bg-slate-50 border-b border-slate-200 text-[11px] uppercase text-slate-600">
                                    <th class="py-3 px-4 font-bold">Waktu</th>
                                    <th class="py-3 px-4 font-bold">Peserta &amp; NISN</th>
                                    <th class="py-3 px-4 font-bold">Keterangan Pelanggaran</th>
                                </tr>
                            </thead>
                            <tbody id="proctorViolationLogsTbody" class="divide-y divide-slate-100"></tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- MENU 3 PENGAWAS: INPUT & UPLOAD SOAL WORD / PDF -->
            <div id="proctorSection-soal" class="hidden grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
                <div class="lg:col-span-5 space-y-5">
                    <div class="bg-white rounded-2xl border border-amber-200 p-5 shadow-sm space-y-3">
                        <div class="flex items-center justify-between gap-2">
                            <h3 class="font-bold text-slate-900 text-sm flex items-center gap-2">
                                <i data-lucide="file-up" class="w-4 h-4 text-amber-600 shrink-0"></i> Upload File Word (.docx) / PDF (.pdf)
                            </h3>
                            <button onclick="downloadSampleQuestionFormat()" class="text-[11px] font-bold text-amber-700 hover:underline shrink-0">
                                Unduh Contoh Format
                            </button>
                        </div>

                        <div>
                            <label class="block text-xs font-bold text-slate-700 mb-1">Pilih Paket Mata Pelajaran Tujuan:</label>
                            <select id="proctorPackageSelector" onchange="selectProctorQuestionPackage(this.value)" class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm font-bold bg-white"></select>
                        </div>

                        <label class="flex flex-col items-center justify-center w-full h-28 border-2 border-dashed border-amber-300 rounded-2xl bg-amber-50/40 hover:bg-amber-50 cursor-pointer transition p-4 text-center">
                            <i data-lucide="upload-cloud" class="w-7 h-7 text-amber-600 mb-1"></i>
                            <span class="text-xs font-bold text-amber-900">Klik Pilih File .DOCX, .PDF, atau .TXT</span>
                            <span class="text-[10px] text-slate-500 mt-0.5">Langsung tersinkron ke Bank Soal Administrator</span>
                            <input type="file" accept=".docx,.pdf,.txt,.csv" onchange="handleDocumentQuestionUpload(event, 'pengawas')" class="hidden">
                        </label>
                    </div>

                    <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
                        <h3 class="font-bold text-slate-900 text-sm mb-3">Input Soal Manual (Sinkron ke Bank Soal)</h3>
                        <form onsubmit="handleAddManualQuestion(event, 'pengawas')" class="space-y-3">
                            <div>
                                <label class="block text-xs font-semibold text-slate-700 mb-1">Teks Pertanyaan *</label>
                                <textarea id="pQText" required rows="2" placeholder="Tuliskan pertanyaan soal..." class="w-full p-2.5 rounded-xl border border-slate-300 text-xs sm:text-sm focus:ring-2 focus:ring-amber-500 focus:outline-none"></textarea>
                            </div>
                            <div class="space-y-2">
                                <div class="flex items-center gap-2">
                                    <span class="w-6 text-xs font-bold text-slate-500 shrink-0">A.</span>
                                    <input type="text" id="pQOptA" required placeholder="Opsi A" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                </div>
                                <div class="flex items-center gap-2">
                                    <span class="w-6 text-xs font-bold text-slate-500 shrink-0">B.</span>
                                    <input type="text" id="pQOptB" required placeholder="Opsi B" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                </div>
                                <div class="flex items-center gap-2">
                                    <span class="w-6 text-xs font-bold text-slate-500 shrink-0">C.</span>
                                    <input type="text" id="pQOptC" required placeholder="Opsi C" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                </div>
                                <div class="flex items-center gap-2">
                                    <span class="w-6 text-xs font-bold text-slate-500 shrink-0">D.</span>
                                    <input type="text" id="pQOptD" required placeholder="Opsi D" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                </div>
                                <div class="flex items-center gap-2">
                                    <span class="w-6 text-xs font-bold text-slate-500 shrink-0">E.</span>
                                    <input type="text" id="pQOptE" placeholder="Opsi E (Opsional)" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                </div>
                            </div>
                            <div class="grid grid-cols-2 gap-3 pt-1">
                                <div>
                                    <label class="block text-xs font-semibold text-slate-700 mb-1">Kunci Jawaban</label>
                                    <select id="pQAnswer" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-bold text-amber-800">
                                        <option value="A">Opsi A</option>
                                        <option value="B">Opsi B</option>
                                        <option value="C">Opsi C</option>
                                        <option value="D">Opsi D</option>
                                        <option value="E">Opsi E</option>
                                    </select>
                                </div>
                                <div>
                                    <label class="block text-xs font-semibold text-slate-700 mb-1">Bobot Poin</label>
                                    <input type="number" id="pQPoints" value="10" min="1" max="100" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-bold">
                                </div>
                            </div>
                            <button type="submit" class="w-full py-2.5 px-4 rounded-xl bg-amber-500 hover:bg-amber-600 text-white text-xs font-bold transition shadow-sm">
                                + Simpan Soal ke Paket Terpilih
                            </button>
                        </form>
                    </div>
                </div>

                <div class="lg:col-span-7 bg-white rounded-2xl border border-slate-200 p-5 shadow-sm space-y-3">
                    <div class="flex items-center justify-between pb-3 border-b border-slate-100">
                        <div>
                            <h3 class="font-bold text-slate-900 text-sm sm:text-base">Daftar Soal Paket Terpilih (<span id="proctorPkgSoalCount">0</span> Soal)</h3>
                            <p class="text-[11px] text-slate-500">Tersinkronisasi langsung dengan Bank Soal Administrator</p>
                        </div>
                    </div>
                    <div id="proctorSoalListContainer" class="space-y-3 max-h-[640px] overflow-y-auto pr-1"></div>
                </div>
            </div>
        </section>

        <!-- ================================================================= -->
        <!-- VIEW E: HALAMAN UJIAN SISWA (ANTI-CHEAT LOCKDOWN + SIRINE)        -->
        <!-- ================================================================= -->
        <section id="view-siswa" class="min-h-screen bg-slate-100 flex flex-col no-select hidden">
            <div class="bg-slate-900 text-white px-4 sm:px-6 py-3.5 shadow-md sticky top-0 z-30">
                <div class="max-w-7xl mx-auto flex flex-wrap items-center justify-between gap-3">
                    <div class="flex items-center gap-3 min-w-0">
                        <div class="w-10 h-10 rounded-xl bg-blue-600 flex items-center justify-center font-extrabold shrink-0">
                            <i data-lucide="shield-check" class="w-5 h-5"></i>
                        </div>
                        <div class="min-w-0">
                            <h2 id="studentExamTitle" class="font-bold text-sm sm:text-base leading-tight truncate">Ujian Berbasis Komputer</h2>
                            <p class="text-xs text-slate-300 truncate">
                                <strong id="studentExamName" class="text-white">-</strong> • NISN: <span id="studentExamNisn">-</span> • <span id="studentExamRoom">-</span>
                            </p>
                        </div>
                    </div>

                    <div class="flex items-center gap-2 sm:gap-3 shrink-0">
                        <div class="hidden sm:flex items-center bg-slate-800 border border-slate-700 rounded-xl p-1">
                            <button onclick="changeExamFontSize('sm')" class="px-2 py-1 rounded-lg text-xs font-bold hover:bg-slate-700" title="Perkecil Teks">A-</button>
                            <button onclick="changeExamFontSize('base')" class="px-2 py-1 rounded-lg text-xs font-bold bg-slate-700" title="Ukuran Normal">A</button>
                            <button onclick="changeExamFontSize('lg')" class="px-2 py-1 rounded-lg text-xs font-bold hover:bg-slate-700" title="Perbesar Teks">A+</button>
                        </div>

                        <div class="bg-slate-800 border border-slate-700 px-3.5 py-1.5 rounded-xl flex items-center gap-2">
                            <i data-lucide="clock" class="w-4 h-4 text-amber-400"></i>
                            <span id="studentTimerDisplay" class="font-mono font-extrabold text-base sm:text-lg text-amber-400">00:00</span>
                        </div>

                        <button onclick="confirmSubmitExam()" class="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-xs sm:text-sm font-bold transition shadow">
                            Selesai Ujian
                        </button>
                    </div>
                </div>
            </div>
