/* ====================================================
   ARKAAN HILMI SUHARSOYO - PORTFOLIO MAIN SCRIPT
   Sistem Informasi · ITS Surabaya
   ==================================================== */

// ============================
// LANDING PAGE
// ============================
let selRole = '';

/**
 * Fungsi untuk memilih role pengunjung
 * @param {string} r - Role yang dipilih
 * @param {HTMLElement} el - Element yang diklik
 */
function pickRole(r, el) {
  selRole = r;
  document.querySelectorAll('.role-card').forEach(c => c.classList.remove('sel'));
  el.classList.add('sel');
}

/**
 * Fungsi untuk masuk ke halaman utama
 */
function enter() {
  const name = document.getElementById('vname').value.trim();
  const role = selRole || 'Pengunjung';
  let w = '';

  // Personalisasi sambutan berdasarkan role dan nama
  if (name) {
    w = `Halo, ${name}! Selamat datang 👋`;
  } else if (role === 'HRD Perusahaan') {
    w = 'Halo, Bapak/Ibu HRD! Selamat datang 👋';
  } else if (role === 'Mahasiswa') {
    w = 'Halo, sesama Mahasiswa! Selamat datang 👋';
  } else if (role === 'Kolega / Rekan') {
    w = 'Halo, Rekan! Selamat datang 👋';
  } else if (role === 'Calon Klien') {
    w = 'Halo! Terima kasih sudah mampir 👋';
  } else {
    w = 'Selamat datang di portfolio Arkaan! 👋';
  }

  // Update welcome message dan transisi ke halaman utama
  document.getElementById('hwelcome').textContent = w;
  document.getElementById('landing').classList.add('hidden');

  setTimeout(() => {
    document.getElementById('navbar').style.display = 'flex';
    document.getElementById('main-content').classList.add('vis');
    document.getElementById('landing').style.display = 'none';
    setTimeout(animateBars, 600);
    initReveal();
  }, 800);
}

// Event listener untuk Enter key pada input nama
document.addEventListener('DOMContentLoaded', () => {
  const vnameInput = document.getElementById('vname');
  if (vnameInput) {
    vnameInput.addEventListener('keydown', e => {
      if (e.key === 'Enter') enter();
    });
  }
});

// ============================
// NAVBAR
// ============================

/**
 * Toggle mobile navigation menu
 */
function toggleNav() {
  document.getElementById('nlinks').classList.toggle('open');
}

/**
 * Update active navigation link based on scroll position
 */
window.addEventListener('scroll', () => {
  const sy = window.scrollY + 90;
  ['home', 'education', 'experience', 'projects', 'training', 'gallery', 'freelance'].forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    const a = document.querySelector(`.n-links a[href="#${id}"]`);
    if (a) {
      a.classList.toggle('active', sy >= el.offsetTop && sy < el.offsetTop + el.offsetHeight);
    }
  });
});

// ============================
// SKILL BARS ANIMATION
// ============================

/**
 * Animasi progress bar untuk skill
 */
function animateBars() {
  document.querySelectorAll('.s-fill').forEach(b => {
    b.style.width = b.getAttribute('data-w') + '%';
  });
}

// ============================
// SCROLL REVEAL ANIMATION
// ============================

/**
 * Inisialisasi Intersection Observer untuk scroll reveal effect
 */
function initReveal() {
  const io = new IntersectionObserver(
    entries => {
      entries.forEach(e => {
        if (e.isIntersecting) {
          e.target.classList.add('vis');
        }
      });
    },
    {
      threshold: 0.1,
      rootMargin: '0px 0px -35px 0px'
    }
  );

  document.querySelectorAll('.reveal').forEach(el => io.observe(el));
}

// ============================
// MESSAGE FORM
// ============================

/**
 * Fungsi untuk mengirim pesan dari form kontak
 */
function sendMsg() {
  const n = document.getElementById('mn').value.trim();
  const m = document.getElementById('mm').value.trim();

  if (!n || !m) {
    alert('Mohon isi nama dan pesan terlebih dahulu.');
    return;
  }

  // Tampilkan pesan sukses
  const ok = document.getElementById('ok-msg');
  ok.style.display = 'block';

  // Reset form
  document.getElementById('mn').value = '';
  document.getElementById('mc').value = '';
  document.getElementById('mm').value = '';
  document.getElementById('mp').value = '';

  // Sembunyikan pesan sukses setelah 5 detik
  setTimeout(() => {
    ok.style.display = 'none';
  }, 5000);
}

// ============================
// DOWNLOAD PPT
// ============================

/**
 * Fungsi untuk mendownload portfolio sebagai PowerPoint
 */
async function dlPPT() {
  const btn = document.querySelector('.btn-ppt');
  const prog = document.getElementById('ppt-prog');
  const fill = document.getElementById('pfill');
  const st = document.getElementById('ppt-status');

  btn.disabled = true;
  btn.innerHTML = '⏳ Memproses…';
  prog.style.display = 'block';

  // Simulasi progress dengan step-by-step
  const steps = [
    [10, 'Mengumpulkan data profil…'],
    [25, 'Membuat slide Pendidikan…'],
    [45, 'Membuat slide Pengalaman…'],
    [62, 'Membuat slide Proyek…'],
    [78, 'Membuat slide Keahlian…'],
    [92, 'Menyusun slide Freelance & Penutup…'],
    [100, 'Selesai! Mengunduh…']
  ];

  for (const [p, s] of steps) {
    fill.style.width = p + '%';
    st.textContent = s;
    await new Promise(r => setTimeout(r, 380));
  }

  // Generate PPT
  await makePPT();

  // Reset button
  btn.disabled = false;
  btn.innerHTML = '✅ Berhasil Diunduh!';
  setTimeout(() => {
    btn.innerHTML = '⬇️ Download Portfolio.pptx';
    prog.style.display = 'none';
    fill.style.width = '0';
  }, 3000);
}

/**
 * Fungsi untuk membuat dan menggenerate file PowerPoint
 */
async function makePPT() {
  // Load PptxGenJS library jika belum ada
  if (!window.PptxGenJS) {
    await new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'https://cdn.jsdelivr.net/npm/pptxgenjs@3/dist/pptxgen.bundle.js';
      s.onload = res;
      s.onerror = rej;
      document.head.appendChild(s);
    });
  }

  const pptx = new PptxGenJS();
  pptx.title = 'Portfolio – Arkaan Hilmi Suharsoyo';
  pptx.author = 'Arkaan Hilmi Suharsoyo';

  // Color constants
  const N = '0D1B3E',
    A = '2563EB',
    G = 'C9A84C',
    W = 'FFFFFF',
    LG = 'F1F5F9',
    DG = '64748B';

  /**
   * Helper function untuk background gelap
   */
  function darkBg(sl) {
    sl.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: N } });
    sl.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.08, fill: { color: G } });
    sl.addShape(pptx.ShapeType.rect, { x: 0, y: 7.42, w: '100%', h: 0.08, fill: { color: G } });
  }

  /**
   * Helper function untuk background terang
   */
  function lightBg(sl) {
    sl.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: LG } });
    sl.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.08, fill: { color: N } });
  }

  /**
   * Helper function untuk nomor halaman
   */
  function pg(sl, n) {
    sl.addText(`${n}`, { x: 9.1, y: 7.1, w: 0.5, h: 0.3, fontSize: 9, color: 'AAAAAA', align: 'right' });
  }

  // SLIDE 1 – COVER
  const s1 = pptx.addSlide();
  darkBg(s1);
  s1.addShape(pptx.ShapeType.rect, { x: 0.6, y: 1.8, w: 8.8, h: 4, fill: { color: '1A2F5E' }, line: { color: '2A4A8A', width: 1 } });
  s1.addShape(pptx.ShapeType.ellipse, { x: 3.7, y: 0.85, w: 1.1, h: 1.1, fill: { color: A }, line: { color: G, width: 3 } });
  s1.addText('AH', { x: 3.7, y: 0.85, w: 1.1, h: 1.1, fontSize: 24, bold: true, color: W, align: 'center', valign: 'middle' });
  s1.addText('ARKAAN HILMI SUHARSOYO', { x: 1, y: 2.05, w: 8, h: 0.5, fontSize: 25, bold: true, color: W, align: 'center', fontFace: 'Georgia' });
  s1.addShape(pptx.ShapeType.rect, { x: 3.5, y: 2.68, w: 3, h: 0.05, fill: { color: G } });
  s1.addText('Mahasiswa Sistem Informasi', { x: 1, y: 2.85, w: 8, h: 0.35, fontSize: 13, color: 'A0B4D0', align: 'center' });
  s1.addText('Institut Teknologi Sepuluh Nopember · Surabaya', { x: 1, y: 3.22, w: 8, h: 0.3, fontSize: 11, color: '7A90A8', align: 'center' });
  s1.addText('🌐 Web Dev  •  📊 Data Analytics  •  🤖 Machine Learning  •  📸 Fotografi', { x: 1, y: 3.82, w: 8, h: 0.32, fontSize: 10, color: 'A0B4D0', align: 'center' });
  s1.addText('P O R T F O L I O', { x: 1, y: 6.95, w: 8, h: 0.28, fontSize: 8, color: G, align: 'center', bold: true, charSpacing: 6 });

  // SLIDE 2 – PROFIL
  const s2 = pptx.addSlide();
  lightBg(s2);
  pg(s2, 2);
  s2.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 3.2, h: 7.5, fill: { color: N } });
  s2.addShape(pptx.ShapeType.ellipse, { x: 0.85, y: 0.8, w: 1.5, h: 1.5, fill: { color: A }, line: { color: G, width: 2 } });
  s2.addText('AH', { x: 0.85, y: 0.8, w: 1.5, h: 1.5, fontSize: 28, bold: true, color: W, align: 'center', valign: 'middle' });
  s2.addText('Arkaan Hilmi\nSuharsoyo', { x: 0.15, y: 2.55, w: 2.9, h: 0.72, fontSize: 14, bold: true, color: W, align: 'center', fontFace: 'Georgia' });
  s2.addText('Sistem Informasi · ITS', { x: 0.15, y: 3.32, w: 2.9, h: 0.32, fontSize: 9, color: 'A0B4D0', align: 'center' });

  [
    ['📧', 'arkaan@email.com'],
    ['💼', 'LinkedIn'],
    ['🐙', 'GitHub'],
    ['📱', 'WhatsApp']
  ].forEach(([ic, lb], i) => {
    s2.addText(`${ic} ${lb}`, { x: 0.15, y: 4.2 + i * 0.4, w: 2.9, h: 0.32, fontSize: 9, color: 'A0B4D0', align: 'center' });
  });

  s2.addText('PROFIL', { x: 3.5, y: 0.3, w: 5.8, h: 0.35, fontSize: 11, bold: true, color: N, charSpacing: 4 });
  s2.addShape(pptx.ShapeType.rect, { x: 3.5, y: 0.7, w: 1, h: 0.05, fill: { color: G } });
  s2.addText(
    'Mahasiswa tekno-kreatif dari ITS Surabaya yang passionate di bidang web development, analisis data, machine learning, dan visualisasi. Berorientasi pada solusi nyata yang berdampak.',
    { x: 3.5, y: 0.9, w: 5.8, h: 0.95, fontSize: 10, color: DG, wrap: true }
  );

  [
    ['IPK', '3.7+'],
    ['Proyek', '10+'],
    ['Org', '5+']
  ].forEach(([lb, vl], i) => {
    const xp = 3.5 + i * 2;
    s2.addShape(pptx.ShapeType.rect, { x: xp, y: 2.1, w: 1.8, h: 0.88, fill: { color: 'E8ECF4' }, line: { color: 'D0D8EC', width: 0.5 } });
    s2.addText(vl, { x: xp, y: 2.18, w: 1.8, h: 0.42, fontSize: 18, bold: true, color: N, align: 'center', fontFace: 'Georgia' });
    s2.addText(lb, { x: xp, y: 2.62, w: 1.8, h: 0.25, fontSize: 8, color: DG, align: 'center' });
  });

  s2.addText('MINAT KE DEPAN', { x: 3.5, y: 3.2, w: 5.8, h: 0.28, fontSize: 9, bold: true, color: N, charSpacing: 2 });
  ['🤖 AI / ML Engineer', '🌐 Full-Stack Dev', '📊 Data Scientist', '☁️ Cloud & DevOps', '📸 Creative Director', '🎓 Educator / Mentor'].forEach((it, i) => {
    const col = i % 3,
      row = Math.floor(i / 3);
    s2.addShape(pptx.ShapeType.rect, { x: 3.5 + col * 1.95, y: 3.6 + row * 0.55, w: 1.82, h: 0.42, fill: { color: W }, line: { color: 'D0D8EC', width: 0.5 } });
    s2.addText(it, { x: 3.5 + col * 1.95, y: 3.6 + row * 0.55, w: 1.82, h: 0.42, fontSize: 8, color: N, align: 'center', valign: 'middle', wrap: true });
  });

  // SLIDE 3 – PENDIDIKAN
  const s3 = pptx.addSlide();
  lightBg(s3);
  pg(s3, 3);
  s3.addText('PENDIDIKAN', { x: 0.5, y: 0.2, w: 9, h: 0.38, fontSize: 11, bold: true, color: N, charSpacing: 4 });
  s3.addShape(pptx.ShapeType.rect, { x: 0.5, y: 0.62, w: 1, h: 0.05, fill: { color: G } });

  [
    { logo: 'ITS', period: '2022 — Sekarang', school: 'Institut Teknologi Sepuluh Nopember', major: 'Sistem Informasi · FTEIC', ipk: 'IPK 3.7+', y: 0.9 },
    { logo: 'SMA', period: '2019 — 2022', school: 'SMA [Nama Sekolah]', major: 'IPA · [Kota Asal]', ipk: 'Nilai Memuaskan', y: 4.0 }
  ].forEach(ed => {
    s3.addShape(pptx.ShapeType.rect, { x: 0.4, y: ed.y, w: 9.2, h: 2.8, fill: { color: W }, line: { color: 'D0D8EC', width: 0.5 } });
    s3.addShape(pptx.ShapeType.rect, { x: 0.4, y: ed.y, w: 0.08, h: 2.8, fill: { color: A } });
    s3.addShape(pptx.ShapeType.rect, { x: 0.6, y: ed.y + 0.28, w: 1, h: 1, fill: { color: N } });
    s3.addText(ed.logo, { x: 0.6, y: ed.y + 0.28, w: 1, h: 1, fontSize: ed.logo === 'ITS' ? 14 : 10, bold: true, color: W, align: 'center', valign: 'middle' });
    s3.addText(ed.period, { x: 1.75, y: ed.y + 0.2, w: 7, h: 0.26, fontSize: 9, color: A, bold: true, fontFace: 'Courier New' });
    s3.addText(ed.school, { x: 1.75, y: ed.y + 0.5, w: 7, h: 0.35, fontSize: 13, bold: true, color: N, fontFace: 'Georgia' });
    s3.addText(ed.major, { x: 1.75, y: ed.y + 0.88, w: 7, h: 0.28, fontSize: 10, color: DG });
    s3.addShape(pptx.ShapeType.rect, { x: 1.75, y: ed.y + 1.22, w: 1.1, h: 0.28, fill: { color: 'FDF5DC' }, line: { color: 'E0D0A0', width: 0.5 } });
    s3.addText(ed.ipk, { x: 1.75, y: ed.y + 1.22, w: 1.1, h: 0.28, fontSize: 9, color: '906020', align: 'center', valign: 'middle', bold: true });
    s3.addText('Capaian: Sistem informasi manajemen, analitik data, dashboard Power BI, aplikasi web full-stack menggunakan Next.js & Express.', {
      x: 1.75,
      y: ed.y + 1.65,
      w: 7.2,
      h: 0.6,
      fontSize: 9,
      color: DG,
      wrap: true
    });
  });

  // SLIDE 4 – PENGALAMAN
  const s4 = pptx.addSlide();
  lightBg(s4);
  pg(s4, 4);
  s4.addText('PENGALAMAN', { x: 0.5, y: 0.2, w: 9, h: 0.38, fontSize: 11, bold: true, color: N, charSpacing: 4 });
  s4.addShape(pptx.ShapeType.rect, { x: 0.5, y: 0.62, w: 1, h: 0.05, fill: { color: G } });

  [
    { icon: '🏛️', type: 'ORG', period: '2023—Kini', title: 'Pengurus HMSI', org: 'Himpunan Mahasiswa SI ITS', desc: 'Kegiatan kemahasiswaan & kepanitiaan' },
    { icon: '💻', type: 'WORK', period: '2024—Kini', title: 'Web Developer', org: 'Freelance', desc: 'Membangun web & aplikasi modern' },
    { icon: '📸', type: 'ORG', period: '2022—2023', title: 'Fotografer & Videografer', org: 'UKM Fotografi ITS', desc: 'Konten visual & dokumentasi kampus' },
    { icon: '📊', type: 'WORK', period: '2024', title: 'Data Analyst Intern', org: '[Perusahaan]', desc: 'Python, Power BI, dashboard analitik' },
    { icon: '📐', type: 'ORG', period: '2023', title: 'Mentor Matematika', org: 'Program Bimbel ITS', desc: 'Mengajar mahasiswa junior' },
    { icon: '🤝', type: 'ORG', period: '2022—2023', title: 'Panitia Acara', org: 'Berbagai Kepanitiaan ITS', desc: 'PKK, seminar, kompetisi mahasiswa' }
  ].forEach((ex, i) => {
    const col = i % 3,
      row = Math.floor(i / 3),
      xp = 0.4 + col * 3.1,
      yp = 0.88 + row * 2.9;
    s4.addShape(pptx.ShapeType.rect, { x: xp, y: yp, w: 2.9, h: 2.6, fill: { color: W }, line: { color: 'D0D8EC', width: 0.5 } });
    const bc = ex.type === 'ORG' ? 'FDF8E1' : 'EFF6FF',
      btc = ex.type === 'ORG' ? '906020' : A;
    s4.addShape(pptx.ShapeType.rect, { x: xp + 1.6, y: yp + 0.12, w: 1.15, h: 0.25, fill: { color: bc }, line: { color: 'E0D0B0', width: 0.3 } });
    s4.addText(ex.type === 'ORG' ? 'ORGANISASI' : 'KERJA', {
      x: xp + 1.6,
      y: yp + 0.12,
      w: 1.15,
      h: 0.25,
      fontSize: 6.5,
      color: btc,
      bold: true,
      align: 'center',
      valign: 'middle',
      charSpacing: 1
    });
    s4.addText(ex.icon, { x: xp + 0.1, y: yp + 0.18, w: 0.62, h: 0.62, fontSize: 20, align: 'center' });
    s4.addText(ex.period, { x: xp + 0.78, y: yp + 0.2, w: 1.9, h: 0.22, fontSize: 8, color: A, fontFace: 'Courier New', bold: true });
    s4.addText(ex.title, { x: xp + 0.1, y: yp + 0.84, w: 2.7, h: 0.35, fontSize: 11, bold: true, color: N, fontFace: 'Georgia', wrap: true });
    s4.addText(ex.org, { x: xp + 0.1, y: yp + 1.24, w: 2.7, h: 0.25, fontSize: 9, color: A });
    s4.addText(ex.desc, { x: xp + 0.1, y: yp + 1.58, w: 2.7, h: 0.5, fontSize: 8.5, color: DG, wrap: true });
  });

  // SLIDE 5 – PROYEK
  const s5 = pptx.addSlide();
  lightBg(s5);
  pg(s5, 5);
  s5.addText('PROYEK & KARYA', { x: 0.5, y: 0.2, w: 9, h: 0.38, fontSize: 11, bold: true, color: N, charSpacing: 4 });
  s5.addShape(pptx.ShapeType.rect, { x: 0.5, y: 0.62, w: 1, h: 0.05, fill: { color: G } });

  [
    { icon: '🌐', name: 'Sistem Manajemen Informasi', tech: 'Next.js · React · TailwindCSS', desc: 'Aplikasi web full-stack dengan dashboard analitik real-time' },
    { icon: '📊', name: 'Dashboard Power BI', tech: 'Power BI · Python · Pandas', desc: 'Visualisasi data interaktif untuk keputusan bisnis' },
    { icon: '🤖', name: 'Model Machine Learning', tech: 'Python · Scikit-learn · TensorFlow', desc: 'Model prediktif klasifikasi & regresi data real-world' },
    { icon: '📱', name: 'Portfolio Website', tech: 'HTML · CSS · JavaScript', desc: 'Website portfolio responsif dengan animasi & download PPT' },
    { icon: '📸', name: 'Konten Kreatif Visual', tech: 'Fotografi · Videografi · Editing', desc: 'Produksi konten foto & video profesional' },
    { icon: '📝', name: 'Laporan & Riset Ilmiah', tech: 'Analisis · Riset · Visualisasi', desc: 'Laporan akademik dengan metodologi terstruktur' }
  ].forEach((pr, i) => {
    const col = i % 3,
      row = Math.floor(i / 3),
      xp = 0.4 + col * 3.1,
      yp = 0.88 + row * 2.9;
    s5.addShape(pptx.ShapeType.rect, { x: xp, y: yp, w: 2.9, h: 2.6, fill: { color: N }, line: { color: '2A4A8A', width: 0.5 } });
    s5.addShape(pptx.ShapeType.rect, { x: xp, y: yp, w: 2.9, h: 1.08, fill: { color: '1A2F5E' } });
    s5.addText(pr.icon, { x: xp, y: yp + 0.08, w: 2.9, h: 0.9, fontSize: 28, align: 'center', valign: 'middle' });
    s5.addText(pr.name, { x: xp + 0.12, y: yp + 1.15, w: 2.65, h: 0.42, fontSize: 10, bold: true, color: W, fontFace: 'Georgia', wrap: true });
    s5.addText(pr.tech, { x: xp + 0.12, y: yp + 1.62, w: 2.65, h: 0.22, fontSize: 7.5, color: G, fontFace: 'Courier New' });
    s5.addText(pr.desc, { x: xp + 0.12, y: yp + 1.9, w: 2.65, h: 0.5, fontSize: 8, color: 'A0B4D0', wrap: true });
  });

  // SLIDE 6 – KEAHLIAN
  const s6 = pptx.addSlide();
  darkBg(s6);
  pg(s6, 6);
  s6.addText('KEAHLIAN & PELATIHAN', { x: 0.5, y: 0.2, w: 9, h: 0.38, fontSize: 11, bold: true, color: W, charSpacing: 4 });
  s6.addShape(pptx.ShapeType.rect, { x: 0.5, y: 0.62, w: 1, h: 0.05, fill: { color: G } });

  [
    { name: 'Web Development', pct: 85 },
    { name: 'Data Analytics', pct: 80 },
    { name: 'Machine Learning', pct: 72 },
    { name: 'Power BI / Visualisasi', pct: 78 },
    { name: 'Fotografi & Video', pct: 75 },
    { name: 'Matematika & Statistik', pct: 82 }
  ].forEach((sk, i) => {
    const yp = 0.95 + i * 0.72;
    s6.addText(sk.name, { x: 0.5, y: yp, w: 3.5, h: 0.28, fontSize: 9.5, color: W, bold: true });
    s6.addText(`${sk.pct}%`, { x: 4.05, y: yp, w: 0.5, h: 0.28, fontSize: 9, color: G, bold: true, fontFace: 'Courier New' });
    s6.addShape(pptx.ShapeType.rect, { x: 0.5, y: yp + 0.3, w: 4.5, h: 0.18, fill: { color: '1A3060' } });
    s6.addShape(pptx.ShapeType.rect, { x: 0.5, y: yp + 0.3, w: (4.5 * sk.pct) / 100, h: 0.18, fill: { color: A } });
  });

  s6.addText('PELATIHAN & SERTIFIKASI', { x: 5.3, y: 0.75, w: 4.3, h: 0.28, fontSize: 9, bold: true, color: G, charSpacing: 2 });
  [
    '🌐 Web Development Bootcamp',
    '📊 Data Analytics Professional (Google)',
    '🤖 Machine Learning Fundamentals',
    '📈 Power BI Data Visualization',
    '🐍 Python for Data Science',
    '📸 Fotografi & Videografi Profesional'
  ].forEach((tr, i) => {
    const yp = 1.1 + i * 0.54;
    s6.addShape(pptx.ShapeType.rect, { x: 5.3, y: yp, w: 4.2, h: 0.42, fill: { color: '122050' }, line: { color: '1E3060', width: 0.5 } });
    s6.addText(tr, { x: 5.45, y: yp, w: 4.0, h: 0.42, fontSize: 9, color: 'A0C0E0', valign: 'middle', wrap: true });
  });

  // SLIDE 7 – FREELANCE
  const s7 = pptx.addSlide();
  darkBg(s7);
  pg(s7, 7);
  s7.addShape(pptx.ShapeType.rect, { x: 3.2, y: 0.12, w: 1.6, h: 0.3, fill: { color: '0A4020' }, line: { color: '2A6040', width: 0.5 } });
  s7.addText('● OPEN FOR WORK', { x: 3.2, y: 0.12, w: 1.6, h: 0.3, fontSize: 7.5, color: '4ADE80', bold: true, align: 'center', valign: 'middle', charSpacing: 1 });
  s7.addText('JASA FREELANCE & REMOTE', { x: 0.5, y: 0.58, w: 9, h: 0.4, fontSize: 14, bold: true, color: W, align: 'center', fontFace: 'Georgia' });
  s7.addText('Tersedia untuk pekerjaan freelance dan remote dalam berbagai bidang keahlian', { x: 0.5, y: 1.0, w: 9, h: 0.28, fontSize: 9.5, color: 'A0B4D0', align: 'center' });

  [
    { icon: '🌐', name: 'Website Development' },
    { icon: '📐', name: 'Tutorial Matematika' },
    { icon: '💻', name: 'Tutorial Pemrograman' },
    { icon: '📝', name: 'Laporan & Tugas' },
    { icon: '📊', name: 'Power BI Visualisasi' },
    { icon: '🤖', name: 'Machine Learning' },
    { icon: '📸', name: 'Fotografi' },
    { icon: '🎬', name: 'Videografi' }
  ].forEach((fl, i) => {
    const col = i % 4,
      row = Math.floor(i / 4),
      xp = 0.4 + col * 2.3,
      yp = 1.48 + row * 2.2;
    s7.addShape(pptx.ShapeType.rect, { x: xp, y: yp, w: 2.15, h: 1.9, fill: { color: '0E2248' }, line: { color: '1A3060', width: 0.5 } });
    s7.addText(fl.icon, { x: xp, y: yp + 0.18, w: 2.15, h: 0.7, fontSize: 26, align: 'center' });
    s7.addText(fl.name, { x: xp + 0.1, y: yp + 1.02, w: 1.95, h: 0.62, fontSize: 9, bold: true, color: W, align: 'center', valign: 'middle', wrap: true });
  });

  s7.addText('📧 arkaan@email.com  •  📱 WhatsApp  •  💼 LinkedIn  •  🐙 GitHub', { x: 0.5, y: 6.9, w: 9, h: 0.28, fontSize: 9, color: G, align: 'center' });

  // SLIDE 8 – PENUTUP
  const s8 = pptx.addSlide();
  darkBg(s8);
  s8.addShape(pptx.ShapeType.rect, { x: 0.5, y: 1.5, w: 9, h: 4.5, fill: { color: '0E2248' }, line: { color: '1A3060', width: 0.5 } });
  s8.addText('"', { x: 0.7, y: 1.55, w: 1, h: 1, fontSize: 80, color: '1E4080', fontFace: 'Georgia', bold: true });
  s8.addText('Saya percaya bahwa teknologi, kreativitas, dan data dapat bersinergi untuk menciptakan solusi yang bermakna bagi masyarakat.', {
    x: 1.2,
    y: 2.35,
    w: 7.8,
    h: 1.2,
    fontSize: 14,
    color: W,
    fontFace: 'Georgia',
    italic: true,
    align: 'center',
    valign: 'middle',
    wrap: true
  });
  s8.addShape(pptx.ShapeType.rect, { x: 3.8, y: 3.8, w: 2.4, h: 0.05, fill: { color: G } });
  s8.addText('Arkaan Hilmi Suharsoyo', { x: 1, y: 3.95, w: 8, h: 0.35, fontSize: 12, color: G, align: 'center', bold: true });
  s8.addText('Sistem Informasi · Institut Teknologi Sepuluh Nopember Surabaya', { x: 1, y: 4.35, w: 8, h: 0.28, fontSize: 9.5, color: 'A0B4D0', align: 'center' });
  s8.addText('Terima Kasih telah mengunjungi Portfolio ini', { x: 1, y: 5.2, w: 8, h: 0.32, fontSize: 11, color: 'A0B4D0', align: 'center' });
  s8.addText('📧 arkaan@email.com  •  📱 WhatsApp  •  💼 LinkedIn  •  🐙 GitHub', { x: 1, y: 5.65, w: 8, h: 0.28, fontSize: 9, color: '708090', align: 'center' });

  // Generate dan download file
  pptx.writeFile({ fileName: 'Portfolio_Arkaan_Hilmi_Suharsoyo.pptx' });
}
