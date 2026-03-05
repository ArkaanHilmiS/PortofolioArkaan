/* ====================================================
   ARKAAN HILMI SUHARSOYO - PORTFOLIO MAIN SCRIPT
   Information Systems · ITS Surabaya
   ==================================================== */

// ============================
// LANDING PAGE
// ============================
let selRole = '';

/**
 * Function to select visitor role
 * @param {string} r - Selected role
 * @param {HTMLElement} el - Clicked element
 */
function pickRole(r, el) {
  selRole = r;
  document.querySelectorAll('.role-card').forEach(c => c.classList.remove('sel'));
  el.classList.add('sel');
}

/**
 * Function to enter the main page
 */
function enter() {
  const name = document.getElementById('vname').value.trim();
  const role = selRole || 'Visitor';
  let w = '';

  // Personalized greeting based on role and name
  if (name) {
    w = `Hello, ${name}! Welcome 👋`;
  } else if (role === 'HR / Recruiter') {
    w = 'Hello! Welcome, esteemed Recruiter 👋';
  } else if (role === 'Student') {
    w = 'Hello, fellow Student! Welcome 👋';
  } else if (role === 'Colleague') {
    w = 'Hello, Colleague! Welcome 👋';
  } else if (role === 'Potential Client') {
    w = 'Hello! Thank you for visiting 👋';
  } else {
    w = "Welcome to Arkaan's Portfolio! 👋";
  }

  // Update welcome message and transition to main page
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

// Event listener for Enter key on name input
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
 * Animate skill progress bars
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
 * Initialize Intersection Observer for scroll reveal effect
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
 * Function to send message from contact form
 */
function sendMsg() {
  const n = document.getElementById('mn').value.trim();
  const m = document.getElementById('mm').value.trim();

  if (!n || !m) {
    alert('Please fill in your name and message first.');
    return;
  }

  // Display success message
  const ok = document.getElementById('ok-msg');
  ok.style.display = 'block';

  // Reset form
  document.getElementById('mn').value = '';
  document.getElementById('mc').value = '';
  document.getElementById('mm').value = '';
  document.getElementById('mp').value = '';

  // Hide success message after 5 seconds
  setTimeout(() => {
    ok.style.display = 'none';
  }, 5000);
}

// ============================
// DOWNLOAD PPT
// ============================

/**
 * Function to download portfolio as PowerPoint
 */
async function dlPPT() {
  const btn = document.querySelector('.btn-ppt');
  const prog = document.getElementById('ppt-prog');
  const fill = document.getElementById('pfill');
  const st = document.getElementById('ppt-status');

  btn.disabled = true;
  btn.innerHTML = '⏳ Processing…';
  prog.style.display = 'block';

  // Simulate progress with step-by-step updates
  const steps = [
    [10, 'Gathering profile data…'],
    [25, 'Creating Education slide…'],
    [45, 'Creating Experience slide…'],
    [62, 'Creating Projects slide…'],
    [78, 'Creating Skills slide…'],
    [92, 'Assembling Freelance & Closing slides…'],
    [100, 'Complete! Downloading…']
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
  btn.innerHTML = '✅ Successfully Downloaded!';
  setTimeout(() => {
    btn.innerHTML = '⬇️ Download Portfolio.pptx';
    prog.style.display = 'none';
    fill.style.width = '0';
  }, 3000);
}

/**
 * Function to generate and create the PowerPoint file
 */
async function makePPT() {
  // Load PptxGenJS library if not loaded yet
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
   * Helper function for dark background
   */
  function darkBg(sl) {
    sl.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: N } });
    sl.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.08, fill: { color: G } });
    sl.addShape(pptx.ShapeType.rect, { x: 0, y: 7.42, w: '100%', h: 0.08, fill: { color: G } });
  }

  /**
   * Helper function for light background
   */
  function lightBg(sl) {
    sl.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: LG } });
    sl.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.08, fill: { color: N } });
  }

  /**
   * Helper function for page numbers
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
  s1.addText('Master\'s Student in Information Systems', { x: 1, y: 2.85, w: 8, h: 0.35, fontSize: 13, color: 'A0B4D0', align: 'center' });
  s1.addText('Institut Teknologi Sepuluh Nopember · Surabaya', { x: 1, y: 3.22, w: 8, h: 0.3, fontSize: 11, color: '7A90A8', align: 'center' });
  s1.addText('🌐 Web Dev  •  📊 Data Analytics  •  🤖 Machine Learning  •  📸 Photography', { x: 1, y: 3.82, w: 8, h: 0.32, fontSize: 10, color: 'A0B4D0', align: 'center' });
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
    ['📧', 'arkaanhilmis23@gmail.com'],
    ['💼', 'LinkedIn'],
    ['🐙', 'GitHub'],
    ['📱', 'Instagram']
  ].forEach(([ic, lb], i) => {
    s2.addText(`${ic} ${lb}`, { x: 0.15, y: 4.2 + i * 0.4, w: 2.9, h: 0.32, fontSize: 9, color: 'A0B4D0', align: 'center' });
  });

  s2.addText('PROFILE', { x: 3.5, y: 0.3, w: 5.8, h: 0.35, fontSize: 11, bold: true, color: N, charSpacing: 4 });
  s2.addShape(pptx.ShapeType.rect, { x: 3.5, y: 0.7, w: 1, h: 0.05, fill: { color: G } });
  s2.addText(
    'A dynamic and results-driven professional from ITS Surabaya with expertise in business consulting, ERP systems, machine learning, and data visualization. Focused on delivering impactful solutions.',
    { x: 3.5, y: 0.9, w: 5.8, h: 0.95, fontSize: 10, color: DG, wrap: true }
  );

  [
    ['GPA', '3.66'],
    ['Team Led', '156'],
    ['KPI', '96%']
  ].forEach(([lb, vl], i) => {
    const xp = 3.5 + i * 2;
    s2.addShape(pptx.ShapeType.rect, { x: xp, y: 2.1, w: 1.8, h: 0.88, fill: { color: 'E8ECF4' }, line: { color: 'D0D8EC', width: 0.5 } });
    s2.addText(vl, { x: xp, y: 2.18, w: 1.8, h: 0.42, fontSize: 18, bold: true, color: N, align: 'center', fontFace: 'Georgia' });
    s2.addText(lb, { x: xp, y: 2.62, w: 1.8, h: 0.25, fontSize: 8, color: DG, align: 'center' });
  });

  s2.addText('FUTURE INTERESTS', { x: 3.5, y: 3.2, w: 5.8, h: 0.28, fontSize: 9, bold: true, color: N, charSpacing: 2 });
  ['🤖 AI / ML Engineer', '🌐 Full-Stack Dev', '📊 Data Scientist', '☁️ Cloud & DevOps', '📸 Creative Director', '🎓 Educator / Mentor'].forEach((it, i) => {
    const col = i % 3,
      row = Math.floor(i / 3);
    s2.addShape(pptx.ShapeType.rect, { x: 3.5 + col * 1.95, y: 3.6 + row * 0.55, w: 1.82, h: 0.42, fill: { color: W }, line: { color: 'D0D8EC', width: 0.5 } });
    s2.addText(it, { x: 3.5 + col * 1.95, y: 3.6 + row * 0.55, w: 1.82, h: 0.42, fontSize: 8, color: N, align: 'center', valign: 'middle', wrap: true });
  });

  // SLIDE 3 – EDUCATION
  const s3 = pptx.addSlide();
  lightBg(s3);
  pg(s3, 3);
  s3.addText('EDUCATION', { x: 0.5, y: 0.2, w: 9, h: 0.38, fontSize: 11, bold: true, color: N, charSpacing: 4 });
  s3.addShape(pptx.ShapeType.rect, { x: 0.5, y: 0.62, w: 1, h: 0.05, fill: { color: G } });

  [
    { logo: 'ITS', period: '2022 — 2026', school: 'Institut Teknologi Sepuluh Nopember', major: 'Information Systems · FTEIC', ipk: 'GPA 3.66', y: 0.9 },
    { logo: 'ITS', period: '2026 — Present', school: 'Institut Teknologi Sepuluh Nopember', major: 'Information Systems · FTEIC (Master)', ipk: 'GPA 3.75', y: 4.0 }
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
    s3.addText('Achievements: Information management systems, data analytics, Power BI dashboards, full-stack web applications using Next.js & Express.', {
      x: 1.75,
      y: ed.y + 1.65,
      w: 7.2,
      h: 0.6,
      fontSize: 9,
      color: DG,
      wrap: true
    });
  });

  // SLIDE 4 – EXPERIENCE
  const s4 = pptx.addSlide();
  lightBg(s4);
  pg(s4, 4);
  s4.addText('EXPERIENCE', { x: 0.5, y: 0.2, w: 9, h: 0.38, fontSize: 11, bold: true, color: N, charSpacing: 4 });
  s4.addShape(pptx.ShapeType.rect, { x: 0.5, y: 0.62, w: 1, h: 0.05, fill: { color: G } });

  [
    { icon: '🏛️', type: 'ORG', period: '2023—Now', title: 'Board Member HMSI', org: 'IS Student Association ITS', desc: 'Student affairs & event committees' },
    { icon: '💻', type: 'WORK', period: '2024—Now', title: 'Web Developer', org: 'Freelance', desc: 'Building modern web apps & sites' },
    { icon: '📸', type: 'ORG', period: '2022—2023', title: 'Photographer & Videographer', org: 'ITS Photography Club', desc: 'Visual content & campus documentation' },
    { icon: '📊', type: 'WORK', period: '2024', title: 'Data Analyst Intern', org: '[Company]', desc: 'Python, Power BI, analytics dashboards' },
    { icon: '📐', type: 'ORG', period: '2023', title: 'Mathematics Mentor', org: 'ITS Tutoring Program', desc: 'Mentoring junior students' },
    { icon: '🤝', type: 'ORG', period: '2022—2023', title: 'Event Committee', org: 'Various ITS Committees', desc: 'Orientation, seminars, competitions' }
  ].forEach((ex, i) => {
    const col = i % 3,
      row = Math.floor(i / 3),
      xp = 0.4 + col * 3.1,
      yp = 0.88 + row * 2.9;
    s4.addShape(pptx.ShapeType.rect, { x: xp, y: yp, w: 2.9, h: 2.6, fill: { color: W }, line: { color: 'D0D8EC', width: 0.5 } });
    const bc = ex.type === 'ORG' ? 'FDF8E1' : 'EFF6FF',
      btc = ex.type === 'ORG' ? '906020' : A;
    s4.addShape(pptx.ShapeType.rect, { x: xp + 1.6, y: yp + 0.12, w: 1.15, h: 0.25, fill: { color: bc }, line: { color: 'E0D0B0', width: 0.3 } });
    s4.addText(ex.type === 'ORG' ? 'ORGANIZATION' : 'WORK', {
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

  // SLIDE 5 – PROJECTS
  const s5 = pptx.addSlide();
  lightBg(s5);
  pg(s5, 5);
  s5.addText('PROJECTS & WORKS', { x: 0.5, y: 0.2, w: 9, h: 0.38, fontSize: 11, bold: true, color: N, charSpacing: 4 });
  s5.addShape(pptx.ShapeType.rect, { x: 0.5, y: 0.62, w: 1, h: 0.05, fill: { color: G } });

  [
    { icon: '🌐', name: 'Information Management System', tech: 'Next.js · React · TailwindCSS', desc: 'Full-stack web app with real-time analytics dashboard' },
    { icon: '📊', name: 'Power BI Dashboard', tech: 'Power BI · Python · Pandas', desc: 'Interactive data visualization for business decisions' },
    { icon: '🤖', name: 'Machine Learning Model', tech: 'Python · Scikit-learn · TensorFlow', desc: 'Predictive classification & regression on real-world data' },
    { icon: '📱', name: 'Portfolio Website', tech: 'HTML · CSS · JavaScript', desc: 'Responsive portfolio with animations & PPT export' },
    { icon: '📸', name: 'Creative Visual Content', tech: 'Photography · Videography · Editing', desc: 'Professional photo & video content production' },
    { icon: '📝', name: 'Research & Scientific Reports', tech: 'Analysis · Research · Visualization', desc: 'Academic reports with structured methodology' }
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

  // SLIDE 6 – SKILLS
  const s6 = pptx.addSlide();
  darkBg(s6);
  pg(s6, 6);
  s6.addText('SKILLS & CERTIFICATIONS', { x: 0.5, y: 0.2, w: 9, h: 0.38, fontSize: 11, bold: true, color: W, charSpacing: 4 });
  s6.addShape(pptx.ShapeType.rect, { x: 0.5, y: 0.62, w: 1, h: 0.05, fill: { color: G } });

  [
    { name: 'Leadership & Team Building', pct: 98 },
    { name: 'Information Systems Analysis', pct: 95 },
    { name: 'Business Intelligence', pct: 85 },
    { name: 'Power BI / Data Visualization', pct: 85 },
    { name: 'Machine Learning', pct: 78 },
    { name: 'Web Development', pct: 78 }
  ].forEach((sk, i) => {
    const yp = 0.95 + i * 0.72;
    s6.addText(sk.name, { x: 0.5, y: yp, w: 3.5, h: 0.28, fontSize: 9.5, color: W, bold: true });
    s6.addText(`${sk.pct}%`, { x: 4.05, y: yp, w: 0.5, h: 0.28, fontSize: 9, color: G, bold: true, fontFace: 'Courier New' });
    s6.addShape(pptx.ShapeType.rect, { x: 0.5, y: yp + 0.3, w: 4.5, h: 0.18, fill: { color: '1A3060' } });
    s6.addShape(pptx.ShapeType.rect, { x: 0.5, y: yp + 0.3, w: (4.5 * sk.pct) / 100, h: 0.18, fill: { color: A } });
  });

  s6.addText('TRAINING & CERTIFICATIONS', { x: 5.3, y: 0.75, w: 4.3, h: 0.28, fontSize: 9, bold: true, color: G, charSpacing: 2 });
  [
    '🌐 Web Development Bootcamp',
    '📊 Data Analytics Professional (Google)',
    '🤖 Machine Learning Fundamentals',
    '📈 Power BI Data Visualization',
    '🐍 Python for Data Science',
    '📸 Professional Photography & Videography'
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
  s7.addText('FREELANCE & REMOTE SERVICES', { x: 0.5, y: 0.58, w: 9, h: 0.4, fontSize: 14, bold: true, color: W, align: 'center', fontFace: 'Georgia' });
  s7.addText('Available for freelance and remote work across various fields of expertise', { x: 0.5, y: 1.0, w: 9, h: 0.28, fontSize: 9.5, color: 'A0B4D0', align: 'center' });

  [
    { icon: '🌐', name: 'Website Development' },
    { icon: '📐', name: 'Mathematics Tutoring' },
    { icon: '💻', name: 'Programming Tutoring' },
    { icon: '📝', name: 'Reports & Assignments' },
    { icon: '📊', name: 'Power BI Visualization' },
    { icon: '🤖', name: 'Machine Learning' },
    { icon: '📸', name: 'Photography' },
    { icon: '🎬', name: 'Videography' }
  ].forEach((fl, i) => {
    const col = i % 4,
      row = Math.floor(i / 4),
      xp = 0.4 + col * 2.3,
      yp = 1.48 + row * 2.2;
    s7.addShape(pptx.ShapeType.rect, { x: xp, y: yp, w: 2.15, h: 1.9, fill: { color: '0E2248' }, line: { color: '1A3060', width: 0.5 } });
    s7.addText(fl.icon, { x: xp, y: yp + 0.18, w: 2.15, h: 0.7, fontSize: 26, align: 'center' });
    s7.addText(fl.name, { x: xp + 0.1, y: yp + 1.02, w: 1.95, h: 0.62, fontSize: 9, bold: true, color: W, align: 'center', valign: 'middle', wrap: true });
  });

  s7.addText('📧 arkaanhilmis23@gmail.com  •  📱 Instagram  •  💼 LinkedIn  •  🐙 GitHub', { x: 0.5, y: 6.9, w: 9, h: 0.28, fontSize: 9, color: G, align: 'center' });

  // SLIDE 8 – PENUTUP
  const s8 = pptx.addSlide();
  darkBg(s8);
  s8.addShape(pptx.ShapeType.rect, { x: 0.5, y: 1.5, w: 9, h: 4.5, fill: { color: '0E2248' }, line: { color: '1A3060', width: 0.5 } });
  s8.addText('"', { x: 0.7, y: 1.55, w: 1, h: 1, fontSize: 80, color: '1E4080', fontFace: 'Georgia', bold: true });
  s8.addText('I believe that technology, creativity, and data can work in synergy to create meaningful solutions for society.', {
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
  s8.addText('Information Systems · Institut Teknologi Sepuluh Nopember Surabaya', { x: 1, y: 4.35, w: 8, h: 0.28, fontSize: 9.5, color: 'A0B4D0', align: 'center' });
  s8.addText('Thank you for visiting this Portfolio', { x: 1, y: 5.2, w: 8, h: 0.32, fontSize: 11, color: 'A0B4D0', align: 'center' });
  s8.addText('📧 arkaanhilmis23@gmail.com  •  📱 Instagram  •  💼 LinkedIn  •  🐙 GitHub', { x: 1, y: 5.65, w: 8, h: 0.28, fontSize: 9, color: '708090', align: 'center' });

  // Generate and download file
  pptx.writeFile({ fileName: 'Portfolio_Arkaan_Hilmi_Suharsoyo.pptx' });
}
