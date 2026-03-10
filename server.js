const express = require('express');
const path    = require('path');
const JSZip   = require('jszip');
const crypto  = require('crypto');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign
} = require('docx');

const app  = express();
const PORT = process.env.PORT || 3001;

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ─────────────────────────────────────────────────────────────────────────────
// AUTH — simple in-memory session store
// To add/change users: edit USERS below or set env vars on Render:
//   ADMIN_EMAIL and ADMIN_PASSWORD
// ─────────────────────────────────────────────────────────────────────────────
const USERS = [
  {
    email:    (process.env.ADMIN_EMAIL    || 'admin@tmosphere.in').toLowerCase(),
    password: (process.env.ADMIN_PASSWORD || 'tmosphere@2024'),
    name:     'Administrator'
  }
];

// Session tokens: Map<token, { email, name, expires }>
const SESSIONS = new Map();
const SESSION_TTL = 8 * 60 * 60 * 1000; // 8 hours

function createSession(email, name, remember) {
  const token   = crypto.randomBytes(32).toString('hex');
  const expires = Date.now() + (remember ? 30 * 24 * 60 * 60 * 1000 : SESSION_TTL);
  SESSIONS.set(token, { email, name, expires });
  return token;
}

function getSession(token) {
  if (!token) return null;
  const sess = SESSIONS.get(token);
  if (!sess) return null;
  if (Date.now() > sess.expires) { SESSIONS.delete(token); return null; }
  return sess;
}

// Auth middleware — reads token from Authorization header or cookie
function requireAuth(req, res, next) {
  const header = req.headers['authorization'] || '';
  const token  = header.startsWith('Bearer ') ? header.slice(7) : null;
  const sess   = getSession(token);
  if (!sess) return res.status(401).json({ error: 'Unauthorized. Please log in.' });
  req.user = sess;
  next();
}

// ─────────────────────────────────────────────────────────────────────────────
// CONSTANTS
// Title = 12pt bold (24 half-pts), Body = 11pt (22 half-pts), Single spacing
// ─────────────────────────────────────────────────────────────────────────────
const FONT       = 'Cambria';
const TITLE_SIZE = 24;   // 12pt — heading / title
const BODY_SIZE  = 22;   // 11pt — all body text
const LINE_RULE  = 240;  // exact single line spacing (twips)
const SA_PARA    = 80;   // spaceAfter between body paragraphs (twips ≈ 4pt)
const SA_SMALL   = 40;   // tight spacing
const SA_NONE    = 0;

const DESIGNATION_MAP = {
  'Private Limited' : 'DIRECTOR',
  'Public Limited'  : 'DIRECTOR',
  'HUF'             : 'KARTA',
  'Trust'           : 'TRUSTEE',
  'Proprietorship'  : 'PROPRIETOR',
  'Partnership'     : 'PARTNER',
  'LLP'             : 'DESIGNATED PARTNER'
};

// Auto-generate "10th day of March, 2026" from YYYY-MM-DD
function toWrittenDate(dateStr) {
  if (!dateStr) return '';
  const [y, m, d] = dateStr.split('-').map(Number);
  const ordinals = ['','1st','2nd','3rd','4th','5th','6th','7th','8th','9th','10th',
    '11th','12th','13th','14th','15th','16th','17th','18th','19th','20th',
    '21st','22nd','23rd','24th','25th','26th','27th','28th','29th','30th','31st'];
  const months = ['January','February','March','April','May','June',
    'July','August','September','October','November','December'];
  return ordinals[d] + ' day of ' + months[m - 1] + ', ' + y;
}

// ─────────────────────────────────────────────────────────────────────────────
// TEXT RUN HELPERS
// ─────────────────────────────────────────────────────────────────────────────
const R = (text, sz) =>
  new TextRun({ text: String(text || ''), font: FONT, size: sz || BODY_SIZE, bold: false });

const Rb = (text, sz) =>
  new TextRun({ text: String(text || ''), font: FONT, size: sz || BODY_SIZE, bold: true });

// Body paragraph — justified, single spacing, tight spaceAfter
const P = (runs, spaceAfter) =>
  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: spaceAfter !== undefined ? spaceAfter : SA_PARA, line: LINE_RULE, lineRule: 'exact' },
    children: Array.isArray(runs) ? runs : [runs],
  });

// Centered paragraph
const PC = (runs, spaceAfter) =>
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: spaceAfter !== undefined ? spaceAfter : SA_NONE, line: LINE_RULE, lineRule: 'exact' },
    children: Array.isArray(runs) ? runs : [runs],
  });

// Left paragraph
const PL = (runs, spaceAfter) =>
  new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { after: spaceAfter !== undefined ? spaceAfter : SA_PARA, line: LINE_RULE, lineRule: 'exact' },
    children: Array.isArray(runs) ? runs : [runs],
  });

// Numbered list item with hanging indent
const ListItem = (n, text) =>
  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: SA_SMALL, line: LINE_RULE, lineRule: 'exact' },
    indent: { left: 360, hanging: 360 },
    children: [R(n + '.  ' + text)],
  });

// Thin horizontal rule
const HRule = (spaceAfter) =>
  new Paragraph({
    spacing: { after: spaceAfter !== undefined ? spaceAfter : SA_PARA, line: LINE_RULE, lineRule: 'exact' },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: '000000', space: 2 } },
    children: [R('')],
  });

// ─────────────────────────────────────────────────────────────────────────────
// SIGNATURE TABLE — 4 rows × 2 cols, invisible borders
// ─────────────────────────────────────────────────────────────────────────────
function noBorder() {
  const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  return { top: none, bottom: none, left: none, right: none, insideH: none, insideV: none };
}

function sigCell(text, isBold) {
  return new TableCell({
    width: { size: 4680, type: WidthType.DXA },
    borders: noBorder(),
    verticalAlign: VerticalAlign.TOP,
    margins: { top: 0, bottom: 0, left: 0, right: 0 },
    children: [
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: 60, line: LINE_RULE, lineRule: 'exact' },
        children: [ isBold ? Rb(text) : R(text) ]
      })
    ]
  });
}

function sigTable(rows) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [4680, 4680],
    borders: noBorder(),
    rows: rows.map(([left, right]) =>
      new TableRow({
        children: [
          sigCell(left[0],  left[1]),
          sigCell(right[0], right[1]),
        ]
      })
    )
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// AFFIDAVIT  —  everything on ONE page
// ─────────────────────────────────────────────────────────────────────────────
function buildAffidavit(d) {
  const designation = DESIGNATION_MAP[d.businessType] || d.businessType.toUpperCase();
  const isProp      = d.businessType === 'Proprietorship';

  const openingRuns = isProp
    ? [ R('I '), Rb(d.applicantName), R(', Proprietor of "'), Rb(d.businessName),
        R('" having registered office at '), Rb(d.registeredAddress), R('.') ]
    : [ R('I '), Rb(d.applicantName), R(', '), Rb(designation),
        R(' of "'), Rb(d.businessName),
        R('" having registered office at '), Rb(d.registeredAddress), R('.') ];

  return new Document({
    styles: { default: { document: { run: { font: FONT, size: BODY_SIZE } } } },
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
        }
      },
      children: [
        PC([Rb('AFFIDAVIT', TITLE_SIZE)], SA_NONE),
        HRule(SA_PARA),

        P(openingRuns, SA_PARA),
        P([R('Do hereby solely affirm as follows:')], SA_PARA),
        P([R('That I am an Indian by nationality and residing at '), Rb(d.residentialAddress), R('.')], SA_PARA),
        P([R('I state that I am familiar and well conversant with the facts and circumstances ' +
             'of the present matters and competent and authorised to swear this affidavit and ' +
             'make the necessary statements in respect thereof.')], SA_PARA),
        // Trademark use paragraph — dynamic based on usageType
        ...((() => {
          const useType = d.usageType || 'proposed'; // 'used' or 'proposed'
          if (useType === 'used' && d.commencementDate) {
            // Already in use — show commencement date
            return [
              P([R('A trademark application is hereby made for registration of the accompanying trademark '),
                 Rb('"' + d.brandName + '"'), R(' in '), Rb('CLASS ' + d.businessClass),
                 R(' and the said mark is already in use for the said '),
                 Rb(d.businessType.toUpperCase()),
                 R('. The mark has been in use since '), Rb(d.commencementDate), R('.')], SA_PARA),
            ];
          } else {
            // Proposed to be used
            return [
              P([R('A trademark application is hereby made for registration of the accompanying trademark '),
                 Rb('"' + d.brandName + '"'), R(' in '), Rb('CLASS ' + d.businessClass),
                 R(' and the said mark has been proposed to be used for the said '),
                 Rb(d.businessType.toUpperCase()), R('.')], SA_PARA),
            ];
          }
        })()),
        P([R('I solemnly state that the content of this affidavit is true to the best of my ' +
             'knowledge and belief and that it conceals nothing and that no part is false.')], 160),

        PL([Rb(d.applicantName)], SA_SMALL),
        PL([R('DATE: ' + d.affidavitDate)], SA_SMALL),
        PL([R('PLACE: ' + d.place)], SA_NONE),
      ]
    }]
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// POWER OF ATTORNEY  —  everything on ONE page
// ─────────────────────────────────────────────────────────────────────────────
function buildPOA(d) {
  const designation = DESIGNATION_MAP[d.businessType] || d.businessType.toUpperCase();
  const isProp      = d.businessType === 'Proprietorship';
  const pronoun     = isProp ? 'I' : 'I/We';

  const openingRuns = isProp
    ? [ R('I '), Rb(d.applicantName), R(', Proprietor of "'), Rb(d.businessName),
        R('" having registered office at '), Rb(d.registeredAddress),
        R(', do hereby appoint: '), Rb(d.agentName),
        R(', having office at '), Rb(d.agentAddress),
        R(', and having '), Rb('Trade Marks Agent Code ' + d.agentCode), R('.') ]
    : [ R('I/We '), Rb(d.applicantName), R(', '), Rb(designation),
        R(' of "'), Rb(d.businessName),
        R('" having registered office at '), Rb(d.registeredAddress),
        R(', do hereby appoint: '), Rb(d.agentName),
        R(', having office at '), Rb(d.agentAddress),
        R(', and having '), Rb('Trade Marks Agent Code ' + d.agentCode), R('.') ];

  const items = [
    'Applying for registration of the following trademark(s) under the Trade Marks Act, 1999 and Rules made thereunder;',
    'Preparing, signing and submitting all applications, requests, forms, responses, and other documents;',
    'Representing me/us before the Registrar of Trade Marks or any other competent authority;',
    'Receiving and responding to all notices, objections, oppositions, and communications;',
    'Making necessary amendments or modifications to the application;',
    'Taking all necessary steps, including appearing at hearings, filing affidavits or appeals, and performing other acts, deeds and things which are necessary or incidental to the registration and protection of the said mark(s).'
  ];

  return new Document({
    styles: { default: { document: { run: { font: FONT, size: BODY_SIZE } } } },
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
        }
      },
      children: [
        // Title block
        PC([Rb('GENERAL POWER OF ATTORNEY FOR TRADEMARK/SERVICE MARK', TITLE_SIZE)], SA_NONE),
        PC([R('(UNDER SECTION 145 OF THE TRADEMARKS ACT, 1999)')], SA_NONE),
        HRule(SA_PARA),

        // Opening
        P(openingRuns, SA_PARA),

        // Scope heading
        P([Rb('As my/our lawful Attorney to act on my/our behalf in respect of:')], SA_SMALL),

        // Numbered list — tight spacing
        ...items.map((item, i) => ListItem(i + 1, item)),

        // Ratification
        new Paragraph({ spacing: { after: 120, before: 60, line: LINE_RULE, lineRule: 'exact' }, alignment: AlignmentType.JUSTIFIED,
          children: [ Rb(pronoun + ' hereby confirm and ratify all acts done by the above-mentioned attorney ' +
            'in pursuance of this authority executed on this '), Rb(d.poaExecutionDate), Rb('.') ] }),

        // Signature table
        sigTable([
          [ ['Signature of Applicant(s):', true],  ['Accepted by:',              true]  ],
          [ ['Name: ' + d.applicantName,   false], ['Name: ' + d.agentName,      false] ],
          [ ['Designation: ' + designation, true], ['Agent Code: ' + d.agentCode,false] ],
          [ ['For: ' + d.businessName,     false], ['',                           false] ],
        ]),

        // Address block — tight
        new Paragraph({ spacing: { after: SA_SMALL, before: 80, line: LINE_RULE, lineRule: 'exact' }, children: [Rb('To,')] }),
        PL([Rb('The Registrar of Trade Marks,')],           SA_SMALL),
        PL([Rb('The Office of the Trade Marks Registry at')], SA_SMALL),
        PL([Rb(d.tmOffice)],                                  SA_SMALL),
        PL([R('Date: ' + d.poaDate)],                       SA_NONE),
      ]
    }]
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// ROUTES
// ─────────────────────────────────────────────────────────────────────────────
app.get('/health', (req, res) => {
  res.json({ status: 'ok', app: 'tmOsphere', version: '3.1.0' });
});

// ── Login ──────────────────────────────────────────────────────────────────
app.post('/api/login', (req, res) => {
  const { email = '', password = '', remember = false } = req.body;
  const user = USERS.find(
    u => u.email === email.toLowerCase().trim() && u.password === password
  );
  if (!user) {
    return res.status(401).json({ error: 'Invalid email or password.' });
  }
  const token = createSession(user.email, user.name, remember);
  res.json({ success: true, token, name: user.name });
});

// ── Logout ─────────────────────────────────────────────────────────────────
app.post('/api/logout', (req, res) => {
  const header = req.headers['authorization'] || '';
  const token  = header.startsWith('Bearer ') ? header.slice(7) : null;
  if (token) SESSIONS.delete(token);
  res.json({ success: true });
});

// ── Session check ───────────────────────────────────────────────────────────
app.get('/api/me', requireAuth, (req, res) => {
  res.json({ email: req.user.email, name: req.user.name });
});

app.post('/api/generate', requireAuth, async (req, res) => {
  try {
    const d = req.body;

    const required = [
      'applicantName', 'businessName', 'brandName',
      'registeredAddress', 'residentialAddress',
      'businessClass', 'businessType', 'affidavitDate', 'place',
      'agentName', 'agentCode', 'agentAddress',
      'poaDate', 'tmOffice'
    ];
    const missing = required.filter(f => !d[f] || !String(d[f]).trim());
    if (missing.length) {
      return res.status(400).json({ error: 'Missing required fields: ' + missing.join(', ') });
    }

    // Auto-generate written execution date from poaDate e.g. "10th day of March, 2026"
    d.poaExecutionDate = toWrittenDate(d.poaDate);

    const [affBuf, poaBuf] = await Promise.all([
      Packer.toBuffer(buildAffidavit(d)),
      Packer.toBuffer(buildPOA(d))
    ]);

    const safeName = (d.brandName || 'TM').replace(/[^a-zA-Z0-9_-]/g, '_');
    const zip = new JSZip();
    zip.file('Affidavit_' + safeName + '.docx', affBuf);
    zip.file('POA_' + safeName + '.docx', poaBuf);
    const zipBuf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });

    res.set({
      'Content-Type': 'application/zip',
      'Content-Disposition': `attachment; filename="TM_Documents_${safeName}.zip"`
    });
    res.send(zipBuf);

  } catch (err) {
    console.error('[generate error]', err);
    res.status(500).json({ error: err.message });
  }
});

app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
  console.log('✅  tmOsphere v4.0 running → http://localhost:' + PORT);
});
