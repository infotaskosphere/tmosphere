const express = require('express');
const path    = require('path');
const JSZip   = require('jszip');
const crypto  = require('crypto');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign
} = require('docx');

let PDFDocument;
try { PDFDocument = require('pdfkit'); } catch(e) { PDFDocument = null; }

const app  = express();
const PORT = process.env.PORT || 3001;

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ─────────────────────────────────────────────────────────────────────────────
// AUTH — HMAC-signed tokens (survive server restarts)
// ─────────────────────────────────────────────────────────────────────────────
const USERS = [
  {
    email:    (process.env.ADMIN_EMAIL    || 'admin@tmosphere.in').toLowerCase(),
    password: (process.env.ADMIN_PASSWORD || 'tmosphere@2024'),
    name:     'Administrator'
  }
];

const APP_SECRET = process.env.APP_SECRET
  || process.env.JWT_SECRET
  || 'tmOsphere_2025@Render#SecureKey$TM_India_TradeMarks_Act_1999!XyZ9q2mK';

const SESSION_TTL      = 8  * 60 * 60 * 1000;
const SESSION_TTL_LONG = 30 * 24 * 60 * 60 * 1000;

function createToken(email, name, remember) {
  const expires = Date.now() + (remember ? SESSION_TTL_LONG : SESSION_TTL);
  const payload = Buffer.from(JSON.stringify({ email, name, expires })).toString('base64url');
  const sig     = crypto.createHmac('sha256', APP_SECRET).update(payload).digest('base64url');
  return payload + '.' + sig;
}

function verifyToken(token) {
  if (!token || typeof token !== 'string') return null;
  const dot = token.lastIndexOf('.');
  if (dot === -1) return null;
  const payload  = token.slice(0, dot);
  const sig      = token.slice(dot + 1);
  const expected = crypto.createHmac('sha256', APP_SECRET).update(payload).digest('base64url');
  if (sig.length !== expected.length) return null;
  try { if (!crypto.timingSafeEqual(Buffer.from(sig), Buffer.from(expected))) return null; } catch { return null; }
  let data;
  try { data = JSON.parse(Buffer.from(payload, 'base64url').toString()); } catch { return null; }
  if (!data || !data.email || !data.expires) return null;
  if (Date.now() > data.expires) return null;
  return { email: data.email, name: data.name };
}

function requireAuth(req, res, next) {
  const header = req.headers['authorization'] || '';
  const token  = header.startsWith('Bearer ') ? header.slice(7) : null;
  const sess   = verifyToken(token);
  if (!sess) return res.status(401).json({ error: 'Unauthorized. Please log in.' });
  req.user = sess;
  next();
}

// ─────────────────────────────────────────────────────────────────────────────
// DOCUMENT CONSTANTS
// Font: Times New Roman | Heading: 14pt Bold Centered | Body: 12pt Justified
// ─────────────────────────────────────────────────────────────────────────────
const FONT        = 'Times New Roman';
const TITLE_SIZE  = 28;   // 14pt in half-points
const SUB_SIZE    = 22;   // 11pt in half-points
const BODY_SIZE   = 24;   // 12pt in half-points
const LINE_RULE   = 240;  // exact single line spacing (twips)
const SA_NONE     = 0;
const SA_SMALL    = 60;
const SA_PARA     = 120;
const SA_SECTION  = 180;

const DESIGNATION_MAP = {
  'Private Limited' : 'DIRECTOR',
  'Public Limited'  : 'DIRECTOR',
  'HUF'             : 'KARTA',
  'Trust'           : 'TRUSTEE',
  'Proprietorship'  : 'PROPRIETOR',
  'Partnership'     : 'PARTNER',
  'LLP'             : 'DESIGNATED PARTNER'
};

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

function toDisplayDate(dateStr) {
  if (!dateStr) return dateStr || '';
  const parts = dateStr.split('-');
  if (parts.length !== 3) return dateStr;
  return parts[2] + '-' + parts[1] + '-' + parts[0];
}

// ─────────────────────────────────────────────────────────────────────────────
// DOCX HELPERS
// ─────────────────────────────────────────────────────────────────────────────
const R  = (text, sz) => new TextRun({ text: String(text ?? ''), font: FONT, size: sz || BODY_SIZE, bold: false });
const Rb = (text, sz) => new TextRun({ text: String(text ?? ''), font: FONT, size: sz || BODY_SIZE, bold: true  });

const P  = (runs, sa, sb) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  spacing: { after: sa !== undefined ? sa : SA_PARA, before: sb || 0, line: LINE_RULE, lineRule: 'exact' },
  children: Array.isArray(runs) ? runs : [runs],
});
const PC = (runs, sa, sb) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: sa !== undefined ? sa : SA_NONE, before: sb || 0, line: LINE_RULE, lineRule: 'exact' },
  children: Array.isArray(runs) ? runs : [runs],
});
const PL = (runs, sa, sb) => new Paragraph({
  alignment: AlignmentType.LEFT,
  spacing: { after: sa !== undefined ? sa : SA_SMALL, before: sb || 0, line: LINE_RULE, lineRule: 'exact' },
  children: Array.isArray(runs) ? runs : [runs],
});

// Empty line spacer
const BLANK = () => new Paragraph({
  spacing: { after: 0, before: 0, line: LINE_RULE, lineRule: 'exact' },
  children: [new TextRun({ text: '', font: FONT, size: BODY_SIZE })],
});

// Horizontal rule line
const HRule = (sa) => new Paragraph({
  spacing: { after: sa !== undefined ? sa : SA_PARA, line: LINE_RULE, lineRule: 'exact' },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000', space: 1 } },
  children: [R('')],
});

// Numbered list item — hanging indent matching sample
const ListItem = (n, text) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  spacing: { after: SA_SMALL, line: LINE_RULE, lineRule: 'exact' },
  indent: { left: 360, hanging: 360 },
  children: [R(n + '.  ' + text)],
});

// ─────────────────────────────────────────────────────────────────────────────
// SIGNATURE TABLE — invisible borders, exactly 2 equal columns
// ─────────────────────────────────────────────────────────────────────────────
function noBorder() {
  const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  return { top: none, bottom: none, left: none, right: none, insideH: none, insideV: none };
}

function sigCell(paraChildren) {
  // paraChildren = array of TextRun arrays (one per row in the cell)
  return new TableCell({
    width: { size: 4680, type: WidthType.DXA },
    borders: noBorder(),
    verticalAlign: VerticalAlign.TOP,
    margins: { top: 0, bottom: 0, left: 0, right: 0 },
    children: paraChildren.map(runs => new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { after: 60, line: LINE_RULE, lineRule: 'exact' },
      children: Array.isArray(runs) ? runs : [runs],
    }))
  });
}

function buildSigTable(d, designation) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [4680, 4680],
    borders: noBorder(),
    rows: [
      new TableRow({ children: [
        sigCell([[Rb('Signature of Applicant(s):')]]),
        sigCell([[Rb('Accepted by:')]]),
      ]}),
      new TableRow({ children: [
        sigCell([[R('Name: '), Rb(d.applicantName)]]),
        sigCell([[R('Name: '), R(d.agentName)]]),
      ]}),
      new TableRow({ children: [
        sigCell([[Rb('Designation: ' + designation)]]),
        sigCell([[R('Agent Code: ' + d.agentCode)]]),
      ]}),
      new TableRow({ children: [
        sigCell([[R('For: ' + d.businessName)]]),
        sigCell([[R('')]]),
      ]}),
    ]
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// BUILD AFFIDAVIT DOCX
// ─────────────────────────────────────────────────────────────────────────────
function buildAffidavitDoc(d) {
  const designation = DESIGNATION_MAP[d.businessType] || d.businessType.toUpperCase();
  const isProp      = d.businessType === 'Proprietorship';

  const openingRuns = isProp
    ? [R('I '),Rb(d.applicantName),R(', Proprietor of "'),Rb(d.businessName),
       R('" having registered office at '),Rb(d.registeredAddress),R('.')]
    : [R('I '),Rb(d.applicantName),R(', '),Rb(designation),
       R(' of "'),Rb(d.businessName),
       R('" having registered office at '),Rb(d.registeredAddress),R('.')];

  const usePara = (d.usageType === 'used' && d.commencementDate)
    ? P([R('A trademark application is hereby made for registration of the accompanying trademark '),
         Rb('"'+d.brandName+'"'),R(' in '),Rb('CLASS '+d.businessClass),
         R(' and the said mark is already in use for the said '),
         Rb(d.businessType.toUpperCase()),
         R('. The mark has been in use since '),Rb(toDisplayDate(d.commencementDate)),R('.')], SA_PARA)
    : P([R('A trademark application is hereby made for registration of the accompanying trademark '),
         Rb('"'+d.brandName+'"'),R(' in '),Rb('CLASS '+d.businessClass),
         R(' and the said mark has been proposed to be used for the said '),
         Rb(d.businessType.toUpperCase()),R('.')], SA_PARA);

  // 10 blank lines to push signature to bottom — matching sample
  const spacers = Array.from({length:10}, () => BLANK());

  return new Document({
    styles: { default: { document: { run: { font: FONT, size: BODY_SIZE } } } },
    sections: [{
      properties: {
        page: {
          size:   { width: 12240, height: 15840 },
          margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
        }
      },
      children: [
        PC([Rb('AFFIDAVIT', TITLE_SIZE)], SA_NONE),
        HRule(SA_PARA),
        P(openingRuns, SA_PARA),
        P([R('Do hereby solely affirm as follows:')], SA_PARA),
        P([R('That I am an Indian by nationality and residing at '),Rb(d.residentialAddress),R('.')], SA_PARA),
        P([R('I state that I am familiar and well conversant with the facts and circumstances of the present matters and competent and authorised to swear this affidavit and make the necessary statements in respect thereof.')], SA_PARA),
        usePara,
        P([R('I solemnly state that the content of this affidavit is true to the best of my knowledge and belief and that it conceals nothing and that no part is false.')], SA_SECTION),
        ...spacers,
        PL([Rb(d.applicantName)], SA_SMALL),
        PL([R('DATE: '+toDisplayDate(d.affidavitDate))], SA_SMALL),
        PL([R('PLACE: '+d.place)], SA_NONE),
      ]
    }]
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// BUILD POA DOCX
// ─────────────────────────────────────────────────────────────────────────────
function buildPOADoc(d) {
  const designation = DESIGNATION_MAP[d.businessType] || d.businessType.toUpperCase();
  const isProp      = d.businessType === 'Proprietorship';
  const pronoun     = isProp ? 'I' : 'I/We';

  const openingRuns = isProp
    ? [R('I '),Rb(d.applicantName),R(', Proprietor of "'),Rb(d.businessName),
       R('" having registered office at '),Rb(d.registeredAddress),
       R(', do hereby appoint: '),Rb(d.agentName),
       R(', having office at '),Rb(d.agentAddress),
       R(', and having '),Rb('Trade Marks Agent Code '+d.agentCode),R('.')]
    : [R('I/We '),Rb(d.applicantName),R(', '),Rb(designation),
       R(' of "'),Rb(d.businessName),
       R('" having registered office at '),Rb(d.registeredAddress),
       R(', do hereby appoint: '),Rb(d.agentName),
       R(', having office at '),Rb(d.agentAddress),
       R(', and having '),Rb('Trade Marks Agent Code '+d.agentCode),R('.')];

  const items = [
    'Applying for registration of the following trademark(s) under the Trade Marks Act, 1999 and Rules made thereunder;',
    'Preparing, signing and submitting all applications, requests, forms, responses, and other documents;',
    'Representing me/us before the Registrar of Trade Marks or any other competent authority;',
    'Receiving and responding to all notices, objections, oppositions, and communications;',
    'Making necessary amendments or modifications to the application;',
    'Taking all necessary steps, including appearing at hearings, filing affidavits or appeals, and performing other acts, deeds and things which are necessary or incidental to the registration and protection of the said mark(s).'
  ];

  // Spacers to push signature block down — matching sample's large gap
  const spacers = Array.from({length:10}, () => BLANK());

  return new Document({
    styles: { default: { document: { run: { font: FONT, size: BODY_SIZE } } } },
    sections: [{
      properties: {
        page: {
          size:   { width: 12240, height: 15840 },
          margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
        }
      },
      children: [
        PC([Rb('GENERAL POWER OF ATTORNEY FOR TRADEMARK/SERVICE MARK', TITLE_SIZE)], SA_NONE),
        PC([R('(UNDER SECTION 145 OF THE TRADEMARKS ACT, 1999)', SUB_SIZE)], SA_NONE),
        HRule(SA_PARA),
        P(openingRuns, SA_PARA),
        P([Rb('As my/our lawful Attorney to act on my/our behalf in respect of:')], SA_SMALL),
        ...items.map((item, i) => ListItem(i+1, item)),
        P([Rb(pronoun+' hereby confirm and ratify all acts done by the above-mentioned attorney in pursuance of this authority executed on this '),
           Rb(d.poaExecutionDate),Rb('.')], SA_SECTION, SA_SMALL),
        ...spacers,
        buildSigTable(d, designation),
        BLANK(), BLANK(), BLANK(),
        PL([Rb('To,')], SA_SMALL),
        PL([Rb('The Registrar of Trade Marks,')], SA_SMALL),
        PL([Rb('The Office of the Trade Marks Registry at')], SA_SMALL),
        PL([Rb(d.tmOffice)], SA_SMALL),
        PL([R('Date: '+toDisplayDate(d.poaDate))], SA_NONE),
      ]
    }]
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// BUILD PDF using pdfkit — exact layout matching sample documents
// ─────────────────────────────────────────────────────────────────────────────
function buildPDF(d, type) {
  return new Promise((resolve, reject) => {
    if (!PDFDocument) {
      return reject(new Error('pdfkit not installed. Run: npm install pdfkit'));
    }

    const doc = new PDFDocument({
      size: 'LETTER',
      margins: { top: 72, right: 72, bottom: 72, left: 72 },
    });

    const chunks = [];
    doc.on('data',  chunk => chunks.push(chunk));
    doc.on('end',   ()    => resolve(Buffer.concat(chunks)));
    doc.on('error', err   => reject(err));

    const designation = DESIGNATION_MAP[d.businessType] || d.businessType.toUpperCase();
    const isProp      = d.businessType === 'Proprietorship';
    const pronoun     = isProp ? 'I' : 'I/We';

    const LEFT   = 72;
    const RIGHT  = doc.page.width - 72;
    const W      = RIGHT - LEFT;
    const H_SIZE = 14;
    const S_SIZE = 11;
    const B_SIZE = 12;
    const LH     = B_SIZE * 1.5;

    const FONT_R = 'Times-Roman';
    const FONT_B = 'Times-Bold';

    // ── Write a paragraph with mixed bold/normal runs, justified ──
    function writeMixed(segments, opts = {}) {
      const fs    = opts.fontSize || B_SIZE;
      const align = opts.align    || 'justify';
      let first   = true;
      const total = segments.length;
      segments.forEach((seg, i) => {
        if (!seg.text) return;
        doc.font(seg.bold ? FONT_B : FONT_R).fontSize(fs);
        const isLast = (i === total - 1);
        doc.text(seg.text, first ? undefined : undefined, first ? undefined : undefined, {
          continued: !isLast,
          align,
          lineGap: 1,
        });
        first = false;
      });
      if (opts.moveDown !== false) doc.moveDown(0.5);
    }

    function writeNormal(text, opts = {}) {
      doc.font(FONT_R).fontSize(opts.fontSize || B_SIZE);
      doc.text(text, { align: opts.align || 'justify', lineGap: 1 });
      if (opts.moveDown !== false) doc.moveDown(0.4);
    }

    function writeBold(text, opts = {}) {
      doc.font(FONT_B).fontSize(opts.fontSize || B_SIZE);
      doc.text(text, { align: opts.align || 'justify', lineGap: 1 });
      if (opts.moveDown !== false) doc.moveDown(0.4);
    }

    function writeHRule() {
      doc.moveDown(0.2);
      const y = doc.y;
      doc.moveTo(LEFT, y).lineTo(RIGHT, y).lineWidth(0.75).stroke();
      doc.moveDown(0.5);
    }

    function writeListItem(n, text) {
      const numW   = 24;
      const textX  = LEFT + numW;
      const textW  = W - numW;
      const startY = doc.y;
      doc.font(FONT_R).fontSize(B_SIZE);
      // Write number
      doc.text(n + '.', LEFT, startY, { width: numW, align: 'left', lineGap: 1 });
      // Write text beside it
      doc.text(text, textX, startY, { width: textW, align: 'justify', lineGap: 1 });
      doc.moveDown(0.2);
    }

    // ── AFFIDAVIT ──────────────────────────────────────────────────
    if (type === 'affidavit') {

      writeBold('AFFIDAVIT', { align: 'center', fontSize: H_SIZE, moveDown: false });
      doc.moveDown(0.3);
      writeHRule();

      if (isProp) {
        writeMixed([
          {text:'I '},{text:d.applicantName,bold:true},{text:', Proprietor of "'},
          {text:d.businessName,bold:true},{text:'" having registered office at '},
          {text:d.registeredAddress,bold:true},{text:'.'}
        ]);
      } else {
        writeMixed([
          {text:'I '},{text:d.applicantName,bold:true},{text:', '},{text:designation,bold:true},
          {text:' of "'},{text:d.businessName,bold:true},{text:'" having registered office at '},
          {text:d.registeredAddress,bold:true},{text:'.'}
        ]);
      }

      writeNormal('Do hereby solely affirm as follows:');

      writeMixed([
        {text:'That I am an Indian by nationality and residing at '},
        {text:d.residentialAddress,bold:true},{text:'.'}
      ]);

      writeNormal('I state that I am familiar and well conversant with the facts and circumstances of the present matters and competent and authorised to swear this affidavit and make the necessary statements in respect thereof.');

      if (d.usageType === 'used' && d.commencementDate) {
        writeMixed([
          {text:'A trademark application is hereby made for registration of the accompanying trademark '},
          {text:'"'+d.brandName+'"',bold:true},{text:' in '},{text:'CLASS '+d.businessClass,bold:true},
          {text:' and the said mark is already in use for the said '},
          {text:d.businessType.toUpperCase(),bold:true},
          {text:'. The mark has been in use since '},{text:toDisplayDate(d.commencementDate),bold:true},{text:'.'}
        ]);
      } else {
        writeMixed([
          {text:'A trademark application is hereby made for registration of the accompanying trademark '},
          {text:'"'+d.brandName+'"',bold:true},{text:' in '},{text:'CLASS '+d.businessClass,bold:true},
          {text:' and the said mark has been proposed to be used for the said '},
          {text:d.businessType.toUpperCase(),bold:true},{text:'.'}
        ]);
      }

      writeNormal('I solemnly state that the content of this affidavit is true to the best of my knowledge and belief and that it conceals nothing and that no part is false.');

      // Push signature to bottom of page
      const sigY = doc.page.height - 72 - (LH * 4.5);
      if (doc.y < sigY) doc.y = sigY;

      doc.font(FONT_B).fontSize(B_SIZE).text(d.applicantName, LEFT, doc.y, { align: 'left', lineGap: 1 });
      doc.moveDown(0.15);
      doc.font(FONT_R).fontSize(B_SIZE).text('DATE: '+toDisplayDate(d.affidavitDate), { align: 'left', lineGap: 1 });
      doc.moveDown(0.15);
      doc.font(FONT_R).fontSize(B_SIZE).text('PLACE: '+d.place, { align: 'left', lineGap: 1 });

    // ── POA ───────────────────────────────────────────────────────
    } else {

      writeBold('GENERAL POWER OF ATTORNEY FOR TRADEMARK/SERVICE MARK', { align: 'center', fontSize: H_SIZE, moveDown: false });
      doc.moveDown(0.25);
      writeNormal('(UNDER SECTION 145 OF THE TRADEMARKS ACT, 1999)', { align: 'center', fontSize: S_SIZE, moveDown: false });
      doc.moveDown(0.3);
      writeHRule();

      if (isProp) {
        writeMixed([
          {text:'I '},{text:d.applicantName,bold:true},{text:', Proprietor of "'},
          {text:d.businessName,bold:true},{text:'" having registered office at '},
          {text:d.registeredAddress,bold:true},{text:', do hereby appoint: '},
          {text:d.agentName,bold:true},{text:', having office at '},
          {text:d.agentAddress,bold:true},{text:', and having '},
          {text:'Trade Marks Agent Code '+d.agentCode,bold:true},{text:'.'}
        ]);
      } else {
        writeMixed([
          {text:'I/We '},{text:d.applicantName,bold:true},{text:', '},{text:designation,bold:true},
          {text:' of "'},{text:d.businessName,bold:true},{text:'" having registered office at '},
          {text:d.registeredAddress,bold:true},{text:', do hereby appoint: '},
          {text:d.agentName,bold:true},{text:', having office at '},
          {text:d.agentAddress,bold:true},{text:', and having '},
          {text:'Trade Marks Agent Code '+d.agentCode,bold:true},{text:'.'}
        ]);
      }

      writeBold('As my/our lawful Attorney to act on my/our behalf in respect of:');

      const items = [
        'Applying for registration of the following trademark(s) under the Trade Marks Act, 1999 and Rules made thereunder;',
        'Preparing, signing and submitting all applications, requests, forms, responses, and other documents;',
        'Representing me/us before the Registrar of Trade Marks or any other competent authority;',
        'Receiving and responding to all notices, objections, oppositions, and communications;',
        'Making necessary amendments or modifications to the application;',
        'Taking all necessary steps, including appearing at hearings, filing affidavits or appeals, and performing other acts, deeds and things which are necessary or incidental to the registration and protection of the said mark(s).'
      ];
      items.forEach((item, i) => writeListItem(i+1, item));

      doc.moveDown(0.3);
      writeMixed([
        {text:pronoun+' hereby confirm and ratify all acts done by the above-mentioned attorney in pursuance of this authority executed on this ',bold:true},
        {text:d.poaExecutionDate,bold:true},{text:'.',bold:true}
      ]);

      // Push signature table to bottom — large gap matching sample
      const addrBlockH = LH * 8;
      const sigTableH  = LH * 5;
      const sigStartY  = doc.page.height - 72 - addrBlockH - sigTableH;
      if (doc.y < sigStartY) doc.y = sigStartY;

      // Signature table — two columns
      const colW  = W / 2;
      const tY    = doc.y;

      doc.font(FONT_B).fontSize(B_SIZE);
      doc.text('Signature of Applicant(s):',  LEFT,         tY,          { width: colW, align: 'left', lineGap: 1 });
      doc.text('Accepted by:',                LEFT + colW,  tY,          { width: colW, align: 'left', lineGap: 1 });

      const r2 = tY + LH;
      doc.font(FONT_R).fontSize(B_SIZE);
      doc.text('Name: '+d.applicantName,  LEFT,        r2, { width: colW, align: 'left', lineGap: 1 });
      doc.text('Name: '+d.agentName,      LEFT+colW,   r2, { width: colW, align: 'left', lineGap: 1 });

      const r3 = r2 + LH;
      doc.font(FONT_B).fontSize(B_SIZE);
      doc.text('Designation: '+designation, LEFT,      r3, { width: colW, align: 'left', lineGap: 1 });
      doc.font(FONT_R).fontSize(B_SIZE);
      doc.text('Agent Code: '+d.agentCode,  LEFT+colW, r3, { width: colW, align: 'left', lineGap: 1 });

      const r4 = r3 + LH;
      doc.font(FONT_R).fontSize(B_SIZE);
      doc.text('For: '+d.businessName, LEFT, r4, { width: colW, align: 'left', lineGap: 1 });

      // Address block
      doc.y = r4 + LH * 2.2;

      doc.font(FONT_B).fontSize(B_SIZE).text('To,',                                       { align: 'left', lineGap: 1 });
      doc.font(FONT_B).fontSize(B_SIZE).text('The Registrar of Trade Marks,',             { align: 'left', lineGap: 1 });
      doc.font(FONT_B).fontSize(B_SIZE).text('The Office of the Trade Marks Registry at', { align: 'left', lineGap: 1 });
      doc.font(FONT_B).fontSize(B_SIZE).text(d.tmOffice,                                  { align: 'left', lineGap: 1 });
      doc.font(FONT_R).fontSize(B_SIZE).text('Date: '+toDisplayDate(d.poaDate),           { align: 'left', lineGap: 1 });
    }

    doc.end();
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// ROUTES
// ─────────────────────────────────────────────────────────────────────────────
app.get('/health', (req, res) => {
  res.json({ status: 'ok', app: 'tmOsphere', version: '5.0.0' });
});

app.post('/api/login', (req, res) => {
  const { email = '', password = '', remember = false } = req.body;
  const user = USERS.find(u => u.email === email.toLowerCase().trim() && u.password === password);
  if (!user) return res.status(401).json({ error: 'Invalid email or password.' });
  const token = createToken(user.email, user.name, remember);
  res.json({ success: true, token, name: user.name });
});

app.post('/api/logout', (req, res) => {
  res.json({ success: true });
});

app.get('/api/me', requireAuth, (req, res) => {
  res.json({ email: req.user.email, name: req.user.name });
});

app.post('/api/generate', requireAuth, async (req, res) => {
  try {
    const d = req.body;

    const required = [
      'applicantName','businessName','brandName',
      'registeredAddress','residentialAddress',
      'businessClass','businessType','affidavitDate','place',
      'agentName','agentCode','agentAddress',
      'poaDate','tmOffice'
    ];
    const missing = required.filter(f => !d[f] || !String(d[f]).trim());
    if (missing.length) {
      return res.status(400).json({ error: 'Missing required fields: ' + missing.join(', ') });
    }

    d.poaExecutionDate = toWrittenDate(d.poaDate);

    // Generate DOCX
    const [affDocx, poaDocx] = await Promise.all([
      Packer.toBuffer(buildAffidavitDoc(d)),
      Packer.toBuffer(buildPOADoc(d))
    ]);

    // Generate PDF
    const [affPdf, poaPdf] = await Promise.all([
      buildPDF(d, 'affidavit'),
      buildPDF(d, 'poa')
    ]);

    const safeName = (d.brandName || 'TM').replace(/[^a-zA-Z0-9_-]/g, '_');
    const zip = new JSZip();
    zip.file('Affidavit_' + safeName + '.docx', affDocx);
    zip.file('POA_'       + safeName + '.docx', poaDocx);
    zip.file('Affidavit_' + safeName + '.pdf',  affPdf);
    zip.file('POA_'       + safeName + '.pdf',  poaPdf);

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
  console.log('✅  tmOsphere v5.0 running → http://localhost:' + PORT);
});
