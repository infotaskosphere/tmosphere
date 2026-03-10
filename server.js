const express    = require('express');
const path       = require('path');
const JSZip      = require('jszip');
const crypto     = require('crypto');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  TabStopPosition, TabStopType, BorderStyle,
  WidthType, UnderlineType
} = require('docx');

let PDFDocument;
try { PDFDocument = require('pdfkit'); } catch(e) { PDFDocument = null; }

const app  = express();
const PORT = process.env.PORT || 3001;
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ─────────────────────────────────────────────────────────────────────────────
// AUTH — HMAC signed tokens (survive restarts)
// ─────────────────────────────────────────────────────────────────────────────
const USERS = [{
  email:    (process.env.ADMIN_EMAIL    || 'admin@tmosphere.in').toLowerCase(),
  password: (process.env.ADMIN_PASSWORD || 'tmosphere@2024'),
  name:     'Administrator'
}];

const APP_SECRET = process.env.APP_SECRET || process.env.JWT_SECRET
  || 'tmOsphere_2025@Render#SecureKey$TM_India_TradeMarks_Act_1999!XyZ9q2mK';

function createToken(email, name, remember) {
  const expires = Date.now() + (remember ? 30*24*60*60*1000 : 8*60*60*1000);
  const payload = Buffer.from(JSON.stringify({ email, name, expires })).toString('base64url');
  const sig     = crypto.createHmac('sha256', APP_SECRET).update(payload).digest('base64url');
  return payload + '.' + sig;
}
function verifyToken(token) {
  if (!token) return null;
  const dot = token.lastIndexOf('.');
  if (dot === -1) return null;
  const payload = token.slice(0, dot), sig = token.slice(dot + 1);
  const exp = crypto.createHmac('sha256', APP_SECRET).update(payload).digest('base64url');
  if (sig.length !== exp.length) return null;
  try { if (!crypto.timingSafeEqual(Buffer.from(sig), Buffer.from(exp))) return null; } catch { return null; }
  let data; try { data = JSON.parse(Buffer.from(payload, 'base64url').toString()); } catch { return null; }
  if (!data?.email || !data?.expires || Date.now() > data.expires) return null;
  return { email: data.email, name: data.name };
}
function requireAuth(req, res, next) {
  const h = req.headers['authorization'] || '';
  const sess = verifyToken(h.startsWith('Bearer ') ? h.slice(7) : null);
  if (!sess) return res.status(401).json({ error: 'Unauthorized. Please log in.' });
  req.user = sess; next();
}

// ─────────────────────────────────────────────────────────────────────────────
// EXACT FORMATTING — reverse-engineered from uploaded sample docs
//
// AFFIDAVIT:
//   Font:         Cambria (body), inherited from Word theme
//   Title:        14pt Bold, Justified (matches sample)
//   Body:         11pt, Justified
//   Spacing:      Word default — after=160 twips, line=259/auto
//   Blank lines:  12 empty paragraphs between body and signature
//   Sig block:    AUTHORISED SIGNATORY (bold, left)
//                 DATE: xx  <tabs>  NAME (bold, justified with tabs)
//                 PLACE: xx <tabs>  SOLE PROPRIETOR (bold)
//
// POA:
//   Font:         Cambria
//   Title:        14pt Bold, Center
//   Body:         11pt, Justified, spaceAfter=63500 EMU (~5pt), lineSpacing=1.0
//   List items:   spaceBefore=63500, spaceAfter=63500, lineSpacing=1.0, no number
//   Sig block:    TAB-separated (not a table): left=applicant, right=attorney
//                 8 blank lines before sig, 4 blank lines before address
// ─────────────────────────────────────────────────────────────────────────────

const FONT       = 'Cambria';
const TITLE_PT   = 28;   // 14pt in half-pts
const BODY_PT    = 22;   // 11pt in half-pts
// Word default spacing (matches docDefaults from sample)
const SA_DEFAULT = 160;  // twips — spaceAfter for normal body paras
const SA_TIGHT   = 0;    // no space after (POA body paras use 63500 EMU = ~50 twips but Word stores 0 effectively)
const LS_DEFAULT = 259;  // auto line spacing (matches sample docDefaults)
const LS_EXACT1  = 240;  // single exact

const DESIGNATION_MAP = {
  'Proprietorship' : 'PROPRIETOR',
  'Partnership'    : 'PARTNER',
  'LLP'            : 'DESIGNATED PARTNER',
  'Private Limited': 'DIRECTOR',
  'Public Limited' : 'DIRECTOR',
  'HUF'            : 'KARTA',
  'Trust'          : 'TRUSTEE',
};
const SOLE_TITLE_MAP = {
  'Proprietorship' : 'SOLE PROPRIETOR',
  'Partnership'    : 'PARTNER',
  'LLP'            : 'DESIGNATED PARTNER',
  'Private Limited': 'DIRECTOR',
  'Public Limited' : 'DIRECTOR',
  'HUF'            : 'KARTA',
  'Trust'          : 'TRUSTEE',
};

function toWrittenDate(s) {
  if (!s) return '';
  const [y,m,d] = s.split('-').map(Number);
  const ord = ['','1st','2nd','3rd','4th','5th','6th','7th','8th','9th','10th',
    '11th','12th','13th','14th','15th','16th','17th','18th','19th','20th',
    '21st','22nd','23rd','24th','25th','26th','27th','28th','29th','30th','31st'];
  const mon = ['January','February','March','April','May','June',
    'July','August','September','October','November','December'];
  return ord[d]+' day of '+mon[m-1]+', '+y;
}
function toDisplayDate(s) {
  if (!s) return '';
  const p = s.split('-');
  return p.length === 3 ? p[2]+'/'+p[1]+'/'+p[0] : s;
}
function toDashDate(s) {
  if (!s) return '';
  const p = s.split('-');
  return p.length === 3 ? p[2]+'-'+p[1]+'-'+p[0] : s;
}

// ── DOCX run helpers ──────────────────────────────────────────────────────────
const R  = (t,sz) => new TextRun({ text:String(t??''), font:FONT, size:sz||BODY_PT, bold:false });
const Rb = (t,sz) => new TextRun({ text:String(t??''), font:FONT, size:sz||BODY_PT, bold:true  });
const TAB = ()    => new TextRun({ text:'\t',          font:FONT, size:BODY_PT });

// Body paragraph — justified, Word default spacing
const P = (runs, sa) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  spacing: { after: sa!==undefined?sa:SA_DEFAULT, line:LS_DEFAULT, lineRule:'auto' },
  children: Array.isArray(runs)?runs:[runs],
});
// Center paragraph
const PC = (runs, sa) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: sa!==undefined?sa:0, line:LS_DEFAULT, lineRule:'auto' },
  children: Array.isArray(runs)?runs:[runs],
});
// Left paragraph
const PL = (runs, sa) => new Paragraph({
  alignment: AlignmentType.LEFT,
  spacing: { after: sa!==undefined?sa:0, line:LS_DEFAULT, lineRule:'auto' },
  children: Array.isArray(runs)?runs:[runs],
});
// Blank line — exactly like sample (no spacing overrides = uses Word default)
const BLANK = () => new Paragraph({
  spacing: { after:SA_DEFAULT, line:LS_DEFAULT, lineRule:'auto' },
  children: [new TextRun({ text:'', font:FONT, size:BODY_PT })],
});
// Blank line with no space after (tight blank)
const BLANK0 = () => new Paragraph({
  spacing: { after:0, line:LS_EXACT1, lineRule:'exact' },
  children: [new TextRun({ text:'', font:FONT, size:BODY_PT })],
});
// Horizontal rule (dashes — matching sample which uses actual dash characters)
const HRule = () => new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  spacing: { after:SA_DEFAULT, line:LS_DEFAULT, lineRule:'auto' },
  children: [R('─'.repeat(86))],  // em-dashes to fill line
});

// POA body paragraph — spaceAfter=63500 EMU=50twips, spaceBefore=50twips, line=single
const PPOA = (runs) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  spacing: { after:50, before:50, line:LS_EXACT1, lineRule:'exact' },
  children: Array.isArray(runs)?runs:[runs],
});

// POA numbered list item (no hanging indent in sample — plain numbered text)
const POAListItem = (n, text) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  spacing: { after:50, before:50, line:LS_EXACT1, lineRule:'exact' },
  children: [R(n+'.  '+text)],
});

// ─────────────────────────────────────────────────────────────────────────────
// BUILD AFFIDAVIT DOCX
// Exactly matches uploaded sample:
//   - Title 14pt bold justified
//   - Dash rule line
//   - Body 11pt justified, Word default spacing
//   - 12 blank lines
//   - AUTHORISED SIGNATORY (bold, left)
//   - blank line
//   - DATE: xx <tabs> APPLICANT NAME (bold, justified)
//   - PLACE: xx <tabs> SOLE PROPRIETOR (bold, justified)
// ─────────────────────────────────────────────────────────────────────────────
function buildAffidavitDoc(d) {
  const desig    = DESIGNATION_MAP[d.businessType]  || d.businessType.toUpperCase();
  const soleTitle= SOLE_TITLE_MAP[d.businessType]   || d.businessType.toUpperCase();
  const isProp   = d.businessType === 'Proprietorship';

  // Opening sentence — matches sample exactly
  // "I CHIRAGKUMAR BABUBHAI PATEL proprietor of PHYNAX PHARMA having registered office at ADDRESS."
  const openRuns = isProp
    ? [R('I '), Rb(d.applicantName.toUpperCase()), R(' proprietor of '), Rb(d.businessName.toUpperCase()),
       R(' having registered office at '), Rb(d.registeredAddress.toUpperCase()), R('.')]
    : [R('I '), Rb(d.applicantName.toUpperCase()), R(', '), Rb(desig),
       R(' of '), Rb(d.businessName.toUpperCase()),
       R(' having registered office at '), Rb(d.registeredAddress.toUpperCase()), R('.')];

  // Trademark paragraph
  const tmPara = (d.usageType === 'used' && d.commencementDate)
    ? P([R('A trademark application is hereby made for registration of '),
         Rb('"'+d.brandName+'"'), R(' in '), Rb('Class '+d.businessClass),
         R(' and the said trademark (device) has been continuously used since '),
         Rb(toDisplayDate(d.commencementDate)), R(' in respect of the said business.')], SA_DEFAULT)
    : P([R('A trademark application is hereby made for registration of '),
         Rb('"'+d.brandName+'"'), R(' in '), Rb('Class '+d.businessClass),
         R(' and the said mark has been proposed to be used for the said '),
         Rb(d.businessType.toUpperCase()), R('.')], SA_DEFAULT);

  // DATE line: "DATE: 02/03/2026     <tabs>      CHIRAGKUMAR BABUBHAI PATEL"
  // Uses tab stops at ~9cm to push name to right side — matching sample
  const dateLine = new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after:0, line:LS_DEFAULT, lineRule:'auto' },
    tabStops: [{ type: TabStopType.LEFT, position: 5040 }],  // ~8.9cm from left
    children: [
      Rb('DATE:'), R('  '), Rb(toDisplayDate(d.affidavitDate)),
      TAB(), TAB(),
      Rb(d.applicantName.toUpperCase()),
    ],
  });
  // PLACE line: "PLACE: SURAT    <tabs>    SOLE PROPRIETOR"
  const placeLine = new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after:0, line:LS_DEFAULT, lineRule:'auto' },
    tabStops: [{ type: TabStopType.LEFT, position: 5040 }],
    children: [
      Rb('PLACE:'), R(' '), Rb(d.place.toUpperCase()),
      TAB(), TAB(),
      Rb(soleTitle),
    ],
  });

  // 12 blank lines = exactly what the sample has (paras 7-18)
  const blanks = Array.from({length:12}, () => BLANK());

  return new Document({
    styles: { default: { document: { run: { font:FONT, size:BODY_PT } } } },
    sections: [{
      properties: {
        page: {
          size:   { width:11906, height:16838 },   // A4
          margin: { top:1440, right:1440, bottom:1440, left:1440 }  // 1 inch all
        }
      },
      children: [
        P([Rb('AFFIDAVIT', TITLE_PT)], SA_DEFAULT),
        HRule(),
        P(openRuns, SA_DEFAULT),
        P([R('That I am an Indian by nationality and residing at '),
           Rb(d.residentialAddress.toUpperCase()), R('.')], SA_DEFAULT),
        P([R('I state that I am familiar and well conversant with the facts and circumstances of the present matters and competent and authorised to swear this affidavit and make the necessary statements in respect thereof.')], SA_DEFAULT),
        tmPara,
        P([R('I solemnly state that the content of this affidavit is true to the best of my knowledge and belief and that it conceals nothing and that no part is false.')], SA_DEFAULT),
        ...blanks,
        PL([Rb('AUTHORISED SIGNATORY')], SA_DEFAULT),
        BLANK0(),
        dateLine,
        placeLine,
      ]
    }]
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// BUILD POA DOCX
// Exactly matches uploaded sample:
//   - Title 14pt bold CENTER
//   - Subtitle normal center
//   - Dash rule
//   - Opening para: justified, spaceAfter=0, line=single
//   - "As my/our..." bold
//   - List items plain numbered (no hanging), spaceAfter=50, spaceBefore=50
//   - Ratification bold
//   - 8 blank lines
//   - Sig block: TAB-separated lines (not a table)
//     "Signature of Applicant(s):  <tabs>  Accepted by:"
//     "Name: XXXX           <tabs>  Name: YYYY"
//     "Designation: XXXX    <tabs>  Agent Code: XXXX"
//     "For: COMPANY"
//   - 4 blank lines
//   - To, / Registrar address
//   - Date
// ─────────────────────────────────────────────────────────────────────────────
function buildPOADoc(d) {
  const desig    = DESIGNATION_MAP[d.businessType]  || d.businessType.toUpperCase();
  const isProp   = d.businessType === 'Proprietorship';
  const pronoun  = isProp ? 'I' : 'I/We';

  // Tab stop for sig block — 4536 twips ≈ 8cm from left margin
  const SIG_TAB = 4536;

  const sigTabStop = [{ type: TabStopType.LEFT, position: SIG_TAB }];

  // Sig line helper
  const sigLine = (leftRuns, rightRuns) => new Paragraph({
    spacing: { after:0, line:LS_EXACT1, lineRule:'exact' },
    tabStops: sigTabStop,
    children: [...leftRuns, TAB(), TAB(), TAB(), TAB(), TAB(), TAB(), ...rightRuns],
  });

  const openRuns = isProp
    ? [R('I'), R(', '), Rb('Mr. '+d.applicantName), R(', Proprietor of '),
       Rb(d.businessName), R(', a Proprietorship firm, having address at '),
       Rb(d.registeredAddress), R(', do hereby appoint: '),
       Rb(d.agentName), R(', Having office at '),
       Rb(d.agentAddress), R(', and having '),
       Rb('Trade Marks Agent Code '+d.agentCode), R('.')]
    : [R('I/We'), R(', '), Rb(d.applicantName), R(', '), R(desig),
       R(' of '), Rb(d.businessName),
       R(', having address at '), Rb(d.registeredAddress),
       R(', do hereby appoint: '), Rb(d.agentName),
       R(', Having office at '), Rb(d.agentAddress),
       R(', and having '), Rb('Trade Marks Agent Code '+d.agentCode), R('.')];

  const items = [
    'Applying for registration of the following trademark(s) under the Trade Marks Act, 1999 and Rules made thereunder;',
    'Preparing, signing and submitting all applications, requests, forms, responses, and other documents;',
    'Representing me/us before the Registrar of Trade Marks or any other competent authority;',
    'Receiving and responding to all notices, objections, oppositions, and communications;',
    'Making necessary amendments or modifications to the application;',
    'Taking all necessary steps, including appearing at hearings, filing affidavits or appeals, and performing other acts, deeds and things which are necessary or incidental to the registration and protection of the said mark(s).'
  ];

  // 8 blank lines before sig block (matches sample paras 11-18)
  const blanks8 = Array.from({length:8}, () => BLANK0());
  // 4 blank lines before address block
  const blanks4 = Array.from({length:4}, () => BLANK0());

  return new Document({
    styles: { default: { document: { run: { font:FONT, size:BODY_PT } } } },
    sections: [{
      properties: {
        page: {
          size:   { width:11906, height:16838 },
          margin: { top:1440, right:1440, bottom:1440, left:1440 }
        }
      },
      children: [
        // Title block — center
        PC([Rb('GENERAL POWER OF ATTORNEY FOR TRADEMARK/SERVICE MARK', TITLE_PT)], 0),
        PC([R('(UNDER SECTION 145 OF THE TRADEMARKS ACT, 1999)')], 0),
        new Paragraph({   // rule with spaceBefore+After like sample
          alignment: AlignmentType.CENTER,
          spacing: { after:50, before:50, line:LS_EXACT1, lineRule:'exact' },
          children: [R('─'.repeat(86))],
        }),

        // Opening paragraph — spaceAfter=0
        PPOA(openRuns),

        // Scope heading
        PPOA([Rb('As my/our lawful Attorney to act on my/our behalf in respect of:')]),

        // Numbered list
        ...items.map((item, i) => POAListItem(i+1, item)),

        // Ratification — bold
        PPOA([Rb(pronoun+' hereby confirm and ratify all acts done by the above-mentioned attorney in pursuance of this authority. Executed on this '+d.poaExecutionDate+'.')]),

        // 8 blank lines
        ...blanks8,

        // Signature block — TAB separated (exactly like sample)
        sigLine([Rb('Signature of Applicant(s):')],  [Rb('Accepted by:')]),
        sigLine([R('Name: '), Rb(d.applicantName)],  [R('Name: '), Rb(d.agentName)]),
        sigLine([R('Designation: '), Rb(desig)],     [R('Agent Code: '), Rb(d.agentCode)]),
        PL([R('For: '), Rb(d.businessName)], 0),

        // 4 blank lines
        ...blanks4,

        // Address block
        PL([Rb('To,')], 0),
        PL([Rb('The Registrar of Trade Marks,')], 0),
        PL([Rb('The Office of the Trade Marks Registry at')], 0),
        PL([Rb(d.tmOffice)], 0),
        BLANK0(),
        BLANK0(),
        BLANK0(),
        PL([R('Date: '), Rb(toDashDate(d.poaDate))], 0),
      ]
    }]
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// BUILD PDF — exact layout matching DOCX output above
// ─────────────────────────────────────────────────────────────────────────────
function buildPDF(d, type) {
  return new Promise((resolve, reject) => {
    if (!PDFDocument) return reject(new Error('pdfkit not installed. Run: npm install pdfkit'));

    const doc = new PDFDocument({ size:'A4', margins:{ top:72, right:72, bottom:72, left:72 } });
    const chunks = [];
    doc.on('data', c => chunks.push(c));
    doc.on('end',  () => resolve(Buffer.concat(chunks)));
    doc.on('error',e  => reject(e));

    const desig     = DESIGNATION_MAP[d.businessType] || d.businessType.toUpperCase();
    const soleTitle = SOLE_TITLE_MAP[d.businessType]  || d.businessType.toUpperCase();
    const isProp    = d.businessType === 'Proprietorship';
    const pronoun   = isProp ? 'I' : 'I/We';

    const L  = 72, R_EDGE = doc.page.width - 72, W = R_EDGE - L;
    const FR = 'Times-Roman', FB = 'Times-Bold';
    const TS = 14, BS = 11, SS = 10;
    const LH = BS * 1.45;

    // Write mixed runs inline
    function mixed(segs, opts={}) {
      const fs = opts.size||BS, al = opts.align||'justify';
      const last = segs.length - 1;
      segs.forEach((s,i) => {
        if (!s.text) return;
        doc.font(s.bold?FB:FR).fontSize(fs);
        doc.text(s.text, { continued: i<last, align:al, lineGap:1 });
      });
      if (opts.gap!==false) doc.moveDown(0.45);
    }
    function bold(t, opts={}) {
      doc.font(FB).fontSize(opts.size||BS);
      doc.text(t, { align:opts.align||'justify', lineGap:1 });
      if (opts.gap!==false) doc.moveDown(0.45);
    }
    function normal(t, opts={}) {
      doc.font(FR).fontSize(opts.size||BS);
      doc.text(t, { align:opts.align||'justify', lineGap:1 });
      if (opts.gap!==false) doc.moveDown(0.45);
    }
    function rule() {
      doc.moveDown(0.2);
      doc.moveTo(L, doc.y).lineTo(R_EDGE, doc.y).lineWidth(0.5).stroke();
      doc.moveDown(0.5);
    }
    function listItem(n, text) {
      const nw = 20, tx = L+nw;
      const y0 = doc.y;
      doc.font(FR).fontSize(BS).text(n+'.', L, y0, { width:nw, align:'left', lineGap:1 });
      doc.text(text, tx, y0, { width:W-nw, align:'justify', lineGap:1 });
      doc.moveDown(0.3);
    }

    // ── AFFIDAVIT ────────────────────────────────────────────────────────────
    if (type === 'affidavit') {

      bold('AFFIDAVIT', { align:'justify', size:TS, gap:false });
      doc.moveDown(0.2); rule();

      if (isProp) {
        mixed([
          {text:'I '},{text:d.applicantName.toUpperCase(),bold:true},
          {text:' proprietor of '},{text:d.businessName.toUpperCase(),bold:true},
          {text:' having registered office at '},{text:d.registeredAddress.toUpperCase(),bold:true},
          {text:'.'}
        ]);
      } else {
        mixed([
          {text:'I '},{text:d.applicantName.toUpperCase(),bold:true},
          {text:', '},{text:desig,bold:true},
          {text:' of '},{text:d.businessName.toUpperCase(),bold:true},
          {text:' having registered office at '},{text:d.registeredAddress.toUpperCase(),bold:true},
          {text:'.'}
        ]);
      }

      mixed([{text:'That I am an Indian by nationality and residing at '},
             {text:d.residentialAddress.toUpperCase(),bold:true},{text:'.'}]);

      normal('I state that I am familiar and well conversant with the facts and circumstances of the present matters and competent and authorised to swear this affidavit and make the necessary statements in respect thereof.');

      if (d.usageType==='used' && d.commencementDate) {
        mixed([{text:'A trademark application is hereby made for registration of '},
               {text:'"'+d.brandName+'"',bold:true},{text:' in '},
               {text:'Class '+d.businessClass,bold:true},
               {text:' and the said trademark (device) has been continuously used since '},
               {text:toDisplayDate(d.commencementDate),bold:true},
               {text:' in respect of the said business.'}]);
      } else {
        mixed([{text:'A trademark application is hereby made for registration of '},
               {text:'"'+d.brandName+'"',bold:true},{text:' in '},
               {text:'Class '+d.businessClass,bold:true},
               {text:' and the said mark has been proposed to be used for the said '},
               {text:d.businessType.toUpperCase(),bold:true},{text:'.'}]);
      }

      normal('I solemnly state that the content of this affidavit is true to the best of my knowledge and belief and that it conceals nothing and that no part is false.');

      // Push signature to bottom — 12 blank lines equivalent
      const sigY = doc.page.height - 72 - (LH * 5);
      if (doc.y < sigY) doc.y = sigY;

      doc.font(FB).fontSize(BS).text('AUTHORISED SIGNATORY', L, doc.y, {align:'left', lineGap:1});
      doc.moveDown(0.5);

      // DATE + NAME on same line using x positioning
      const nameX = L + W * 0.52;
      const rowY  = doc.y;
      doc.font(FB).fontSize(BS).text('DATE:  '+toDisplayDate(d.affidavitDate), L, rowY, {width: W*0.45, align:'left', lineGap:1});
      doc.font(FB).fontSize(BS).text(d.applicantName.toUpperCase(), nameX, rowY, {width: W*0.48, align:'left', lineGap:1});
      doc.moveDown(0.25);

      // PLACE + SOLE PROPRIETOR on same line
      const row2Y = doc.y;
      doc.font(FB).fontSize(BS).text('PLACE:  '+d.place.toUpperCase(), L, row2Y, {width: W*0.45, align:'left', lineGap:1});
      doc.font(FB).fontSize(BS).text(soleTitle, nameX, row2Y, {width: W*0.48, align:'left', lineGap:1});

    // ── POA ─────────────────────────────────────────────────────────────────
    } else {

      bold('GENERAL POWER OF ATTORNEY FOR TRADEMARK/SERVICE MARK', { align:'center', size:TS, gap:false });
      doc.moveDown(0.2);
      normal('(UNDER SECTION 145 OF THE TRADEMARKS ACT, 1999)', { align:'center', size:SS, gap:false });
      doc.moveDown(0.2); rule();

      if (isProp) {
        mixed([{text:'I'},{text:', '},{text:'Mr. '+d.applicantName,bold:true},
               {text:', Proprietor of '},{text:d.businessName,bold:true},
               {text:', a Proprietorship firm, having address at '},{text:d.registeredAddress,bold:true},
               {text:', do hereby appoint: '},{text:d.agentName,bold:true},
               {text:', Having office at '},{text:d.agentAddress,bold:true},
               {text:', and having '},{text:'Trade Marks Agent Code '+d.agentCode,bold:true},{text:'.'}]);
      } else {
        mixed([{text:'I/We '},{text:d.applicantName,bold:true},{text:', '},{text:desig,bold:true},
               {text:' of '},{text:d.businessName,bold:true},
               {text:', having address at '},{text:d.registeredAddress,bold:true},
               {text:', do hereby appoint: '},{text:d.agentName,bold:true},
               {text:', Having office at '},{text:d.agentAddress,bold:true},
               {text:', and having '},{text:'Trade Marks Agent Code '+d.agentCode,bold:true},{text:'.'}]);
      }

      bold('As my/our lawful Attorney to act on my/our behalf in respect of:');

      const items = [
        'Applying for registration of the following trademark(s) under the Trade Marks Act, 1999 and Rules made thereunder;',
        'Preparing, signing and submitting all applications, requests, forms, responses, and other documents;',
        'Representing me/us before the Registrar of Trade Marks or any other competent authority;',
        'Receiving and responding to all notices, objections, oppositions, and communications;',
        'Making necessary amendments or modifications to the application;',
        'Taking all necessary steps, including appearing at hearings, filing affidavits or appeals, and performing other acts, deeds and things which are necessary or incidental to the registration and protection of the said mark(s).'
      ];
      items.forEach((item,i) => listItem(i+1, item));

      bold(pronoun+' hereby confirm and ratify all acts done by the above-mentioned attorney in pursuance of this authority. Executed on this '+d.poaExecutionDate+'.');

      // Large gap — push sig block down
      const addrH  = LH * 7;
      const sigH   = LH * 5;
      const sigY   = doc.page.height - 72 - addrH - sigH;
      if (doc.y < sigY) doc.y = sigY;

      // Sig block — two columns
      const col2X = L + W * 0.52;
      let y = doc.y;

      doc.font(FB).fontSize(BS).text('Signature of Applicant(s):', L,     y, {width:W*0.45, align:'left', lineGap:1});
      doc.font(FB).fontSize(BS).text('Accepted by:',               col2X, y, {width:W*0.45, align:'left', lineGap:1});
      y += LH;
      doc.font(FR).fontSize(BS).text('Name: ',            L,     y, {continued:true, lineGap:1});
      doc.font(FB).fontSize(BS).text(d.applicantName,     {continued:false, lineGap:1});
      doc.font(FR).fontSize(BS).text('Name: ',            col2X, y, {continued:true, lineGap:1});
      doc.font(FB).fontSize(BS).text(d.agentName,         {continued:false, lineGap:1});
      y += LH;
      doc.font(FR).fontSize(BS).text('Designation: ',     L,     y, {continued:true, lineGap:1});
      doc.font(FB).fontSize(BS).text(desig,               {continued:false, lineGap:1});
      doc.font(FR).fontSize(BS).text('Agent Code: ',      col2X, y, {continued:true, lineGap:1});
      doc.font(FB).fontSize(BS).text(d.agentCode,         {continued:false, lineGap:1});
      y += LH;
      doc.font(FR).fontSize(BS).text('For: ',             L,     y, {continued:true, lineGap:1});
      doc.font(FB).fontSize(BS).text(d.businessName,      {continued:false, lineGap:1});

      // Address block
      doc.y = y + LH * 2.5;
      doc.font(FB).fontSize(BS).text('To,',                                       {align:'left', lineGap:1});
      doc.font(FB).fontSize(BS).text('The Registrar of Trade Marks,',             {align:'left', lineGap:1});
      doc.font(FB).fontSize(BS).text('The Office of the Trade Marks Registry at', {align:'left', lineGap:1});
      doc.font(FB).fontSize(BS).text(d.tmOffice,                                  {align:'left', lineGap:1});
      doc.moveDown(0.4);
      doc.font(FR).fontSize(BS).text('Date: ', {continued:true, lineGap:1});
      doc.font(FB).fontSize(BS).text(toDashDate(d.poaDate), {continued:false, lineGap:1});
    }

    doc.end();
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// ROUTES
// ─────────────────────────────────────────────────────────────────────────────
app.get('/health', (_,res) => res.json({ status:'ok', app:'tmOsphere', version:'5.1.0' }));

app.post('/api/login', (req, res) => {
  const { email='', password='', remember=false } = req.body;
  const user = USERS.find(u => u.email===email.toLowerCase().trim() && u.password===password);
  if (!user) return res.status(401).json({ error:'Invalid email or password.' });
  res.json({ success:true, token:createToken(user.email, user.name, remember), name:user.name });
});

app.post('/api/logout', (_,res) => res.json({ success:true }));

app.get('/api/me', requireAuth, (req,res) =>
  res.json({ email:req.user.email, name:req.user.name }));

app.post('/api/generate', requireAuth, async (req, res) => {
  try {
    const d = req.body;
    const required = ['applicantName','businessName','brandName','registeredAddress',
      'residentialAddress','businessClass','businessType','affidavitDate','place',
      'agentName','agentCode','agentAddress','poaDate','tmOffice'];
    const missing = required.filter(f => !d[f]?.toString().trim());
    if (missing.length) return res.status(400).json({ error:'Missing: '+missing.join(', ') });

    d.poaExecutionDate = toWrittenDate(d.poaDate);

    const [affDocx, poaDocx, affPdf, poaPdf] = await Promise.all([
      Packer.toBuffer(buildAffidavitDoc(d)),
      Packer.toBuffer(buildPOADoc(d)),
      buildPDF(d, 'affidavit'),
      buildPDF(d, 'poa'),
    ]);

    const safe = (d.brandName||'TM').replace(/[^a-zA-Z0-9_-]/g,'_');
    const zip  = new JSZip();
    zip.file('Affidavit_'+safe+'.docx', affDocx);
    zip.file('POA_'      +safe+'.docx', poaDocx);
    zip.file('Affidavit_'+safe+'.pdf',  affPdf);
    zip.file('POA_'      +safe+'.pdf',  poaPdf);

    const buf = await zip.generateAsync({ type:'nodebuffer', compression:'DEFLATE' });
    res.set({
      'Content-Type': 'application/zip',
      'Content-Disposition': `attachment; filename="TM_Documents_${safe}.zip"`
    });
    res.send(buf);

  } catch(err) {
    console.error('[generate error]', err);
    res.status(500).json({ error:err.message });
  }
});

app.get('*', (req,res) =>
  res.sendFile(path.join(__dirname, 'public', 'index.html')));

app.listen(PORT, () =>
  console.log('✅  tmOsphere v5.1 running → http://localhost:'+PORT));
