const express = require('express');
const path    = require('path');
const JSZip   = require('jszip');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType
} = require('docx');

const app  = express();
const PORT = process.env.PORT || 3001;

// ── Middleware ─────────────────────────────────────────────────────────────
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ── Document helpers ───────────────────────────────────────────────────────
const bold   = (t, sz = 22) => new TextRun({ text: String(t), bold: true, size: sz });
const normal = (t, sz = 22) => new TextRun({ text: String(t), size: sz });
const empty  = ()            => new Paragraph({ children: [] });
const LINE   = '─'.repeat(95);

// ── Affidavit builder ──────────────────────────────────────────────────────
function buildAffidavit(d) {
  return new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: 'AFFIDAVIT', bold: true, size: 32 })]
        }),
        new Paragraph({ children: [normal(LINE, 18)] }),
        empty(),
        new Paragraph({
          children: [
            normal('I '), bold(d.applicantName),
            normal(' proprietor of "'), bold(d.brandName),
            normal('" having registered office AT- '), bold(d.registeredAddress)
          ]
        }),
        empty(),
        new Paragraph({ children: [normal('Do hereby solely affirm as follow:')] }),
        empty(),
        new Paragraph({
          children: [
            normal('That I am an Indian by nationality and residing at '),
            bold(d.residentialAddress), normal('.')
          ]
        }),
        empty(),
        new Paragraph({
          children: [normal(
            'I state that I am familiar and well conversant with the facts and circumstances ' +
            'of the present matters and competent and authorised to swear this affidavit and ' +
            'make the necessary statements in respect thereof.'
          )]
        }),
        empty(),
        new Paragraph({
          children: [
            normal('A trademark application is hereby made for registration of the accompanying trademark '),
            bold('"' + d.brandName + '"'), normal(' in '),
            bold('CLASS ' + d.businessClass),
            normal(' and the said mark has been proposed to be used for the said ' + d.businessType + '.')
          ]
        }),
        empty(),
        new Paragraph({
          children: [normal(
            'I solemnly state that the content of this affidavit is true to the best of my ' +
            'knowledge and belief and that it conceals nothing and that no part is false.'
          )]
        }),
        empty(), empty(),
        new Paragraph({ children: [bold(d.applicantName)] }),
        empty(),
        new Paragraph({ children: [normal('DATE: ' + d.affidavitDate)] }),
        empty(),
        new Paragraph({ children: [normal('PLACE: ' + d.place)] }),
      ]
    }]
  });
}

// ── Power of Attorney builder ──────────────────────────────────────────────
function buildPOA(d) {
  const items = [
    'Applying for registration of the following trademark(s) under the Trade Marks Act, 1999 and Rules made thereunder;',
    'Preparing, signing and submitting all applications, requests, forms, responses, and other documents;',
    'Representing me/us before the Registrar of Trade Marks or any other competent authority;',
    'Receiving and responding to all notices, objections, oppositions, and communications;',
    'Making necessary amendments or modifications to the application;',
    'Taking all necessary steps, including appearing at hearings, filing affidavits or appeals, and performing other acts, deeds and things which are necessary or incidental to the registration and protection of the said mark(s).'
  ];

  return new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: 'GENERAL POWER OF ATTORNEY FOR TRADEMARK/SERVICE MARK', bold: true, size: 26 })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [normal('(UNDER SECTION 145 OF THE TRADEMARKS ACT, 1999)', 20)]
        }),
        new Paragraph({ children: [normal(LINE, 18)] }),
        empty(),
        new Paragraph({
          children: [
            normal('I '), bold(d.applicantName),
            normal(' proprietor of "'), bold(d.brandName),
            normal('" having registered office AT - '), bold(d.registeredAddress),
            normal(', do hereby appoint: '), bold(d.agentName),
            normal(', Having office at '), bold(d.agentAddress),
            normal(', and having '), bold('Trade Marks Agent Code ' + d.agentCode), normal('.')
          ]
        }),
        empty(),
        new Paragraph({ children: [bold('As my/our lawful Attorney to act on my/our behalf in respect of:')] }),
        empty(),
        ...items.map((item, i) =>
          new Paragraph({ children: [normal((i + 1) + '.  ' + item)] })
        ),
        empty(),
        new Paragraph({
          children: [
            bold('I/We hereby confirm and ratify all acts done by the above-mentioned attorney ' +
                 'in pursuance of this authority executed on this '),
            bold(d.poaExecutionDate), bold('.')
          ]
        }),
        empty(), empty(),
        new Paragraph({
          children: [bold('Signature of Applicant(s):                                              Accepted by:')]
        }),
        new Paragraph({
          children: [
            normal('Name: ' + d.applicantName + '                                                    Name: '),
            bold(d.agentName)
          ]
        }),
        new Paragraph({
          children: [
            bold('Designation: SOLE PROPRIETOR'),
            normal('                                               Agent Code: '),
            bold(d.agentCode)
          ]
        }),
        empty(),
        new Paragraph({ children: [normal('For: '), bold(d.brandName)] }),
        empty(),
        new Paragraph({ children: [bold('To,')] }),
        new Paragraph({ children: [bold('The Registrar of Trade Marks,')] }),
        new Paragraph({ children: [bold('The Office of the Trade Marks Registry at')] }),
        new Paragraph({ children: [bold(d.tmOffice)] }),
        empty(),
        new Paragraph({ children: [normal('Date: ' + d.poaDate)] }),
      ]
    }]
  });
}

// ── API routes ─────────────────────────────────────────────────────────────
app.get('/health', (req, res) => {
  res.json({ status: 'ok', app: 'tmOsphere', version: '1.0.0' });
});

app.post('/api/generate', async (req, res) => {
  try {
    const d = req.body;

    // Basic validation
    const required = [
      'applicantName','brandName','registeredAddress','residentialAddress',
      'businessClass','businessType','affidavitDate','place',
      'agentName','agentCode','agentAddress','poaDate','poaExecutionDate','tmOffice'
    ];
    const missing = required.filter(f => !d[f] || !String(d[f]).trim());
    if (missing.length) {
      return res.status(400).json({ error: 'Missing fields: ' + missing.join(', ') });
    }

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

// ── Catch-all → serve frontend ─────────────────────────────────────────────
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ── Start ──────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`✅  tmOsphere running → http://localhost:${PORT}`);
});
