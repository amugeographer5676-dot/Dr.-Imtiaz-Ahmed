const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, LevelFormat, PageNumber, PageBreak
} = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const headerBorders = {
  top: { style: BorderStyle.SINGLE, size: 1, color: "1F4E79" },
  bottom: { style: BorderStyle.SINGLE, size: 1, color: "1F4E79" },
  left: { style: BorderStyle.SINGLE, size: 1, color: "1F4E79" },
  right: { style: BorderStyle.SINGLE, size: 1, color: "1F4E79" }
};

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 160 },
    children: [new TextRun({ text, bold: true, size: 32, font: "Arial", color: "1F4E79" })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: 26, font: "Arial", color: "2E75B6" })]
  });
}

function h3(text) {
  return new Paragraph({
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, bold: true, size: 22, font: "Arial", color: "1F4E79" })]
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 80 },
    children: [new TextRun({ text, font: "Arial", size: 20, italic: opts.italic || false, color: opts.color || "000000" })]
  });
}

function bullet(text, bold = false) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, font: "Arial", size: 20, bold })]
  });
}

function divider(color = "2E75B6") {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color, space: 1 } },
    spacing: { before: 160, after: 160 },
    children: []
  });
}

function warningBox(text) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({
      children: [new TableCell({
        borders: {
          top: { style: BorderStyle.SINGLE, size: 4, color: "C00000" },
          bottom: { style: BorderStyle.SINGLE, size: 4, color: "C00000" },
          left: { style: BorderStyle.SINGLE, size: 4, color: "C00000" },
          right: { style: BorderStyle.SINGLE, size: 4, color: "C00000" },
        },
        shading: { fill: "FFF2F2", type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 180, right: 180 },
        children: [new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 18, color: "C00000" })] })]
      })]
    })]
  });
}

function infoBox(text) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({
      children: [new TableCell({
        borders: {
          top: { style: BorderStyle.SINGLE, size: 4, color: "2E75B6" },
          bottom: { style: BorderStyle.SINGLE, size: 4, color: "2E75B6" },
          left: { style: BorderStyle.SINGLE, size: 4, color: "2E75B6" },
          right: { style: BorderStyle.SINGLE, size: 4, color: "2E75B6" },
        },
        shading: { fill: "EBF3FB", type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 180, right: 180 },
        children: [new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 18, color: "1F4E79" })] })]
      })]
    })]
  });
}

function makeTable(rows, colWidths) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: rows.map((row, ri) =>
      new TableRow({
        children: row.map((cell, ci) =>
          new TableCell({
            borders: ri === 0 ? headerBorders : borders,
            width: { size: colWidths[ci], type: WidthType.DXA },
            shading: ri === 0
              ? { fill: "1F4E79", type: ShadingType.CLEAR }
              : ri % 2 === 0 ? { fill: "F5F8FC", type: ShadingType.CLEAR } : { fill: "FFFFFF", type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({
              children: [new TextRun({
                text: cell, font: "Arial", size: 18,
                bold: ri === 0,
                color: ri === 0 ? "FFFFFF" : "000000"
              })]
            })]
          })
        )
      })
    )
  });
}

function spacer(n = 1) {
  return Array.from({ length: n }, () => new Paragraph({ children: [new TextRun("")] }));
}

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "\u2022",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } }
      }]
    }]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: "1F4E79" },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [

      // ── COVER ──────────────────────────────────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 800, after: 80 },
        children: [new TextRun({ text: "US Study & Career Resources", bold: true, size: 52, font: "Arial", color: "1F4E79" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 80, after: 80 },
        children: [new TextRun({ text: "For Dr. Imtiaz Ahmed", size: 30, font: "Arial", color: "2E75B6" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 40, after: 40 },
        children: [new TextRun({ text: "International Applicant from India  |  West Coast USA Focus", size: 22, font: "Arial", italic: true, color: "555555" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 40, after: 600 },
        children: [new TextRun({ text: "May 2026", size: 20, font: "Arial", color: "999999" })]
      }),
      divider(),
      ...spacer(1),

      // ── REALITY CHECK ──────────────────────────────────────────────────────
      h1("Honest Assessment: What Works for You"),
      para("You are currently based in India, hold a PhD, and have strong research credentials in MENA geopolitics and diaspora studies. This section cuts through all resources and tells you exactly which ones are realistic for your situation."),
      ...spacer(1),

      makeTable([
        ["Resource", "Relevant?", "Why / Why Not"],
        ["UCSD GPS PhD", "YES — Top Priority", "Fully funded, open to internationals, perfect research fit"],
        ["USC Price School PhD / MPP", "YES — Backup", "Open to internationals; partial funding available; strong policy network"],
        ["EPA AAAS Fellowship", "YES — Monitor", "Requires PhD (you have it); open to international researchers; apply once in US"],
        ["RAND MNSP", "NOT YET", "Awaiting US govt approval to admit international students — cannot apply now"],
        ["UCLA Luskin Scholarships", "NO", "Undergraduate scholarships only — not applicable for a PhD holder"],
        ["CSULB CHHS", "NO", "Health & human services focus — does not match your research profile"],
        ["USC CSII", "For Reference Only", "Research center, not a degree program — useful for networking once in the US"],
      ], [2600, 1800, 4960]),
      ...spacer(2),
      new Paragraph({ children: [new PageBreak()] }),

      // ── SECTION 1: UCSD GPS ────────────────────────────────────────────────
      h1("1.  UCSD GPS — PhD in Political Science & International Affairs"),
      para("YOUR TOP PRIORITY — Fully Funded | Perfect Research Fit | International Students Welcome", { italic: true }),
      divider(),

      h2("Why This Is Your Best Option"),
      para("As an international PhD holder based in India, UCSD GPS is the single strongest path for you. It is fully funded (tuition waiver + annual stipend), actively recruits international scholars, and has faculty whose work directly overlaps with your published research on MENA conflict, Yemeni diaspora, and geopolitics."),
      ...spacer(1),

      h2("Program Details"),
      makeTable([
        ["Detail", "Information"],
        ["Program", "PhD in Political Science & International Affairs"],
        ["Duration", "4–6 years"],
        ["Funding", "Full tuition waiver + annual stipend (~$30,000–34,000/year)"],
        ["Location", "La Jolla, San Diego, CA (West Coast)"],
        ["Format", "In-person; F-1 student visa (OPT eligible after completion)"],
        ["Application Deadline", "December 1, 2026 (for Fall 2027 entry)"],
        ["GRE", "Required"],
        ["TOEFL Minimum", "90 iBT  |  IELTS Minimum: 7.0"],
        ["Application Fee", "$155 (international applicants)"],
        ["Website", "gps.ucsd.edu/academics/phd.html"],
      ], [3500, 5860]),
      ...spacer(1),

      h2("Your Competitive Strengths for This Program"),
      bullet("PhD already completed — signals research maturity and readiness"),
      bullet("9 peer-reviewed publications — exceptionally strong for applicants at this stage"),
      bullet("Fieldwork in 4 countries including conflict-affected environments"),
      bullet("Arabic language fluency — rare and highly valued in MENA-focused programs"),
      bullet("Presentations at Georgetown University and University of Cambridge"),
      bullet("Junior Research Fellowship — demonstrates recognized scholarly standing"),
      ...spacer(1),

      h2("Application Requirements"),
      makeTable([
        ["Requirement", "Details / Notes"],
        ["Personal Statement", "1,000–1,500 words; cover research goals, why GPS, named faculty, post-PhD plans"],
        ["CV / Resume", "PDF format; highlight publications, fieldwork, language skills"],
        ["3 Letters of Recommendation", "Minimum 1 academic (PhD supervisor ideal) + 1 professional"],
        ["GRE Scores", "Target: Verbal 160+, Quant 155+, AWA 4.5+. Register at ets.org"],
        ["TOEFL / IELTS", "Min TOEFL 90 iBT or IELTS 7.0 band"],
        ["Official Transcripts", "From all institutions attended"],
        ["Application Fee", "$155 USD (non-refundable)"],
      ], [3000, 6360]),
      ...spacer(1),

      h2("Key Faculty to Target"),
      para("Contact these faculty before applying. A positive response from a potential advisor significantly strengthens your application."),
      ...spacer(1),
      makeTable([
        ["Faculty", "Why They Match You", "Email"],
        ["Prof. Eli Berman\n(Professor, Economics/GPS)", "MENA conflict, economics of religion, terrorism & insurgency — directly maps to your Yemen crisis and Red Sea research. Research Director at IGCC.", "elberman@ucsd.edu"],
        ["Prof. Samuel Bazzi\n(Associate Dean)", "Migration shapes culture & community; barriers to labor mobility; diaspora political economy — mirrors your Yemeni diaspora fieldwork in India.", "sbazzi@ucsd.edu"],
        ["Prof. Barbara F. Walter\n(Professor)", "Civil wars, international security, terrorism — directly relevant to your Yemen civil war publications.", "gps.ucsd.edu/faculty"],
        ["Prof. Aila Matanock\n(Associate Professor)", "Post-conflict statebuilding, international enforcement — aligns with your Georgetown paper on sustainable peace in MENA.", "gps.ucsd.edu/faculty"],
      ], [2600, 4360, 2400]),
      ...spacer(1),

      h2("Cold Email — Prof. Eli Berman"),
      infoBox("Send: June–July 2026  |  To: elberman@ucsd.edu  |  Subject: Prospective PhD Student – MENA Conflict & Diaspora Research"),
      ...spacer(1),
      para("Dear Professor Berman,"),
      para("I am writing to introduce myself as a prospective PhD applicant for Fall 2027 at UC San Diego's School of Global Policy and Strategy. My name is Dr. Imtiaz Ahmed, and my research sits at the intersection of conflict, religion, and diaspora politics in the Middle East — an area where your work has been particularly formative for me."),
      para("Your book 'Radical, Religious and Violent' and your research on the economics of radical religious organizations directly inform how I approach my own work. My doctoral research focused on the Yemeni diaspora in India, examining how displacement, identity, and geopolitical crisis shape transnational community behavior. I have published 9 peer-reviewed articles and presented at Georgetown University and the University of Cambridge on topics including the Yemen civil war, Red Sea security, and post-conflict reconstruction in MENA."),
      para("I am particularly drawn to your affiliation with IGCC and the ESOC project, as my fieldwork experience across four countries — including conflict-affected environments — would complement the empirical and policy-focused research your lab undertakes. I would welcome the opportunity to contribute to this work as a PhD student."),
      para("I would be grateful if you could let me know whether you are accepting doctoral students for Fall 2027, and whether my background might be a good fit for your current or upcoming projects. I am happy to share my CV, publications, or research statement at your convenience."),
      para("Warm regards,\nDr. Imtiaz Ahmed\n[Your email]  |  [Your website URL]"),
      ...spacer(1),

      h2("Cold Email — Prof. Samuel Bazzi"),
      infoBox("Send: June–July 2026  |  To: sbazzi@ucsd.edu  |  Subject: Prospective PhD Student – Diaspora, Migration & Political Economy Research"),
      ...spacer(1),
      para("Dear Professor Bazzi,"),
      para("I am writing to express my interest in joining UC San Diego's GPS PhD program for Fall 2027, with the hope of working under your supervision. My name is Dr. Imtiaz Ahmed, and my research focuses on diaspora identity, migration, and political economy in the context of the Middle East — areas that closely parallel your own work on how migration shapes culture and community."),
      para("My doctoral thesis examined the Yemeni diaspora community in India, exploring how displacement, transnational identity, and crisis-driven migration intersect with political behavior and community resilience. This work involved multi-country fieldwork across India, Yemen, and the broader MENA region, and has resulted in 9 peer-reviewed publications and presentations at Georgetown University and the University of Cambridge."),
      para("While your regional expertise centers on Southeast Asia, I believe the theoretical frameworks you develop — particularly around diversity, mobility barriers, and policy responses — are directly applicable to MENA diaspora contexts, and I would be eager to contribute a comparative dimension to your ongoing research agenda."),
      para("I would greatly appreciate the chance to speak with you about your current projects and whether you are considering doctoral students for Fall 2027. I am happy to provide my CV, a sample publication, or a research statement at your request."),
      para("Warm regards,\nDr. Imtiaz Ahmed\n[Your email]  |  [Your website URL]"),
      ...spacer(1),

      h2("Preparation Timeline"),
      makeTable([
        ["Period", "Task", "Details"],
        ["Now — May/June 2026", "Contact faculty", "Email Profs. Berman and Bazzi. Introduce yourself, ask if they're taking PhD students for Fall 2027."],
        ["September 2026", "GRE Prep Begins", "Take diagnostic test. Focus: Verbal 160+, Quant 155+, AWA 4.5+. Use Magoosh or Manhattan Prep."],
        ["September 2026", "Request Recommenders", "Ask all 3 recommenders. Provide your CV, SOP outline, and specific points to address."],
        ["October 2026", "Take GRE", "Aim to sit the test by end of October. Leaves November open for a retake if needed."],
        ["October 2026", "Draft SOP", "First full draft of Statement of Purpose. Get feedback from a mentor or colleague."],
        ["November 2026", "Finalize Application", "Polish SOP. Confirm all 3 letters submitted by Nov 25. Gather transcripts and TOEFL scores."],
        ["December 1, 2026", "SUBMIT", "Submit PhD application for Fall 2027 entry. This is a hard deadline — no extensions."],
      ], [2200, 2200, 4960]),
      ...spacer(2),
      new Paragraph({ children: [new PageBreak()] }),

      // ── SECTION 2: USC PRICE ──────────────────────────────────────────────
      h1("2.  USC Sol Price School of Public Policy"),
      para("Backup Option — Apply in Parallel | International Students Eligible | Partial Funding Available", { italic: true }),
      divider(),
      para("USC Price is ranked #3 nationally (U.S. News) and offers a PhD (DPPD) and MPP that are open to international students. While funding is less certain than UCSD GPS, the school has strong fellowships and a powerful alumni network in policy and government."),
      ...spacer(1),

      h2("Relevant Degrees for You"),
      makeTable([
        ["Degree", "Best For", "Duration"],
        ["PhD / DPPD (Doctor of Policy, Planning & Development)", "If you want to continue academic research in the US long-term", "4–5 years"],
        ["Master of Public Policy (MPP)", "If you want to pivot toward US policy practice faster", "2 years"],
        ["Master of Public Administration (MPA)", "If you want government/NGO career track", "2 years"],
      ], [3200, 3760, 2400]),
      ...spacer(1),

      h2("Scholarships & Financial Aid (International Applicants)"),
      makeTable([
        ["Award", "Details"],
        ["Dean's Merit Scholarship", "Flagship award — ranges from half to full tuition; merit-based; renewable Year 2 (min 3.0 GPA). Deadline: December 15."],
        ["Professional Scholarship", "Essay-based; for students with demonstrated contributions to public, private, or nonprofit fields"],
        ["Peace Corps Fellowship", "25% tuition scholarship; stackable with Dean's Merit for up to full tuition coverage"],
        ["Federal Financial Aid", "Available to eligible students; grad PLUS loans and other federal programs"],
      ], [3000, 6360]),
      ...spacer(1),

      infoBox("Strategy: Apply to USC Price as a parallel backup alongside UCSD GPS. If admitted to both, compare funding packages before deciding. UCSD GPS full funding will likely be stronger, but USC's policy network and LA location are distinct advantages."),
      ...spacer(2),
      new Paragraph({ children: [new PageBreak()] }),

      // ── SECTION 3: EPA ────────────────────────────────────────────────────
      h1("3.  U.S. EPA Fellowships"),
      para("Post-Arrival Option — Requires PhD (You Qualify) | Federal Research Experience | Monitor Annually", { italic: true }),
      divider(),
      para("Once you are in the US on a student or research visa, EPA fellowships become relevant. You already hold a PhD which makes you eligible for the most competitive federal fellowship tracks."),
      ...spacer(1),

      warningBox("Note: Most EPA fellowship pathways require you to already be in the US or enrolled at a US institution. Apply to these after arriving, not before."),
      ...spacer(1),

      h2("Most Relevant Fellowships for Dr. Ahmed"),
      makeTable([
        ["Fellowship", "Eligibility", "When to Apply"],
        ["AAAS Science & Technology Policy Fellowship", "Doctoral degree (PhD) holders — you qualify. Fellows work in EPA offices on national security, environment, science communication.", "Applications open each fall. Apply via aaas.org"],
        ["Presidential Management Fellows (PMF)", "Recent advanced degree graduates (within 2 years of graduation). 2-year federal leadership program with rotations across agencies including EPA.", "Monitor at pmf.gov. Apply after arriving in US."],
        ["NAS/NRC Research Associateship", "PhD or equivalent for post-doctoral research at federal labs including EPA. Up to 3 years.", "Rolling applications at nas.edu/rap"],
        ["ASPPH/EPA Environmental Health Fellowship", "Early-career public health professionals within 5 years of graduation from ASPPH member school. 1-year placement.", "Applications: December–January each year via aspph.org"],
      ], [2600, 3760, 3000]),
      ...spacer(2),
      new Paragraph({ children: [new PageBreak()] }),

      // ── SECTION 4: RAND NOTE ─────────────────────────────────────────────
      h1("4.  RAND School of Public Policy — On Hold for Now"),
      divider(),
      warningBox("RAND MNSP is currently NOT accepting international students. The program is awaiting U.S. government approval to admit international applicants. You cannot apply at this time."),
      ...spacer(1),
      para("When international admissions open, RAND MNSP will be worth revisiting. Here is what to know for future reference:"),
      ...spacer(1),
      makeTable([
        ["Detail", "Information"],
        ["Degree", "Master of National Security Policy (MNSP)"],
        ["Duration", "9 months (full-time) or up to 3 years (part-time)"],
        ["Campuses", "Santa Monica, CA  |  Washington, D.C."],
        ["Tuition", "$50,000 full-time; up to $20,000 scholarship available"],
        ["Why It Fits You", "National security, foreign policy, and defense focus — directly aligned with your MENA and Yemen research"],
        ["Action", "Complete the Request for Information form at rand.edu to be notified when international admissions open"],
      ], [3000, 6360]),
      ...spacer(2),
      new Paragraph({ children: [new PageBreak()] }),

      // ── MASTER PLAN ──────────────────────────────────────────────────────
      h1("Your Master Action Plan — From India to the US"),
      divider(),
      makeTable([
        ["Timeline", "Action", "Priority"],
        ["June–July 2026", "Email Prof. Eli Berman and Prof. Samuel Bazzi at UCSD GPS. Introduce yourself and ask if they are taking PhD students for Fall 2027.", "CRITICAL"],
        ["August 2026", "Register for GRE at ets.org. Book your test slot for late October. Begin prep with Magoosh or Manhattan Prep.", "HIGH"],
        ["September 2026", "Ask 3 recommenders. Provide CV + SOP outline + specific points you want each letter to address.", "HIGH"],
        ["September 2026", "Begin drafting your Statement of Purpose. Focus: research arc, why GPS/UCSD, named faculty, post-PhD goals.", "HIGH"],
        ["October 2026", "Sit GRE. Target: Verbal 160+, Quant 155+, AWA 4.5+. Also begin USC Price application in parallel.", "HIGH"],
        ["November 2026", "Finalize SOP. Confirm all 3 recommendation letters submitted by Nov 25. Gather TOEFL + transcripts.", "HIGH"],
        ["December 1, 2026", "SUBMIT UCSD GPS PhD application. Also submit USC Price application (deadline Dec 15 for scholarships).", "DEADLINE"],
        ["Jan–Mar 2027", "Interview invitations typically sent Jan–Feb. Prepare to discuss your research agenda and fit with GPS.", "PREPARE"],
        ["April 2027", "Decision letters issued. Compare funding packages if admitted to multiple programs.", "DECIDE"],
        ["After Arrival in US", "Apply for EPA AAAS Fellowship or PMF once on F-1 visa and enrolled. Monitor RAND for international admissions opening.", "FUTURE"],
      ], [2000, 4960, 2400]),
      ...spacer(2),
      divider(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 80 },
        children: [new TextRun({ text: "Prepared for Dr. Imtiaz Ahmed  |  International Applicant from India  |  May 2026", font: "Arial", size: 18, italic: true, color: "888888" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 40, after: 40 },
        children: [new TextRun({ text: "Sources: UCSD GPS, USC Price, EPA, RAND official websites + personal research compilation", font: "Arial", size: 16, italic: true, color: "AAAAAA" })]
      }),

    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("Dr_Ahmed_Resources_International.docx", buf);
  console.log("Done");
});
