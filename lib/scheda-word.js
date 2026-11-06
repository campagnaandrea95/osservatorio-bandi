/* Osservatorio bandi – esportazione della scheda in Word (.docx), generata nel browser del visitatore.
   Usa la libreria docx (MIT), caricata solo al primo clic su "Scarica Word". */
(function (root) {
  "use strict";

  var DOCX_SRC = "lib/docx-8.5.0.umd.js";
  var loading = null;

  function loadDocx() {
    if (root.docx) return Promise.resolve(root.docx);
    if (loading) return loading;
    loading = new Promise(function (resolve, reject) {
      var s = document.createElement("script");
      s.src = DOCX_SRC;
      s.async = true;
      s.onload = function () { root.docx ? resolve(root.docx) : reject(new Error("docx non disponibile")); };
      s.onerror = function () { loading = null; reject(new Error("caricamento non riuscito")); };
      document.head.appendChild(s);
    });
    return loading;
  }

  var ACCENT = "1F5C66";
  var FONT = "Calibri";

  function itDate(isoStr) {
    var d = String(isoStr || "").slice(0, 10);
    return /^\d{4}-\d{2}-\d{2}$/.test(d) ? d.slice(8, 10) + "/" + d.slice(5, 7) + "/" + d.slice(0, 4) : "n.d.";
  }

  function isUrl(u) { return /^https?:\/\/\S+$/i.test(u || ""); }

  /* Trasforma il testo di una voce (stesse regole della pagina web) in blocchi Word */
  function blocks(D, text, opts) {
    opts = opts || {};
    var out = [];
    var lines = String(text || "").replace(/\r/g, "").split("\n");
    var i = 0;
    while (i < lines.length) {
      var l = lines[i];
      if (!l.trim()) { i++; continue; }

      if (opts.links) {
        l.trim().split(/\s+/).forEach(function (u) {
          if (isUrl(u)) {
            out.push(new D.Paragraph({
              spacing: { after: 60 },
              children: [new D.ExternalHyperlink({ link: u, children: [new D.TextRun({ text: u, style: "Hyperlink" })] })]
            }));
          } else {
            out.push(new D.Paragraph({ spacing: { after: 60 }, children: [new D.TextRun(u)] }));
          }
        });
        i++; continue;
      }

      if (/^\s*[-•]\s+/.test(l)) {
        var lvl = /^\s{2,}[-•]\s+/.test(l) ? 1 : 0;
        out.push(new D.Paragraph({
          numbering: { reference: "elenco", level: lvl },
          spacing: { after: 40 },
          children: [new D.TextRun(l.replace(/^\s*[-•]\s+/, ""))]
        }));
        i++; continue;
      }

      if (l.indexOf(" | ") !== -1) {
        var righe = [];
        while (i < lines.length && lines[i].indexOf(" | ") !== -1) {
          righe.push(lines[i].split(" | ").map(function (c) { return c.trim(); }));
          i++;
        }
        var ncol = righe.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
        var larg = Math.floor(9000 / ncol);
        out.push(new D.Table({
          width: { size: 9000, type: D.WidthType.DXA },
          columnWidths: Array(ncol).fill(larg),
          rows: righe.map(function (r, ri) {
            var celle = [];
            for (var c = 0; c < ncol; c++) {
              celle.push(new D.TableCell({
                width: { size: larg, type: D.WidthType.DXA },
                shading: ri === 0 ? { type: D.ShadingType.CLEAR, color: "auto", fill: "E6EEF0" } : undefined,
                margins: { top: 60, bottom: 60, left: 100, right: 100 },
                children: [new D.Paragraph({ children: [new D.TextRun({ text: r[c] || "", bold: ri === 0, size: 18 })] })]
              }));
            }
            return new D.TableRow({ tableHeader: ri === 0, children: celle });
          })
        }));
        out.push(new D.Paragraph({ spacing: { after: 60 }, children: [] }));
        continue;
      }

      if (/^N\.B\./.test(l.trim())) {
        out.push(new D.Paragraph({
          spacing: { before: 60, after: 80 },
          shading: { type: D.ShadingType.CLEAR, color: "auto", fill: "F2F4F5" },
          border: { left: { style: D.BorderStyle.SINGLE, size: 12, color: ACCENT, space: 6 } },
          children: [new D.TextRun({ text: l.trim(), italics: true, size: 20 })]
        }));
        i++; continue;
      }

      out.push(new D.Paragraph({ spacing: { after: 80 }, children: [new D.TextRun(l)] }));
      i++;
    }
    if (!out.length) out.push(new D.Paragraph({ children: [new D.TextRun("n.d.")] }));
    return out;
  }

  function sezione(D, voce, valore) {
    return [new D.Paragraph({ heading: D.HeadingLevel.HEADING_2, children: [new D.TextRun(voce || "")] })]
      .concat(blocks(D, valore, { links: voce === "Link" }));
  }

  function buildDocument(D, sc, r) {
    sc = sc || {}; r = r || {};
    var titolo = sc.titolo || r.titolo || "Scheda";
    var n = sc.n != null ? sc.n : r.n;
    var corpo = [];

    corpo.push(new D.Paragraph({
      spacing: { after: 60 },
      children: [new D.TextRun({ text: "SCHEDA DI APPROFONDIMENTO", size: 18, color: "6B7A80", characterSpacing: 20 })]
    }));
    corpo.push(new D.Paragraph({ heading: D.HeadingLevel.TITLE, children: [new D.TextRun(titolo)] }));

    var info = [];
    if (n != null && n !== "") info.push("N. " + n);
    if (r.ente) info.push(r.ente);
    if (r.territorio) info.push(r.territorio);
    info.push("scheda del " + itDate(sc.generata_il));
    corpo.push(new D.Paragraph({
      spacing: { after: 200 },
      border: { bottom: { style: D.BorderStyle.SINGLE, size: 6, color: "C9D3D6", space: 6 } },
      children: [new D.TextRun({ text: info.join("  ·  "), size: 18, color: "6B7A80" })]
    }));

    (sc.voci || []).forEach(function (v) { corpo = corpo.concat(sezione(D, v.voce, v.valore)); });

    if ((sc.appendici || []).length) {
      corpo.push(new D.Paragraph({ heading: D.HeadingLevel.HEADING_1, pageBreakBefore: false, children: [new D.TextRun("Appendici")] }));
      sc.appendici.forEach(function (v) { corpo = corpo.concat(sezione(D, v.voce, v.valore)); });
    }

    var fonti = (sc.fonti || []).filter(isUrl);
    if (fonti.length) {
      corpo.push(new D.Paragraph({ heading: D.HeadingLevel.HEADING_2, children: [new D.TextRun("Fonti consultate")] }));
      fonti.forEach(function (u) {
        corpo.push(new D.Paragraph({
          spacing: { after: 40 },
          children: [new D.ExternalHyperlink({ link: u, children: [new D.TextRun({ text: u, style: "Hyperlink", size: 18 })] })]
        }));
      });
    }

    corpo.push(new D.Paragraph({
      spacing: { before: 300 },
      border: { top: { style: D.BorderStyle.SINGLE, size: 6, color: "C9D3D6", space: 6 } },
      children: [new D.TextRun({
        text: "Scheda redatta alla data indicata sulla base delle fonti elencate. Prima di presentare domanda verifica sull'avviso ufficiale eventuali rettifiche e proroghe.",
        size: 16, color: "6B7A80", italics: true
      })]
    }));

    return new D.Document({
      creator: "Osservatorio bandi",
      title: titolo,
      description: "Scheda di approfondimento – Osservatorio bandi",
      styles: {
        default: { document: { run: { font: FONT, size: 21 } } },
        paragraphStyles: [
          { id: "Title", name: "Title", basedOn: "Normal", next: "Normal", quickFormat: true,
            run: { size: 32, bold: true, color: "16262B", font: FONT }, paragraph: { spacing: { after: 80 } } },
          { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
            run: { size: 26, bold: true, color: ACCENT, font: FONT }, paragraph: { spacing: { before: 320, after: 120 }, outlineLevel: 0 } },
          { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
            run: { size: 23, bold: true, color: ACCENT, font: FONT }, paragraph: { spacing: { before: 220, after: 80 }, keepNext: true, outlineLevel: 1 } }
        ]
      },
      numbering: {
        config: [{
          reference: "elenco",
          levels: [
            { level: 0, format: D.LevelFormat.BULLET, text: "•", alignment: D.AlignmentType.LEFT,
              style: { paragraph: { indent: { left: 360, hanging: 240 } } } },
            { level: 1, format: D.LevelFormat.BULLET, text: "–", alignment: D.AlignmentType.LEFT,
              style: { paragraph: { indent: { left: 720, hanging: 240 } } } }
          ]
        }]
      },
      sections: [{
        properties: {
          page: { size: { width: 11906, height: 16838 }, margin: { top: 1300, right: 1300, bottom: 1300, left: 1300 } }
        },
        headers: {
          default: new D.Header({ children: [new D.Paragraph({
            alignment: D.AlignmentType.RIGHT,
            children: [new D.TextRun({ text: "Osservatorio bandi", size: 16, color: "6B7A80" })]
          })] })
        },
        footers: {
          default: new D.Footer({ children: [new D.Paragraph({
            alignment: D.AlignmentType.RIGHT,
            children: [new D.TextRun({ children: ["Pagina ", D.PageNumber.CURRENT, " di ", D.PageNumber.TOTAL_PAGES], size: 16, color: "6B7A80" })]
          })] })
        },
        children: corpo
      }]
    });
  }

  function fileName(sc, r) {
    var n = (sc && sc.n != null) ? sc.n : (r && r.n);
    var base = String((sc && sc.titolo) || (r && r.titolo) || "scheda")
      .normalize("NFD").replace(/[̀-ͯ]/g, "")
      .replace(/[^A-Za-z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 60);
    return "Scheda" + (n != null && n !== "" ? "-N" + n : "") + "-" + (base || "bando") + ".docx";
  }

  function download(sc, r) {
    return loadDocx().then(function (D) {
      return D.Packer.toBlob(buildDocument(D, sc, r));
    }).then(function (blob) {
      var a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = fileName(sc, r);
      document.body.appendChild(a);
      a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
    });
  }

  root.SchedaWord = { download: download, buildDocument: buildDocument, fileName: fileName };
})(typeof window !== "undefined" ? window : globalThis);
