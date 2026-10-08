/*
 * vml-documento-docx.js — v4.0
 * Gera documentos Word (.docx) dentro do timbrado oficial VML A4
 * (VML_Letterhead_A4.docx, brand kit 2026 — logotipo no topo da 1ª página, Spark e vml.com no rodapé).
 * Genérico: serve briefing, relatório, proposta, ata ou qualquer texto estruturado.
 *
 * Dependência, carregada antes deste arquivo:
 *   https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js
 * Templates na pasta templates/, ao lado deste script (registro em TEMPLATES). Padrão: vml-a4.
 *
 * Contrato de dados (D):
 * {
 *   meta:   { template: "vml-a4", tipo: "Briefing", titulo, detalhes: ["KitKat (Nestlé)", "Equipe criativa", "v1"], data: "08/10/2026" },
 *   resumo: [ { rotulo: "Prazo final", valor: "Agosto" } ],          // até 3 itens, opcional
 *   secoes: [ { rotulo: "Contexto", tipo: "padrao", paragrafos: ["..."], itens: ["..."] } ]
 * }
 * Tipos de seção:
 *   padrao      → texto corrido e lista com travessão
 *   destaque    → texto em corpo maior (a frase principal do documento)
 *   numerada    → itens numerados 01, 02...
 *   cronograma  → itens "Marco: data" em duas colunas, data em negrito
 *   caixa       → conteúdo dentro de caixa com contorno preto (pendências, riscos)
 * Sem "tipo", o engine deduz pelo rótulo (objetivo → destaque; entregáveis/próximos passos → numerada;
 * crono/prazo/cronograma → cronograma; pontos a confirmar/pendências/riscos → caixa).
 * Compatível com a v3: meta.cliente, meta.publico e meta.versao viram detalhes quando meta.detalhes não vier.
 * Linhas que começam com "Recomendação:" recebem o rótulo em negrito.
 *
 * Navegador: VMLDocumento.preview(D, elemento); VMLDocumento.download(D)
 */
(function (root) {
  /*
   * Registro de templates. Cada template é um .docx oficial em templates/<id>.docx com a geometria dele.
   * Para incluir um novo: subir o .docx em templates/ e acrescentar uma entrada aqui com as medidas do timbrado.
   */
  var TEMPLATES = {
    "vml-a4": {
      arquivo: "templates/vml-a4.docx",
      nome: "Timbrado VML A4 (brand kit 2026)",
      tokens: {}
    }
  };
  var PADRAO = "vml-a4";

  function factory(JSZip, loadTemplate) {
    var BASE = {
      book: "Snowflake Sans Book",
      bold: "Snowflake Sans",
      color: "020303",
      gray: "53544A",
      rule: "C2C4B8",
      black: "191919",
      colL: 1921, colR: 6700, labelW: 2611, labelInd: -690,
      sz: { body: 18, label: 16, small: 14, title: 44, meta: 18, lead: 24, value: 21 },
      line: 367,
      top: 1890
    };
    var T = BASE;
    function useTemplate(id) {
      var t = TEMPLATES[id] || TEMPLATES[PADRAO];
      T = Object.assign({}, BASE, t.tokens);
      return t;
    }

    function esc(s) {
      return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    }

    function rPr(o) {
      o = o || {};
      var f = o.b ? T.bold : T.book, sz = o.sz || T.sz.body;
      return "<w:rPr><w:rFonts w:ascii=\"" + f + "\" w:hAnsi=\"" + f + "\" w:cs=\"" + f + "\" w:eastAsia=\"" + f + "\"/>" +
        (o.b ? "<w:b/><w:bCs/>" : "") + (o.caps ? "<w:caps/>" : "") +
        "<w:color w:val=\"" + (o.color || T.color) + "\"/><w:sz w:val=\"" + sz + "\"/><w:szCs w:val=\"" + sz + "\"/></w:rPr>";
    }
    function r(text, o) { return "<w:r>" + rPr(o) + "<w:t xml:space=\"preserve\">" + esc(text) + "</w:t></w:r>"; }
    function tab(o) { return "<w:r>" + rPr(o) + "<w:tab/></w:r>"; }

    function p(runs, o) {
      o = o || {};
      var left = o.left != null ? o.left : T.colL, right = o.right != null ? o.right : 399;
      var ind = "<w:ind w:left=\"" + left + "\" w:right=\"" + right + "\"" + (o.hanging ? " w:hanging=\"" + o.hanging + "\"" : "") + "/>";
      return "<w:p><w:pPr><w:pStyle w:val=\"BodyText\"/>" + (o.keepNext ? "<w:keepNext/>" : "") +
        (o.tabs ? "<w:tabs>" + o.tabs + "</w:tabs>" : "") +
        "<w:spacing w:before=\"" + (o.before || 0) + "\" w:after=\"" + (o.after || 0) + "\" w:line=\"" + (o.line || T.line) + "\" w:lineRule=\"auto\"/>" +
        ind + "</w:pPr>" + (runs || "") + "</w:p>";
    }

    function border(side, sz, color) { return "<w:" + side + " w:val=\"single\" w:sz=\"" + sz + "\" w:space=\"0\" w:color=\"" + color + "\"/>"; }
    function noBorder(side) { return "<w:" + side + " w:val=\"nil\"/>"; }

    function cell(width, content, o) {
      o = o || {};
      var b = o.box
        ? border("top", 8, T.black) + border("left", 8, T.black) + border("bottom", 8, T.black) + border("right", 8, T.black)
        : (o.topRule ? border("top", o.topRule.sz, o.topRule.color) : noBorder("top")) + noBorder("left") + noBorder("bottom") + noBorder("right");
      var m = o.mar || { top: 0, left: 0, bottom: 0, right: 0 };
      return "<w:tc><w:tcPr><w:tcW w:w=\"" + width + "\" w:type=\"dxa\"/><w:tcBorders>" + b + "</w:tcBorders>" +
        "<w:tcMar><w:top w:w=\"" + m.top + "\" w:type=\"dxa\"/><w:left w:w=\"" + m.left + "\" w:type=\"dxa\"/><w:bottom w:w=\"" + m.bottom + "\" w:type=\"dxa\"/><w:right w:w=\"" + m.right + "\" w:type=\"dxa\"/></w:tcMar>" +
        "</w:tcPr>" + (content || p("", { left: 0, right: 0 })) + "</w:tc>";
    }

    function table(widths, rowsXml, o) {
      o = o || {};
      var total = widths.reduce(function (a, b) { return a + b; }, 0);
      return "<w:tbl><w:tblPr><w:tblW w:w=\"" + total + "\" w:type=\"dxa\"/><w:tblInd w:w=\"" + (o.ind || 0) + "\" w:type=\"dxa\"/>" +
        "<w:tblBorders>" + noBorder("top") + noBorder("left") + noBorder("bottom") + noBorder("right") + noBorder("insideH") + noBorder("insideV") + "</w:tblBorders>" +
        "<w:tblLayout w:type=\"fixed\"/><w:tblCellMar><w:left w:w=\"0\" w:type=\"dxa\"/><w:right w:w=\"0\" w:type=\"dxa\"/></w:tblCellMar><w:tblLook w:val=\"0000\"/></w:tblPr>" +
        "<w:tblGrid>" + widths.map(function (w) { return "<w:gridCol w:w=\"" + w + "\"/>"; }).join("") + "</w:tblGrid>" + rowsXml + "</w:tbl>";
    }

    function spacer(after) { return p("", { left: 0, right: 0, after: after || 0, line: 240 }); }

    function inline(text, o) {
      var m = /^Recomenda[çc][aã]o:\s*/i.exec(text);
      return m ? r("Recomendação: ", { b: true, sz: (o && o.sz) }) + r(text.slice(m[0].length), o) : r(text, o);
    }

    function norm(s) {
      return String(s || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z ]/g, " ").trim();
    }
    var TIPOS = { padrao: 1, destaque: 1, numerada: 1, cronograma: 1, caixa: 1 };
    function kind(s) {
      if (s && TIPOS[s.tipo]) return s.tipo;
      var n = norm(s && s.rotulo);
      if (/^objetivo/.test(n)) return "destaque";
      if (/^(entregaveis|proximos passos)/.test(n)) return "numerada";
      if (/^(crono|prazo|cronograma)/.test(n)) return "cronograma";
      if (/^(pontos a confirmar|pendencias|riscos)/.test(n)) return "caixa";
      return "padrao";
    }
    function detalhes(meta) {
      return (meta.detalhes && meta.detalhes.length ? meta.detalhes : [meta.cliente, meta.publico, meta.versao]).filter(Boolean);
    }

    function sectionContent(s) {
      var k = kind(s), x = "", C = { left: 0, right: 0 }, crono = [];
      (s.paragrafos || []).forEach(function (t) {
        x += k === "destaque"
          ? p(inline(t, { sz: T.sz.lead }), { left: 0, right: 0, after: 100, line: 320 })
          : p(inline(t), { left: 0, right: 0, after: 100 });
      });
      (s.itens || []).forEach(function (t, i) {
        if (k === "numerada") {
          x += p(r(String(i + 1).padStart(2, "0"), { b: true }) + tab() + inline(t),
            { left: 454, right: 0, hanging: 454, after: 80, tabs: "<w:tab w:val=\"left\" w:pos=\"454\"/>" });
          return;
        }
        var m = k === "cronograma" ? /^(.+?):\s*(.+)$/.exec(t) : null;
        if (m) { crono.push(m); return; }
        x += p(r("–") + tab() + inline(t), { left: 283, right: 0, hanging: 283, after: 60, tabs: "<w:tab w:val=\"left\" w:pos=\"283\"/>" });
      });
      if (crono.length) {
        var cw = [2000, T.colR - 2000];
        x += table(cw, crono.map(function (m) {
          var M = { top: 0, left: 0, bottom: 80, right: 160 };
          return "<w:tr><w:trPr><w:cantSplit/></w:trPr>" +
            cell(cw[0], p(r(m[2], { b: true }), { left: 0, right: 0, line: 276 }), { mar: M }) +
            cell(cw[1], p(r(m[1]), { left: 0, right: 0, line: 276 }), { mar: { top: 0, left: 0, bottom: 80, right: 0 } }) + "</w:tr>";
        }).join("")) + p("", { left: 0, right: 0, line: 120 });
      }
      return x || p("", C);
    }

    function bodyXml(D) {
      var meta = D.meta || {}, x = "";
      x += p(r(meta.tipo || "Documento", { b: true, caps: true, sz: T.sz.label }) + (meta.data ? tab() + r(meta.data) : ""),
        { tabs: "<w:tab w:val=\"right\" w:pos=\"" + (T.colL + T.colR) + "\"/>", before: T.top, after: 200, line: 276 });
      x += p(r(meta.titulo || meta.tipo || "Documento", { sz: T.sz.title }), { after: 120, line: 252 });
      var metaLine = detalhes(meta).join("  |  ");
      if (metaLine) x += p(r(metaLine, { color: T.gray }), { after: 360, line: 276 });

      var res = (D.resumo || []).slice(0, 3);
      if (res.length) {
        var w = Math.floor(T.colR / res.length), widths = res.map(function () { return w; });
        widths[widths.length - 1] = T.colR - w * (res.length - 1);
        var row = "<w:tr><w:trPr><w:cantSplit/></w:trPr>" + res.map(function (it, i) {
          var c = p(r(it.rotulo, { b: true, caps: true, sz: T.sz.small, color: T.gray }), { left: 0, right: 200, after: 60, line: 240 }) +
                  p(r(it.valor, { sz: T.sz.value }), { left: 0, right: 200, line: 276 });
          return cell(widths[i], c, { topRule: { sz: 12, color: T.black }, mar: { top: 140, left: 0, bottom: 0, right: 0 } });
        }).join("") + "</w:tr>";
        x += table(widths, row, { ind: T.colL }) + spacer(360);
      }

      var rows = (D.secoes || []).map(function (s) {
        var k = kind(s);
        var label = p(r(String(s.rotulo || "").replace(/:\s*$/, ""), { b: true, caps: true, sz: T.sz.label }), { left: 0, right: 280, line: 276 });
        var rule = { sz: 4, color: T.rule };
        var left = cell(T.labelW, label, { topRule: rule, mar: { top: 160, left: 0, bottom: 240, right: 0 } });
        var right = k === "caixa"
          ? cell(T.colR, sectionContent(s), { box: true, mar: { top: 160, left: 200, bottom: 80, right: 200 } })
          : cell(T.colR, sectionContent(s), { topRule: rule, mar: { top: 160, left: 0, bottom: 240, right: 0 } });
        return "<w:tr>" + left + right + "</w:tr>";
      }).join("");
      if (rows) x += table([T.labelW, T.colR], rows, { ind: T.labelInd }) + spacer();
      return x;
    }

    function slug(s) {
      return String(s || "briefing").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "")
        .replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 60);
    }
    function stamp(dateStr) {
      var m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(dateStr || "");
      if (m) return m[3] + m[2] + m[1];
      var d = new Date();
      return d.getFullYear() + String(d.getMonth() + 1).padStart(2, "0") + String(d.getDate()).padStart(2, "0");
    }
    function fileName(D) {
      var meta = D.meta || {};
      return slug(meta.tipo || "documento") + "-" + slug(meta.titulo) + "-" + stamp(meta.data) + ".docx";
    }

    function build(D, type) {
      var tpl = useTemplate((D.meta || {}).template);
      return Promise.resolve(loadTemplate(tpl.arquivo)).then(function (buf) { return JSZip.loadAsync(buf); }).then(function (zip) {
        return zip.file("word/document.xml").async("string").then(function (doc) {
          var i = doc.indexOf("<w:body>") + "<w:body>".length, j = doc.indexOf("<w:sectPr");
          zip.file("word/document.xml", doc.slice(0, i) + bodyXml(D) + doc.slice(j));
          return zip.generateAsync({ type: type || "blob", compression: "DEFLATE", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
        });
      });
    }

    function h(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
    function ht(t) {
      var m = /^Recomenda[çc][aã]o:\s*/i.exec(t);
      return m ? "<b>Recomendação:</b> " + h(t.slice(m[0].length)) : h(t);
    }

    function preview(D, el) {
      var meta = D.meta || {};
      var css = "<style>.vb{font-family:'Snowflake Sans Book','Snowflake Sans',Arial,sans-serif;color:#020303;background:#fff;max-width:820px;padding:56px;box-sizing:border-box;font-size:13px;line-height:1.53}" +
        ".vb-lab{font-family:'Snowflake Sans',Arial,sans-serif;font-weight:700;text-transform:uppercase;font-size:11px;line-height:1.3}" +
        ".vb-in{margin-left:168px}.vb-top{display:flex;justify-content:space-between}.vb h1{font-weight:400;font-size:32px;line-height:1.1;margin:14px 0 8px}" +
        ".vb-meta{color:#53544A;margin-bottom:28px}.vb-res{display:flex;gap:0;margin-bottom:28px}.vb-res div{flex:1;border-top:2px solid #191919;padding-top:10px;padding-right:14px}" +
        ".vb-res .vb-lab{color:#53544A;font-size:10px}.vb-res span{display:block;font-size:15px;margin-top:4px}" +
        ".vb-row{display:flex;border-top:1px solid #C2C4B8;padding:12px 0 18px}.vb-row>.vb-lab{width:168px;flex:none;padding-right:16px;box-sizing:border-box}.vb-row>div:last-child{flex:1}" +
        ".vb p{margin:0 0 8px}.vb-obj p{font-size:17px;line-height:1.4}.vb ul,.vb ol{margin:0;padding-left:20px}.vb ul{list-style:'– '}" +
        ".vb ol{list-style:decimal-leading-zero}.vb ol li::marker{font-weight:700}.vb-box{border:1px solid #191919;padding:12px 14px}" +
        ".vb-cr{display:flex}.vb-cr b{width:110px;flex:none}</style>";
      var x = css + '<div class="vb"><div class="vb-in"><div class="vb-top"><span class="vb-lab">' + h(meta.tipo || "Documento") + '</span><span>' + h(meta.data || "") + "</span></div>" +
        "<h1>" + h(meta.titulo || meta.tipo || "Documento") + '</h1><div class="vb-meta">' + h(detalhes(meta).join("  |  ")) + "</div>";
      var res = (D.resumo || []).slice(0, 3);
      if (res.length) x += '<div class="vb-res">' + res.map(function (it) { return '<div><div class="vb-lab">' + h(it.rotulo) + "</div><span>" + h(it.valor) + "</span></div>"; }).join("") + "</div>";
      x += "</div>";
      (D.secoes || []).forEach(function (s) {
        var k = kind(s), c = "";
        (s.paragrafos || []).forEach(function (t) { c += "<p>" + ht(t) + "</p>"; });
        var it = s.itens || [];
        if (it.length) {
          if (k === "numerada") c += "<ol>" + it.map(function (t) { return "<li>" + ht(t) + "</li>"; }).join("") + "</ol>";
          else if (k === "cronograma") c += it.map(function (t) { var m = /^(.+?):\s*(.+)$/.exec(t); return m ? '<div class="vb-cr"><b>' + h(m[2]) + "</b><span>" + h(m[1]) + "</span></div>" : "<p>" + ht(t) + "</p>"; }).join("");
          else c += "<ul>" + it.map(function (t) { return "<li>" + ht(t) + "</li>"; }).join("") + "</ul>";
        }
        x += '<div class="vb-row"><div class="vb-lab">' + h(String(s.rotulo || "").replace(/:\s*$/, "")) + "</div><div" +
          (k === "destaque" ? ' class="vb-obj"' : k === "caixa" ? ' class="vb-box"' : "") + ">" + c + "</div></div>";
      });
      el.innerHTML = x + "</div>";
    }

    function download(D) {
      return build(D, "blob").then(function (blob) {
        var a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = fileName(D);
        document.body.appendChild(a);
        a.click();
        setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
      });
    }

    return { version: "4.0", templates: Object.keys(TEMPLATES), tokens: BASE, build: build, bodyXml: bodyXml, fileName: fileName, preview: preview, download: download };
  }

  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory;
  } else {
    var src = (document.currentScript && document.currentScript.src) || "";
    var base = src.slice(0, src.lastIndexOf("/") + 1);
    root.VMLDocumento = root.VMLBriefing = factory(root.JSZip, function (arquivo) {
      var url = base + arquivo;
      return fetch(url).then(function (res) {
        if (!res.ok) throw new Error("Template do documento não encontrado: " + url);
        return res.arrayBuffer();
      });
    });
  }
})(typeof window !== "undefined" ? window : this);
