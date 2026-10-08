# vml-documento-engine

Gera documentos Word (.docx) no timbrado oficial VML a partir de um objeto de dados `D`.
Usado por agentes do WPP Open: o agente monta `D` e um artefato HTML carrega este engine, mostra a prévia e baixa o Word.

## Uso no artefato

```html
<div id="vb"></div>
<button id="baixar">Baixar Word</button>
<script src="https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js"></script>
<script src="https://cdn.jsdelivr.net/gh/tomcoliveira/vml-documento-engine@v4.0/vml-documento-docx.js"></script>
<script>
const D = {
  meta: { template: "vml-a4", tipo: "Relatório", titulo: "Título", detalhes: ["VML Brasil"], data: "08/10/2026" },
  resumo: [ { rotulo: "Prazo", valor: "Agosto" } ],
  secoes: [ { rotulo: "Contexto", tipo: "padrao", paragrafos: ["Texto."], itens: [] } ]
};
VMLDocumento.preview(D, document.getElementById("vb"));
document.getElementById("baixar").onclick = () => VMLDocumento.download(D);
</script>
```

## Tipos de seção

- `padrao`: texto corrido e lista com travessão
- `destaque`: frase principal em corpo maior
- `numerada`: itens 01, 02, 03
- `cronograma`: itens no formato "Marco: data", em duas colunas
- `caixa`: conteúdo em caixa com contorno (pendências, riscos)

## Templates

Ficam em `templates/` e são registrados no objeto `TEMPLATES` do engine.

- `vml-a4`: timbrado VML A4 do brand kit 2026, sem o endereço de Kansas City no rodapé; logotipo só na primeira página.

Para incluir um template: subir o .docx oficial em `templates/`, registrar em `TEMPLATES` com a geometria dele e publicar uma nova tag.

## Versões

Cada tag é um endereço fixo no jsDelivr. Mudança de diagramação = nova tag + troca da versão no prompt do agente.
