import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  ShadingType,
  UnderlineType,
} from "docx";
import fs from "fs";

// ─────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────
const bold = (text, size = 24) =>
  new TextRun({ text, bold: true, size, font: "Times New Roman" });

const normal = (text, size = 24) =>
  new TextRun({ text, size, font: "Times New Roman" });

const italic = (text, size = 24) =>
  new TextRun({ text, italics: true, size, font: "Times New Roman" });

const boldItalic = (text, size = 24) =>
  new TextRun({ text, bold: true, italics: true, size, font: "Times New Roman" });

const para = (runs, align = AlignmentType.JUSTIFIED, spacing = 360) =>
  new Paragraph({
    children: Array.isArray(runs) ? runs : [runs],
    alignment: align,
    spacing: { after: spacing, line: 360 },
  });

const emptyLine = () =>
  new Paragraph({ children: [new TextRun("")], spacing: { after: 120 } });

const centeredBold = (text, size = 24) =>
  para([bold(text, size)], AlignmentType.CENTER, 200);

const sectionTitle = (text) =>
  new Paragraph({
    children: [bold(text, 24)],
    alignment: AlignmentType.CENTER,
    spacing: { before: 400, after: 200, line: 360 },
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
    },
  });

const subTitle = (text) =>
  new Paragraph({
    children: [bold(text, 24)],
    alignment: AlignmentType.LEFT,
    spacing: { before: 300, after: 160, line: 360 },
  });

const quotePara = (text) =>
  new Paragraph({
    children: [italic(text, 22)],
    alignment: AlignmentType.JUSTIFIED,
    indent: { left: 720, right: 720 },
    spacing: { before: 160, after: 160, line: 320 },
  });

// ─────────────────────────────────────────────
// Tabela de dados do requerimento
// ─────────────────────────────────────────────
const tableRow = (campo, informacao) =>
  new TableRow({
    children: [
      new TableCell({
        children: [
          new Paragraph({
            children: [bold(campo, 22)],
            alignment: AlignmentType.LEFT,
            spacing: { after: 80 },
          }),
        ],
        shading: { type: ShadingType.CLEAR, fill: "E8E8E8" },
        width: { size: 40, type: WidthType.PERCENTAGE },
      }),
      new TableCell({
        children: [
          new Paragraph({
            children: [normal(informacao, 22)],
            alignment: AlignmentType.LEFT,
            spacing: { after: 80 },
          }),
        ],
        width: { size: 60, type: WidthType.PERCENTAGE },
      }),
    ],
  });

const dadosTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    tableRow("Nome da Segurada", "Micheli Petroski Fragoso"),
    tableRow("Nome da Criança", "Maria Gabrieli Lourenço Fragoso"),
    tableRow("Data de Nascimento da Criança", "13/03/2024"),
    tableRow("Número do Benefício (NB)", "209.147.265-9"),
    tableRow("Data do Requerimento (DER)", "14/02/2025"),
    tableRow("Motivo do Indeferimento", "Falta de comprovação da qualidade de segurada especial"),
  ],
});

// ─────────────────────────────────────────────
// Documento
// ─────────────────────────────────────────────
const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "Times New Roman", size: 24 },
      },
    },
  },
  sections: [
    {
      properties: {
        page: {
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 },
        },
      },
      children: [
        // ── Endereçamento ──────────────────────────────────────────
        new Paragraph({
          children: [
            bold(
              "EXCELENTÍSSIMO(A) SENHOR(A) DOUTOR(A) JUIZ(A) FEDERAL DO JUIZADO ESPECIAL FEDERAL DE HUMAITÁ – AMAZONAS",
              24
            ),
          ],
          alignment: AlignmentType.CENTER,
          spacing: { after: 600, line: 360 },
        }),

        emptyLine(),
        emptyLine(),
        emptyLine(),

        // ── Qualificação da Autora ─────────────────────────────────
        para([
          bold("MICHELI PETROSKI FRAGOSO"),
          normal(
            ", brasileira, solteira, portadora da Cédula de Identidade RG n.º 32.016.174 SSP/MT, inscrita no CPF sob o n.º 086.877.461-80, residente e domiciliada na BR 319, KM 170, Zona Rural, Humaitá/AM, CEP 69.800-000, por intermédio de sua advogada infra-assinada, vem, respeitosamente, à presença de Vossa Excelência, propor a presente:"
          ),
        ]),

        emptyLine(),
        emptyLine(),

        // ── Título da Ação ─────────────────────────────────────────
        new Paragraph({
          children: [bold("AÇÃO DE CONCESSÃO DE SALÁRIO-MATERNIDADE RURAL", 28)],
          alignment: AlignmentType.CENTER,
          spacing: { before: 200, after: 200, line: 360 },
        }),

        emptyLine(),

        // ── Réu ───────────────────────────────────────────────────
        para([
          normal("Em face do "),
          bold("INSTITUTO NACIONAL DO SEGURO SOCIAL – INSS"),
          normal(
            ", autarquia federal, inscrita no CNPJ sob o n.º 29.979.036/0001-40, com sede na Rua Dep. Francisco Monteiro Neto, Humaitá/AM, CEP 69.800-000, pelos seguintes fundamentos fáticos e jurídicos:"
          ),
        ]),

        emptyLine(),
        emptyLine(),

        // ══════════════════════════════════════════════════════════
        // 1. JUSTIÇA GRATUITA
        // ══════════════════════════════════════════════════════════
        sectionTitle("I – DA JUSTIÇA GRATUITA"),

        para([
          normal(
            "A Autora é trabalhadora rural em regime de economia familiar, com renda limitada e voltada à subsistência, declarando, para os devidos fins e sob as penas da lei, ser pessoa hipossuficiente, sem condições de arcar com as custas processuais e honorários advocatícios sem prejuízo do próprio sustento e de sua família."
          ),
        ]),

        para([
          normal(
            "O benefício encontra amparo no art. 5.º, inciso LXXIV, da Constituição Federal de 1988, bem como no art. 98 do Código de Processo Civil (Lei n.º 13.105/2015). Requer-se, portanto, a concessão dos benefícios da Justiça Gratuita, isentando a Requerente do pagamento das custas processuais e demais despesas decorrentes do presente feito."
          ),
        ]),

        para([
          normal(
            "Registra-se que, com o advento do CPC/2015, a Lei n.º 1.060/50 foi parcialmente revogada, sendo o art. 98 e seguintes do CPC o regramento atualmente aplicável à gratuidade de justiça."
          ),
        ]),

        emptyLine(),

        // ══════════════════════════════════════════════════════════
        // 2. DOS FATOS
        // ══════════════════════════════════════════════════════════
        sectionTitle("II – DOS FATOS"),

        para([
          normal(
            "A Autora, Sra. Micheli Petroski Fragoso, reside em união estável com seu companheiro, Sr. ROBSON LOURENÇO DE SOUZA, na zona rural do Município de Humaitá/AM. Antes de se estabelecerem naquele município, o casal residia na propriedade dos avós paternos da criança, em Rondolândia/MT, endereço que consta na certidão de nascimento."
          ),
        ]),

        para([
          normal(
            "Desde agosto de 2022, o casal reside e trabalha na propriedade rural do sogro da Autora (avô paterno da criança), Sr. PAULO GABRIEL LOURENÇO, onde todos se dedicam à atividade agropecuária em regime de economia familiar, nos termos do art. 11, VII, da Lei n.º 8.213/1991."
          ),
        ]),

        para([
          normal(
            "A Autora engravidou em junho de 2023, e o nascimento de sua filha, MARIA GABRIELI LOURENÇO FRAGOSO, ocorreu em 13 de março de 2024. Em 14/02/2025, a Autora protocolou o requerimento administrativo de salário-maternidade rural perante o INSS (Protocolo n.º 1186879712). Contudo, a autarquia previdenciária indeferiu o pedido sob a alegação de \"ausência de comprovação da qualidade de segurada especial na data do afastamento\"."
          ),
        ]),

        para([
          normal(
            "Tal decisão administrativa é manifestamente equivocada e ignora a realidade fática e documental apresentada. A Autora sempre residiu em área rural, em terras integrantes do núcleo familiar de seu companheiro, onde exerce atividade agropecuária indispensável à subsistência. O fato de ter dado à luz em Cacoal/RO não descaracteriza sua condição de segurada especial, uma vez que o deslocamento se deu exclusivamente em razão da busca de melhor estrutura hospitalar, mantendo-se o vínculo com a atividade rural em Humaitá/AM."
          ),
        ]),

        para([
          normal(
            "Importa ressaltar que a residência da Autora não possui energia elétrica convencional, utilizando apenas placa solar, e os comprovantes de endereço estão em nome de seu sogro, o que é absolutamente comum em áreas rurais e não pode constituir óbice à comprovação de sua residência e atividade. Ademais, o acompanhamento pré-natal da Autora foi realizado integralmente em Humaitá/AM, conforme comprovado pela Ficha de Cadastro do Usuário da Secretaria Municipal de Saúde de Humaitá/AM, datada de 18/08/2023, onde consta seu endereço rural (BR 319, KM 170) e a anotação \"gestante\"."
          ),
        ]),

        emptyLine(),

        subTitle("2.1. Dados do Requerimento Administrativo"),

        dadosTable,

        emptyLine(),
        emptyLine(),

        // ══════════════════════════════════════════════════════════
        // 3. DO DIREITO
        // ══════════════════════════════════════════════════════════
        sectionTitle("III – DO DIREITO"),

        para([
          normal(
            "O direito reclamado pela Requerente encontra amparo na Constituição Federal de 1988, na legislação previdenciária infraconstitucional e na jurisprudência consolidada dos Tribunais Superiores, conforme se demonstra a seguir."
          ),
        ]),

        subTitle("3.1. Do Fundamento Constitucional"),

        para([
          normal(
            "A Constituição Federal de 1988 estabelece, em seu art. 201, inciso II, a proteção à maternidade como um dos pilares da Previdência Social:"
          ),
        ]),

        quotePara(
          "Art. 201. A previdência social será organizada sob a forma do Regime Geral de Previdência Social, de caráter contributivo e de filiação obrigatória, observados critérios que preservem o equilíbrio financeiro e atuarial, e atenderá, na forma da lei, a: [...] II – proteção à maternidade, especialmente à gestante;"
        ),

        para([
          normal(
            "No mesmo sentido, o art. 7.º, inciso XVIII, da Carta Magna garante a licença à gestante, sem prejuízo do emprego e do salário, com a duração de cento e vinte dias:"
          ),
        ]),

        quotePara(
          "Art. 7.º São direitos dos trabalhadores urbanos e rurais, além de outros que visem à melhoria de sua condição social: [...] XVIII – licença à gestante, sem prejuízo do emprego e do salário, com duração de cento e vinte dias."
        ),

        subTitle("3.2. Do Fundamento Legal – Lei n.º 8.213/1991"),

        para([
          normal(
            "No âmbito da legislação infraconstitucional, a Lei n.º 8.213/1991 define a figura da segurada especial e os requisitos para a concessão do benefício:"
          ),
        ]),

        quotePara(
          "Art. 11. São segurados obrigatórios da previdência social as seguintes pessoas físicas: [...] VII – como segurado especial: a pessoa física residente no imóvel rural ou em aglomerado urbano ou rural próximo a ele que, individualmente ou em regime de economia familiar, ainda que com o auxílio eventual de terceiro, na condição de: a) produtor, seja proprietário, usufrutuário, possuidor, assentado, parceiro ou meeiro outorgados, comodatário ou arrendatário rurais, que explore atividade agropecuária em área de até 4 (quatro) módulos fiscais; [...] c) cônjuge ou companheiro, bem como filho maior de 16 (dezesseis) anos de idade ou a este equiparado, do segurado de que tratam as alíneas a e b deste inciso, que, comprovadamente, trabalhem com o grupo familiar respectivo."
        ),

        para([
          normal(
            "Especificamente sobre o salário-maternidade, o art. 71 e seguintes da Lei n.º 8.213/1991, combinado com o art. 39, parágrafo único, dispõem que o benefício é devido à segurada especial que comprove o exercício de atividade rural nos dez meses imediatamente anteriores à data do parto, ainda que de forma descontínua:"
          ),
        ]),

        quotePara(
          "Art. 39, parágrafo único. Para a segurada especial fica garantida a concessão do salário-maternidade no valor de 1 (um) salário mínimo, desde que comprove o exercício de atividade rural, ainda que de forma descontínua, nos 10 (dez) meses imediatamente anteriores à data do parto ou do requerimento do benefício, quando requerido antes do parto."
        ),

        para([
          normal(
            "Diferentemente de normas temporárias anteriores, o ordenamento jurídico vigente não prevê prazo decadencial de 180 dias para o requerimento do salário-maternidade. O direito ao benefício submete-se apenas ao prazo prescricional de 5 (cinco) anos, conforme pacificado pelos Tribunais Superiores."
          ),
        ]),

        subTitle("3.3. Do Início de Prova Material em Nome do Grupo Familiar"),

        para([
          normal(
            "É pacífico o entendimento jurisprudencial de que documentos em nome de outros membros do grupo familiar (marido, pai ou sogro) servem como início de prova material para a trabalhadora rural. Nesse sentido, a Súmula n.º 73 do Tribunal Regional Federal da 4.ª Região e a jurisprudência consolidada do Superior Tribunal de Justiça (REsp 1.304.479/SP, submetido ao rito dos recursos repetitivos) reafirmam que o trabalho da mulher no campo é, via de regra, exercido em regime de colaboração mútua com o grupo familiar."
          ),
        ]),

        para([
          normal(
            "Para comprovar sua qualidade de segurada especial, a Autora apresenta os seguintes documentos:"
          ),
        ]),

        new Paragraph({
          children: [
            bold("a) "),
            bold("Contrato Particular de Compra e Venda (21/10/2022): "),
            normal(
              "Comprova a aquisição do \"Sítio Santo Antônio\", Gleba Acará, em Humaitá/AM, pelo sogro da Autora, Sr. Paulo Gabriel Lourenço. O documento possui firma reconhecida em cartório na mesma data, demonstrando a posse da terra pelo núcleo familiar muito antes do período de carência."
            ),
          ],
          alignment: AlignmentType.JUSTIFIED,
          indent: { left: 360 },
          spacing: { after: 200, line: 360 },
        }),

        new Paragraph({
          children: [
            bold("b) "),
            bold("Guia de Trânsito Animal – GTA (22/01/2024): "),
            normal(
              "Emitida pela ADAF/Amazonas, em nome do sogro da Autora, comprovando a movimentação de bovinos (fêmeas adultas) para fins de recria. Este documento situa-se exatamente dentro do período de carência (10 meses anteriores ao parto ocorrido em março/2024)."
            ),
          ],
          alignment: AlignmentType.JUSTIFIED,
          indent: { left: 360 },
          spacing: { after: 200, line: 360 },
        }),

        new Paragraph({
          children: [
            bold("c) "),
            bold("Nota Fiscal de Aquisição de Ferramentas Rurais (22/05/2024): "),
            normal(
              "Nota de aquisição de motosserra profissional em nome do companheiro da Autora (Robson Lourenço de Souza), reforçando o vínculo com a localidade rural (BR 319, KM 170, Realidade/AM)."
            ),
          ],
          alignment: AlignmentType.JUSTIFIED,
          indent: { left: 360 },
          spacing: { after: 200, line: 360 },
        }),

        new Paragraph({
          children: [
            bold("d) "),
            bold("Certidão de Nascimento da Filha: "),
            normal(
              "Documento que registra a residência dos pais em zona rural, corroborando o vínculo da Autora com a atividade campesina."
            ),
          ],
          alignment: AlignmentType.JUSTIFIED,
          indent: { left: 360 },
          spacing: { after: 200, line: 360 },
        }),

        new Paragraph({
          children: [
            bold("e) "),
            bold("Ficha de Cadastro do Usuário – Secretaria Municipal de Saúde de Humaitá/AM (18/08/2023): "),
            normal(
              "Comprova o acompanhamento pré-natal da Autora na localidade rural, com registro de seu endereço (BR 319, KM 170) e a anotação \"gestante\", em data que se insere no período de carência."
            ),
          ],
          alignment: AlignmentType.JUSTIFIED,
          indent: { left: 360 },
          spacing: { after: 200, line: 360 },
        }),

        subTitle("3.4. Da Carência e da Qualidade de Segurada Especial"),

        para([
          normal(
            "Considerando que o parto ocorreu em 13/03/2024, o período de carência de 10 (dez) meses compreende o intervalo entre maio de 2023 e março de 2024. O Contrato de Compra e Venda de outubro de 2022 e a GTA de janeiro de 2024 cobrem integralmente este lapso temporal, demonstrando que a Autora, inserida no regime de economia familiar, preenche todos os requisitos legais para a concessão do benefício."
          ),
        ]),

        subTitle("3.5. Da Jurisprudência Aplicável"),

        para([
          normal(
            "A jurisprudência dos Tribunais Regionais Federais é uníssona em garantir o direito ao salário-maternidade em casos análogos ao presente:"
          ),
        ]),

        quotePara(
          "PREVIDENCIÁRIO. SALÁRIO-MATERNIDADE. TRABALHADORA RURAL. SEGURADA ESPECIAL. ATIVIDADE CAMPESINA COMPROVADA. INÍCIO DE PROVA MATERIAL. PROVA TESTEMUNHAL. CONCESSÃO DO BENEFÍCIO. SENTENÇA DE IMPROCEDÊNCIA REFORMADA. 1. O salário-maternidade é devido à segurada especial, no valor de 01 (um) salário mínimo mensal durante 120 dias, a contar da data do parto ou dos 28 (vinte e oito) dias que o antecederam, desde que comprovado o exercício de atividade rural, ainda que de forma descontínua, nos dez meses imediatamente anteriores ao início do benefício (arts. 39, parágrafo único, e 71 c/c 25, da Lei n.º 8.213/1991). 2. Documentos em nome de membros do grupo familiar constituem início razoável de prova material, a ser complementado pela prova testemunhal. 3. Recurso provido."
        ),

        para([
          normal(
            "Nesse mesmo sentido, o Superior Tribunal de Justiça, no julgamento do REsp 1.304.479/SP (Tema 554), firmou a tese de que a prova material produzida por um membro do grupo familiar aproveita aos demais integrantes, desde que corroborada por prova testemunhal idônea."
          ),
        ]),

        emptyLine(),

        // ══════════════════════════════════════════════════════════
        // 4. DOS PEDIDOS
        // ══════════════════════════════════════════════════════════
        sectionTitle("IV – DOS PEDIDOS"),

        para([
          normal(
            "Diante de todo o exposto, requer a Vossa Excelência:"
          ),
        ]),

        new Paragraph({
          children: [
            bold("a) "),
            normal(
              "A concessão dos benefícios da Justiça Gratuita, nos termos do art. 98 do CPC, dada a hipossuficiência econômica da Autora, dispensando-a do recolhimento de custas e despesas processuais;"
            ),
          ],
          alignment: AlignmentType.JUSTIFIED,
          indent: { left: 360 },
          spacing: { after: 200, line: 360 },
        }),

        new Paragraph({
          children: [
            bold("b) "),
            normal(
              "A citação do INSS, na pessoa de seu representante legal, para, querendo, apresentar contestação no prazo legal;"
            ),
          ],
          alignment: AlignmentType.JUSTIFIED,
          indent: { left: 360 },
          spacing: { after: 200, line: 360 },
        }),

        new Paragraph({
          children: [
            bold("c) "),
            normal(
              "A produção de todas as provas admitidas em direito, especialmente a prova testemunhal, essencial para corroborar o início de prova material apresentado e detalhar a rotina de labor rural da Autora;"
            ),
          ],
          alignment: AlignmentType.JUSTIFIED,
          indent: { left: 360 },
          spacing: { after: 200, line: 360 },
        }),

        new Paragraph({
          children: [
            bold("d) "),
            normal(
              "O julgamento de PROCEDÊNCIA TOTAL da demanda, condenando o INSS a:"
            ),
          ],
          alignment: AlignmentType.JUSTIFIED,
          indent: { left: 360 },
          spacing: { after: 120, line: 360 },
        }),

        new Paragraph({
          children: [
            normal(
              "d.1) Conceder o benefício de Salário-Maternidade Rural (NB 209.147.265-9) à Autora, com DIB fixada na data do parto (13/03/2024);"
            ),
          ],
          alignment: AlignmentType.JUSTIFIED,
          indent: { left: 720 },
          spacing: { after: 120, line: 360 },
        }),

        new Paragraph({
          children: [
            normal(
              "d.2) Efetuar o pagamento das parcelas vencidas desde a data do parto (13/03/2024), devidamente corrigidas pelo INPC e acrescidas de juros de mora, nos termos da legislação vigente e da jurisprudência do STJ (Tema 905);"
            ),
          ],
          alignment: AlignmentType.JUSTIFIED,
          indent: { left: 720 },
          spacing: { after: 200, line: 360 },
        }),

        new Paragraph({
          children: [
            bold("e) "),
            normal(
              "A condenação do INSS ao pagamento de honorários advocatícios sucumbenciais, nos termos do art. 85 do CPC, fixados entre 10% e 20% sobre o valor da condenação."
            ),
          ],
          alignment: AlignmentType.JUSTIFIED,
          indent: { left: 360 },
          spacing: { after: 200, line: 360 },
        }),

        emptyLine(),

        // ── Valor da Causa ─────────────────────────────────────────
        para([
          bold("Valor da causa: "),
          normal(
            "R$ 5.648,00 (cinco mil, seiscentos e quarenta e oito reais), correspondente a 4 (quatro) salários mínimos vigentes na data do parto, para fins de alçada do Juizado Especial Federal."
          ),
        ]),

        emptyLine(),

        para([normal("Termos em que,")]),
        para([bold("Pede deferimento.")]),

        emptyLine(),
        emptyLine(),

        new Paragraph({
          children: [normal("Humaitá/AM, 06 de abril de 2026.")],
          alignment: AlignmentType.CENTER,
          spacing: { after: 600, line: 360 },
        }),

        emptyLine(),
        emptyLine(),
        emptyLine(),

        new Paragraph({
          children: [bold("HEVELLYN PRYSCYLLA MEDEIROS ROBERTO")],
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
        }),
        new Paragraph({
          children: [normal("Advogada – OAB/RO n.º 6.595")],
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("uploads/petição - michele salario maternidade - CORRIGIDA.docx", buffer);
console.log("Documento gerado com sucesso!");
