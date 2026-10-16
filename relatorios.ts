import ExcelJS from "exceljs";
import path from "node:path";

const meses = [
    { value: "01", label: "Janeiro" },
    { value: "02", label: "Fevereiro" },
    { value: "03", label: "Março" },
    { value: "04", label: "Abril" },
    { value: "05", label: "Maio" },
    { value: "06", label: "Junho" },
    { value: "07", label: "Julho" },
    { value: "08", label: "Agosto" },
    { value: "09", label: "Setembro" },
    { value: "10", label: "Outubro" },
    { value: "11", label: "Novembro" },
    { value: "12", label: "Dezembro" },
  ]

const mesAtual = 8

const colaboradores = [
    { nome: "Lucas Moraes", cargo: "Desenvolvedor", senioridade: "SENIOR", fornecedor: "SINOSBYTE", squad: "SUSTENTAÇÃO", lider: "Luciano Muller", alocacao: 1, valorHora: 130 },
    { nome: "Bianca Franzon", cargo: "Desenvolvedora", senioridade: "JUNIOR", fornecedor: "SINOSBYTE", squad: "SUSTENTAÇÃO", lider: "Luciano Muller", alocacao: 1, valorHora: 130 },
    { nome: "João Teste 3", cargo: "Desenvolvedor", senioridade: "SENIOR", fornecedor: "SINOSBYTE", squad: "SUSTENTAÇÃO", lider: "Luciano Muller", alocacao: 1, valorHora: 130 }
  ];

const horasLancadas = [
  { entrada: "09:00", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:00", observacoes: "Teste 1"},
  { entrada: "09:30", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:30", observacoes: "Teste 2"},
  { entrada: "07:45", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "15:05", observacoes: "Teste 3"},
]

async function gerarRelatorioHardcoded() {
  const workbook = new ExcelJS.Workbook();
  const wsConsolidado = workbook.addWorksheet("Consolidado");

  const r1 = wsConsolidado.addRow(["CONSOLIDADO — SINOSBYTE"]);
  r1.font = { bold: true, size: 16, color: { argb: "1F4E78" }}
  wsConsolidado.addRow(["Preencha o Mês e o Ano de referência abaixo — todas as abas se ajustam sozinhas."]);
  wsConsolidado.addRow([]);
  
  const row4 = wsConsolidado.addRow(["Mês de referência (1 a 12):", meses[mesAtual-1].value, "", "Ano de referência", 2026, "", "Mês:", meses[mesAtual-1].label]);
  row4.font = { bold: true };
  // row4.getCell(2).numFmt = "0";
  // row4.getCell(5).numFmt = "0";

  wsConsolidado.addRow([]);
  wsConsolidado.addRow(["Revise a coluna VALOR HORA (em azul) se houve reajuste. As colunas TOTAL DE HORAS e TOTAL R$ MENSAL são calculadas automaticamente."]);
  wsConsolidado.addRow([]);

  wsConsolidado.getColumn(1).width = 30;
  wsConsolidado.getColumn(2).width = 20;
  wsConsolidado.getColumn(3).width = 15;
  wsConsolidado.getColumn(4).width = 20;
  wsConsolidado.getColumn(5).width = 20;
  wsConsolidado.getColumn(6).width = 20;
  wsConsolidado.getColumn(7).width = 20;
  wsConsolidado.getColumn(8).width = 20;
  wsConsolidado.getColumn(9).width = 20;
  wsConsolidado.getColumn(10).width = 20;
  wsConsolidado.getColumn(11).width = 10;

  const headerConsolidado = ["COLABORADOR", "CARGO/FUNÇÃO", "SENIORIDADE", "FORNECEDOR", "SQUAD", "RESPONSÁVEL IMEDIATO", "% ALOCAÇÃO", "VALOR HORA", "TOTAL DE HORAS", "TOTAL R$ MENSAL", "MÊS"];
  const r8 = wsConsolidado.addRow(headerConsolidado);
  r8.font = { bold: true, color: { argb: "FFFFFF" } };
  r8.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "1F4E78" } };

  colaboradores.forEach((c, index) => {
    const rowNum = 9 + index;
    wsConsolidado.addRow([
      c.nome, c.cargo, c.senioridade, c.fornecedor, c.squad, c.lider, c.alocacao, c.valorHora, { formula: `='${c.nome}'!H40` }, { formula: `=H${rowNum}*I${rowNum}` }, meses[mesAtual-1].label
    ]);
  });

  const totalRowIndex = 9 + colaboradores.length;
  wsConsolidado.addRow(["", "", "", "", "", "", "", "TOTAL GERAL", "", { formula: `=SUM(J9:J${totalRowIndex-1})` } ]);

  console.log(wsConsolidado.getRow(8).values)
  // wsConsolidado.getColumn("H").numFmt = '"R$ "#,##0.00';
  // wsConsolidado.getColumn("J").numFmt = '"R$ "#,##0.00';
  // wsConsolidado.getColumn("G").numFmt = '0%';

  // Abas de cada Colaborador
  colaboradores.forEach((c) => {
    const wsPonto = workbook.addWorksheet(c.nome);

    const r1ponto = wsPonto.addRow([`REGISTRO DE PONTO — ${c.nome} — SINOSBYTE — ${meses[mesAtual-1].label}/${2026}`]);
    r1ponto.font = { bold: true, size: 16, color: { argb: "1F4E78" } };
    wsPonto.addRow(["Preencha apenas as células amarelas. O Total de Horas e o Total do Mês são calculados automaticamente"]);
    wsPonto.addRow([]);

    wsPonto.addRow(["Colaborador:", c.nome, "", "Cargo:", c.cargo, "", "Empresa:", "SINOSBYTE"]);
    wsPonto.addRow(["Squad:", c.squad, "", "Valor Hora (R$):", c.valorHora, "", "Mês/Ano:", `${meses[mesAtual-1].label}/2026`]);
    wsPonto.addRow([]);

    wsPonto.getColumn(1).width = 20;
    wsPonto.getColumn(2).width = 20;
    wsPonto.getColumn(3).width = 20;
    wsPonto.getColumn(4).width = 20;
    wsPonto.getColumn(5).width = 20;
    wsPonto.getColumn(6).width = 20;
    wsPonto.getColumn(7).width = 20;
    wsPonto.getColumn(8).width = 20;
    
    const cabecalhoPonto = ["Data", "Dia da Semana", "Entrada", "Saída Almoço", "Retorno Almoço", "Saída", "Total Horas", "Observações"];
    const r7Ponto = wsPonto.addRow(cabecalhoPonto);
    r7Ponto.font = { bold: true, color: { argb: "FFFFFF" } };
    r7Ponto.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "1F4E78" } };



    const diasDaSemana = ["Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "Sábado", "Domingo"];
    
    for (let dia = 1; dia <= 6; dia++) {
      const dataStr = `${dia}/6/2026`;
      const dataObj = new Date(2026, 5, dia);
      const diaSemanaIndex = (dataObj.getDay() + 6) % 7; 
      const nomeDiaSemana = diasDaSemana[diaSemanaIndex];
      const rowIdx = 8 + dia;
      const entrada = horasLancadas[(dia - 1) % horasLancadas.length].entrada;
      const saidaAlmoco = horasLancadas[(dia - 1) % horasLancadas.length].saidaAlmoco;
      const retornoAlmoco = horasLancadas[(dia - 1) % horasLancadas.length].retornoAlmoco;
      const saida = horasLancadas[(dia - 1) % horasLancadas.length].saida;

      const formulaHoras = `=IF(OR(C${rowIdx}="",D${rowIdx}="",E${rowIdx}="",F${rowIdx}=""),0,(D${rowIdx}-C${rowIdx})+(F${rowIdx}-E${rowIdx}))*24`;

      wsPonto.addRow([dataStr, nomeDiaSemana, entrada, saidaAlmoco, retornoAlmoco, saida, { formula: formulaHoras }, ""]);
    }

    wsPonto.addRow([]);
    const rTotalHoras = wsPonto.addRow(["", "", "", "", "", "TOTAL HORAS DO MÊS", { formula: `=SUM(G9:G38)` }]);
    rTotalHoras.font = { bold: true };
    
    // Acho que isso tá errado, mas fiz como tava no modelo inicial
    const rTotalRs = wsPonto.addRow(["", "", "", "", "", "TOTAL R$ MENSAL (estimado)", { formula: `=G40*E5` }]);
    rTotalRs.font = { bold: true };
    rTotalRs.getCell(7).numFmt = '"R$ "#,##0.00';

    wsPonto.getColumn("G").numFmt = '0.00';
    wsPonto.getColumn("E").numFmt = '"R$ "#,##0.00';
  });

  const output = path.join(process.cwd(), "relatorio_teste_sinosbyte.xlsx");
  await workbook.xlsx.writeFile(output);
  console.log(`Relatório gerado com sucesso em: ${output}`);
}

gerarRelatorioHardcoded();