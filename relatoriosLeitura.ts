import ExcelJS from "exceljs";
import path from "node:path";

const mesAtual = 5

const colaboradores = [
    { nome: "Lucas Moraes", cargo: "Desenvolvedor", senioridade: "SENIOR", fornecedor: "SINOSBYTE", squad: "SUSTENTAÇÃO", lider: "Luciano Muller", alocacao: 1, valorHora: 130 },
    { nome: "Bianca Franzon", cargo: "Desenvolvedora", senioridade: "JUNIOR", fornecedor: "SINOSBYTE", squad: "SUSTENTAÇÃO", lider: "Luciano Muller", alocacao: 1, valorHora: 100 },
    { nome: "Guilherme Fagundes", cargo: "Desenvolvedor", senioridade: "SENIOR", fornecedor: "SINOSBYTE", squad: "SUSTENTAÇÃO", lider: "Luciano Muller", alocacao: 1, valorHora: 110 },
    { nome: "Joszef Barrionuevo", cargo: "Desenvolvedor", senioridade: "SENIOR", fornecedor: "SINOSBYTE", squad: "SUSTENTAÇÃO", lider: "Luciano Muller", alocacao: 1, valorHora: 140 }
  ];

const horasLancadas = [
  { colaborador: "Bianca Franzon", dia: 1,entrada: "09:00", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:00", observacoes: "Teste 1"},
  { colaborador: "Bianca Franzon", dia: 2,entrada: "", saidaAlmoco: "", retornoAlmoco: "13:00", saida: "18:00", observacoes: "Teste 31"},
  { colaborador: "Bianca Franzon", dia: 3,entrada: "", saidaAlmoco: "", retornoAlmoco: "13:00", saida: "18:00", observacoes: "Teste 1"},
  { colaborador: "Bianca Franzon", dia: 4,entrada: "09:30", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:00", observacoes: "Teste 31"},
  { colaborador: "Bianca Franzon", dia: 7,entrada: "09:30", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:00", observacoes: "Teste 43211"},
  { colaborador: "Bianca Franzon", dia: 8,entrada: "09:00", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:00", observacoes: "Teste 1"},
  { colaborador: "Bianca Franzon", dia: 9,entrada: "09:00", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:00", observacoes: "Teste 51"},
  { colaborador: "Bianca Franzon", dia: 10,entrada: "09:30", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:00", observacoes: "Teste 361"},
  { colaborador: "Bianca Franzon", dia: 11,entrada: "09:00", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:00", observacoes: "Teste 21"},
  { colaborador: "Bianca Franzon", dia: 14,entrada: "09:30", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:00", observacoes: "Teste 41"},
  { colaborador: "Bianca Franzon", dia: 15,entrada: "09:00", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:00", observacoes: "Teste 221"},
  { colaborador: "Lucas Moraes",dia: 4, entrada: "09:30", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:30", observacoes: "Teste 2"},
  { colaborador: "Bianca Franzon", dia: 1, entrada: "09:30", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:30", observacoes: "Teste 2"},
  { colaborador: "Lucas Moraes",dia: 2, entrada: "09:30", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:30", observacoes: "Teste 2"},
  { colaborador: "Lucas Moraes",dia: 3, entrada: "09:30", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:30", observacoes: "Teste 2"},
  { colaborador: "Bianca Franzon", dia: 4, entrada: "09:30", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:30", observacoes: "Teste 2"},
  { colaborador: "Bianca Franzon", dia: 5, entrada: "07:45", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "15:05", observacoes: "Teste 3"},
  { colaborador: "Joszef Barrionuevo", dia: 1, entrada: "08:45", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:35", observacoes: "Test"},
  { colaborador: "Joszef Barrionuevo", dia: 3, entrada: "09:30", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:10", observacoes: "Test"},
  { colaborador: "Joszef Barrionuevo", dia: 4, entrada: "08:45", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:35", observacoes: "Test"},
  { colaborador: "Joszef Barrionuevo", dia: 5, entrada: "09:30", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:35", observacoes: "Test"},
  { colaborador: "Joszef Barrionuevo", dia: 6, entrada: "09:30", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:10", observacoes: "Test"},
  { colaborador: "Joszef Barrionuevo", dia: 7, entrada: "08:45", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:35", observacoes: "Test"},
  { colaborador: "Joszef Barrionuevo", dia: 8, entrada: "09:30", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:10", observacoes: "Test"},
  { colaborador: "Joszef Barrionuevo", dia: 12, entrada: "09:30", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:35", observacoes: "Test"},
  { colaborador: "Joszef Barrionuevo", dia: 13, entrada: "08:45", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:10", observacoes: "Test"},
  { colaborador: "Joszef Barrionuevo", dia: 14, entrada: "09:30", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:35", observacoes: "Test"},
  { colaborador: "Joszef Barrionuevo", dia: 15, entrada: "08:45", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:10", observacoes: "Test"},
  { colaborador: "Guilherme Fagundes", dia: 1, entrada: "08:45", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "19:00", observacoes: "Test"},
  { colaborador: "Guilherme Fagundes", dia: 4, entrada: "09:00", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:10", observacoes: "Test"},
  { colaborador: "Guilherme Fagundes", dia: 5, entrada: "08:45", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "17:40", observacoes: "Test"},
  { colaborador: "Guilherme Fagundes", dia: 6, entrada: "09:00", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:10", observacoes: "Test"},
  { colaborador: "Guilherme Fagundes", dia: 7, entrada: "09:00", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "19:00", observacoes: "Test"},
  { colaborador: "Guilherme Fagundes", dia: 8, entrada: "08:45", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "19:00", observacoes: "Test"},
  { colaborador: "Guilherme Fagundes", dia: 11, entrada: "09:00", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:10", observacoes: "Test"},
  { colaborador: "Guilherme Fagundes", dia: 12, entrada: "09:00", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "19:00", observacoes: "Test"},
  { colaborador: "Guilherme Fagundes", dia: 13, entrada: "08:45", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "17:40", observacoes: "Test"},
  { colaborador: "Guilherme Fagundes", dia: 14, entrada: "08:45", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:10", observacoes: "Test"},
  { colaborador: "Guilherme Fagundes", dia: 15, entrada: "08:45", saidaAlmoco: "12:00", retornoAlmoco: "13:00", saida: "18:10", observacoes: "Test"},
]

// Auxiliar para copiar estilo de uma célula para outra (preserva bordas, fontes, alinhamentos e formatos)
function copiarEstiloCelula(celulaOrigem: ExcelJS.Cell, celulaDestino: ExcelJS.Cell) {
  if (celulaOrigem.style) {
    celulaDestino.style = JSON.parse(JSON.stringify(celulaOrigem.style));
  }
}

async function gerarRelatorioHardcoded() {
  const workbook = new ExcelJS.Workbook();

  await workbook.xlsx.readFile('SINOSBYTE-TEMPLATE-Vazio.xlsx');

  const wsConsolidado = workbook.getWorksheet("Consolidado");
  if (wsConsolidado) {
    wsConsolidado.getCell("B4").value = mesAtual; 
    wsConsolidado.getCell("E4").value = 2026;

    // Insere os colaboradores a partir da linha 9 (conforme estruturado no template)
    colaboradores.forEach((c, index) => {
      const rowNum = 9 + index;
      const linhaModelo = wsConsolidado.getRow(9);
      const linhaAtual = wsConsolidado.getRow(rowNum);

      
      linhaAtual.getCell("A").value = c.nome;
      linhaAtual.getCell("B").value = c.cargo;
      linhaAtual.getCell("C").value = c.senioridade;
      linhaAtual.getCell("D").value = c.fornecedor;
      linhaAtual.getCell("E").value = c.squad;
      linhaAtual.getCell("F").value = c.lider;
      linhaAtual.getCell("G").value = c.alocacao;
      linhaAtual.getCell("H").value = c.valorHora;

      
      // Mantém a fórmula que puxa o total de horas da aba do colaborador correspondente (ex: 'Lucas Moraes'!H40)
      linhaAtual.getCell("I").value = { formula: `'${c.nome}'!G40` };
      linhaAtual.getCell("J").value = { formula: `H${rowNum}*I${rowNum}` };
      linhaAtual.getCell("K").value = { formula: `$H$4` };

      linhaAtual.eachCell({includeEmpty: true }, (celulaAlvo, colNumber) => {
        const celulamodelo = linhaModelo.getCell(colNumber);
        copiarEstiloCelula(celulamodelo, celulaAlvo)
      })
    });

    const totalRowIndex = 9 + colaboradores.length;
    const linhaTotal = wsConsolidado.getRow(totalRowIndex);

    linhaTotal.getCell("H").value = "TOTAL GERAL";
    linhaTotal.getCell("J").value = { formula: `SUM(J9:J${totalRowIndex - 1})` };
    
    linhaTotal.getCell("J").numFmt = '"R$ "#,##0.00';
    
    linhaTotal.font = { bold: true };
  }

  const abaModelo = workbook.getWorksheet(2);

  for (const c of colaboradores){
    let wsPonto = workbook.getWorksheet(c.nome);


    if (!wsPonto && abaModelo) {
      wsPonto = workbook.addWorksheet(c.nome);

      // DEPOIS TESTA AQUI
      abaModelo.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        const newRow = wsPonto!.getRow(rowNumber);
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const newCell = newRow.getCell(colNumber);
          newCell.value = cell.value;
          copiarEstiloCelula(cell, newCell);
        });
      });

      // AQUI TBM
      abaModelo.columns?.forEach((col, i) => {
        if (wsPonto!.getColumn(i + 1)) {
          wsPonto!.getColumn(i + 1).width = col.width;
        }
      });
    }

    if (wsPonto) {
      wsPonto.mergeCells("A1:H1");
      wsPonto.mergeCells("A2:H2");

      wsPonto.getCell("A1").value = { formula: `"REGISTRO DE PONTO — ${c.nome} — SINOSBYTE — " & Consolidado!$H$4 & "/" & Consolidado!$E$4` };

      wsPonto.getCell("B4").value = c.nome;
      wsPonto.getCell("E4").value = c.cargo;
      wsPonto.getCell("H4").value = "SINOSBYTE";
      wsPonto.getCell("B5").value = c.squad;
      wsPonto.getCell("E5").value = c.valorHora;


      const horasDoColaborador = horasLancadas.filter(h => h.colaborador === c.nome);

      horasDoColaborador.forEach((ponto) => {
        const rowIdx = 7 + ponto.dia;
        const row = wsPonto!.getRow(rowIdx);

        row.getCell("C").value = ponto.entrada || "00:00";
        row.getCell("D").value = ponto.saidaAlmoco || "00:00";
        row.getCell("E").value = ponto.retornoAlmoco || "00:00";
        row.getCell("F").value = ponto.saida || "00:00";
        if (ponto.observacoes) row.getCell("H").value = ponto.observacoes;
      })
    }

  }

  // Se a aba modelo genérica ainda existir no workbook final, podemos removê-la para ficar limpo
  workbook.removeWorksheet(2);



  const output = path.join(process.cwd(), "relatorio_teste_sinosbyte.xlsx");
  await workbook.xlsx.writeFile(output);
  console.log(`Relatório gerado com sucesso em: ${output}`);
}

gerarRelatorioHardcoded();