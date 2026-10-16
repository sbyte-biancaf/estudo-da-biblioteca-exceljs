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

  await workbook.xlsx.readFile('relatorio_teste_sinosbyte.xlsx');

  const wsConsolidado = workbook.getWorksheet("Consolidado");
  
  if (!wsConsolidado) {
    console.error("Aba 'Consolidado' não encontrada no arquivo!");
    return;
  }
  
  wsConsolidado.getCell("B8").value = "testeMudanca";
  console.log(wsConsolidado?.getRow(8).values)

  // const output = path.join(process.cwd(), "relatorio_teste_sinosbyte.xlsx");
  // await workbook.xlsx.writeFile(output);
  // console.log(`Relatório gerado com sucesso em: ${output}`);
}

gerarRelatorioHardcoded();