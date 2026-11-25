import ExcelJS from 'exceljs'

export async function exemploFicticio({ semPedidos = false, semEnderecoFiscal = false, semId = false, outraAba = false, comPacote = false } = {}) {
  const livro = new ExcelJS.Workbook()
  const aba = livro.addWorksheet(outraAba ? 'Resumo' : 'Vendas BR')
  aba.getCell('A2').value = 'Relatório fictício para testes'
  aba.getCell('A6').value = ' vendas '
  aba.mergeCells('A6:T6')
  aba.getCell('V6').value = 'ANÚNCIOS'
  aba.mergeCells('V6:AA6')
  aba.getCell('AB6').value = 'Faturamento   ao comprador'
  aba.mergeCells('AB6:AG6')
  aba.getCell('AH6').value = 'Compradores'
  aba.mergeCells('AH6:AO6')
  aba.getCell('AP6').value = 'Envios'
  aba.mergeCells('AP6:AV6')
  aba.getCell('AW6').value = 'Devoluções'
  aba.mergeCells('AW6:BI6')

  const colunas = {
    A: semId ? null : ' N.º  de  venda ',
    G: '  UNIDADES ',
    V: 'SKU',
    X: 'TÍTULO DO ANÚNCIO',
    AC: 'Dados pessoais ou da empresa',
    AE: semEnderecoFiscal ? null : 'Endereço',
    AK: 'Endereço',
    AT: 'Identificador do envio',
    AW: 'Unidades',
  }
  for (const [coluna, titulo] of Object.entries(colunas)) aba.getCell(coluna + '7').value = titulo

  if (!semPedidos) {
    const dados = [
      [8, '10001', 2, ' KIT   2 TEMPORIZADOR Timer Bivolt', 'Cliente Fictício', 'Rua de Teste, 1\nBloco 2'],
      [9, '10001', 1, 'Temporizador Timer Mecânico Tomada Bivolt', null, null],
      [11, '10002', '1', 'Timer Temporizador Digital 24 Horas', null, '=Endereço em texto'],
      [12, '10003', 1, 'Kit 12 Presilhas para Cabelo Original Premium', null, null],
      [13, '10004', 1, 'Pote de Vidro 640ml Transparente com Tampa Hermética Premium Original Oferta', '@Dados em texto', null],
      [14, '10005', 1, 'Kit 3 Temporizadores Timer Mecânico', null, null],
    ]
    for (const [linha, id, unidades, titulo, cliente, endereco] of dados) {
      if (!semId) aba.getCell('A' + linha).value = id
      aba.getCell('G' + linha).value = unidades
      aba.getCell('V' + linha).value = 'SKU-' + String(linha).padStart(2, '0')
      aba.getCell('X' + linha).value = titulo
      aba.getCell('AC' + linha).value = cliente
      aba.getCell('AE' + linha).value = endereco
      aba.getCell('AK' + linha).value = 'Endereço de comprador fictício'
      aba.getCell('AT' + linha).value = String(70000000000 + linha)
      aba.getCell('AW' + linha).value = 99
    }
  }
  if (comPacote) {
    aba.getCell('B7').value = 'Data de venda'
    aba.getCell('C7').value = 'Estado'
    aba.getCell('E7').value = 'Pacote de diversos produtos'
    aba.getCell('AH7').value = 'Comprador'
    aba.getCell('AM7').value = 'Estado'
    const pacote = [
      ['10006', 'Pacote de 3 produtos', ' ', ' ', ' ', 'Cliente do Pacote', 'Rua Fictícia, 2'],
      ['10007', 'Etiqueta impressa', 'Sim', 2, 'Pote de Vidro 640 ml', ' ', ' '],
      ['10008', 'Etiqueta impressa', 'Sim', 1, 'Kit 2 Temporizador Bivolt', ' ', ' '],
      ['10009', 'Etiqueta impressa', 'Sim', 1, 'Timer Temporizador Digital', ' ', ' '],
      ['10010', 'Etiqueta impressa', 'Sim', 1, 'Caneca Cerâmica', 'Cliente do Pacote', 'Rua Fictícia, 2'],
    ]
    pacote.forEach(([id, estado, flag, unidades, titulo, cliente, endereco], indice) => {
      const linha = 15 + indice
      const campos = { A: id, B: '15 de novembro de 2026 21:05 hs.', C: estado, E: flag, G: unidades, X: titulo, AC: cliente, AE: endereco,
        V: indice ? 'SKU-' + linha : ' ', AT: indice === 0 || indice === 4 ? String(70000000000 + linha) : ' ' }
      for (const [coluna, valor] of Object.entries(campos)) aba.getCell(coluna + linha).value = valor
    })
  }
  const rodape = comPacote ? 23 : 18
  aba.getCell('A' + rodape).value = 'Total geral'
  aba.getCell('G' + rodape).value = 125
  return Buffer.from(await livro.xlsx.writeBuffer())
}
