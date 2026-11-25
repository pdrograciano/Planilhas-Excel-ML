import assert from 'node:assert/strict'
import { readFile, writeFile } from 'node:fs/promises'
import { test } from 'node:test'
import ExcelJS from 'exceljs'
import { PDFDocument } from 'pdf-lib'
import { nomeCurto } from '../src/formatador/item.ts'
import { formatarPlanilha, lerEtiquetas, renumerarPlanilhaPorItens } from '../src/formatador/planilha.ts'
import { gerarPdfEtiquetas } from '../src/formatador/etiquetas.ts'
import { exemploFicticio } from './exemplo-ficticio.mjs'

async function ler(bytes) {
  const arquivo = new ExcelJS.Workbook()
  await arquivo.xlsx.load(bytes)
  return arquivo
}

test('abrevia temporizadores, kits de outras quantidades e produtos com medidas', () => {
  assert.equal(nomeCurto(' KIT   2 Temporizador Timer'), 'Kit 02 Timer')
  assert.equal(nomeCurto('Temporizador Timer Mecânico Tomada Bivolt'), 'Timer Unitário')
  assert.equal(nomeCurto('Timer Temporizador analógico 24 Horas'), 'Timer Unitário')
  assert.equal(nomeCurto('Timer 3 unidades'), 'Kit 3 Timer')
  assert.equal(nomeCurto('Kit 3 Temporizadores Timer'), 'Kit 3 Timer')
  assert.equal(nomeCurto('Temporizador Mecânico Kit 3'), 'Kit 3 Timer')
  assert.equal(nomeCurto('Par de Tapetes'), 'Kit 2 Tapetes')
  assert.equal(nomeCurto('Kit 2 Potes de Vidro'), 'Kit 2 Potes Vidro')
  assert.equal(nomeCurto('Kit 12 Presilhas Para Cabelo Premium'), 'Kit 12 Presilhas')
  assert.equal(nomeCurto('Pote de Vidro 640ml Transparente com Tampa Hermética Premium'), 'Pote Vidro 640 ml')
  assert.equal(nomeCurto(''), '')
})

test('usa os nomes de item cadastrados para mini potes, Trident e porta-ampolas', async () => {
  const original = await ler(await exemploFicticio())
  const vendas = original.getWorksheet('Vendas BR')
  const titulos = [
    ' KIT   COM 4 Mini Potes para Temperos',
    '4 Mini Potes com Tampa',
    'Trident Chiclete de Menta',
    'TRIDENT CHICLETES Hortelã',
    'Porta 4 Ampolas com Tampa',
    'Kit 2 Porta Ampola para Viagem',
  ]
  for (const [indice, linha] of [8, 9, 11, 12, 13, 14].entries()) {
    vendas.getCell(linha, 24).value = titulos[indice]
  }
  const { bytes, porItens } = await formatarPlanilha(await original.xlsx.writeBuffer())
  const normal = (await ler(bytes)).getWorksheet('Planilha1')
  const ordenada = (await ler(porItens)).getWorksheet('Planilha1')
  assert.deepEqual(normal.getColumn('D').values.slice(2), titulos)
  assert.deepEqual(normal.getColumn('C').values.slice(2), [
    'Mini Potes', 'Mini Potes', 'Trident', 'Trident', 'Porta Ampola Unitário', 'Kit 2 Porta Ampola',
  ])
  assert.deepEqual(ordenada.getColumn('C').values.slice(2), [
    'Kit 2 Porta Ampola', 'Mini Potes', 'Mini Potes', 'Porta Ampola Unitário', 'Trident', 'Trident',
  ])
})

test('gera A:F a partir de Vendas BR e mantém posições, valores, tipos e estilos', async () => {
  const original = await exemploFicticio()
  const resultado = await formatarPlanilha(original)
  assert.equal(resultado.linhas, 6)
  const livro = await ler(resultado.bytes)
  assert.deepEqual(livro.worksheets.map((aba) => aba.name), ['Planilha1'])
  const aba = livro.getWorksheet('Planilha1')
  assert.equal(aba.rowCount, 7)
  assert.equal(aba.columnCount, 6)
  assert.deepEqual(aba.getRow(1).values.slice(1), ['ORDEM', 'Unidades', 'ITEM', 'Título do anúncio', 'Dados pessoais ou da empresa', 'Endereço'])
  assert.deepEqual(aba.getColumn('A').values.slice(2), [1, 2, 3, 4, 5, 6])
  assert.deepEqual(aba.getColumn('B').values.slice(2), [2, 1, 1, 1, 1, 1])
  assert.deepEqual(aba.getColumn('C').values.slice(2), [
    'Kit 02 Timer', 'Timer Unitário', 'Timer Unitário', 'Kit 12 Presilhas', 'Pote Vidro 640 ml', 'Kit 3 Timer',
  ])
  assert.equal(aba.getCell('D2').value, ' KIT   2 TEMPORIZADOR Timer Bivolt')
  assert.equal(aba.getCell('E2').value, 'Cliente Fictício')
  assert.equal(aba.getCell('E3').value, '')
  assert.equal(aba.getCell('F2').value, 'Rua de Teste, 1\nBloco 2')
  assert.equal(aba.getCell('F4').value, '=Endereço em texto')
  assert.equal(aba.getCell('E6').value, '@Dados em texto')
  assert.equal(aba.getCell('A2').font.color.argb, 'FFFF0000')
  assert.equal(aba.getCell('A2').fill.fgColor.argb, 'FFC0E6F5')
  assert.equal(aba.getCell('B2').font.bold, true)
  assert.equal(aba.getCell('C2').alignment.shrinkToFit, true)
  assert.equal(aba.getCell('F2').alignment.wrapText, true)
  assert.equal(aba.getCell('F2').border.bottom.style, 'thin')
  assert.deepEqual(aba.columns.map((coluna) => coluna.width), [12.6640625, 12.6640625, 18.6640625, 68.5546875, 36.21875, 82.6640625])
  assert.equal(aba.getRow(1).height, 60)
  assert.equal(aba.getRow(2).height, 30)
  assert.equal(aba.getRow(3).height, 22.5)
  assert.deepEqual(aba.getRow(1).values.slice(1).map((_, i) => aba.getCell(1, i + 1).fill.fgColor.argb), [
    'FFF1C232', 'FFF1C232', 'FFF1C232', 'FFF6B26B', 'FFFFD966', 'FFFFD966',
  ])
})

test('importa relatório sem número de venda quando as demais células identificam os pedidos', async () => {
  const resultado = await formatarPlanilha(await exemploFicticio({ semId: true }))
  assert.equal(resultado.linhas, 6)
})

test('ordena por ITEM mantendo as linhas completas e a sequência original dentro dos grupos', async () => {
  const { bytes, porItens } = await formatarPlanilha(await exemploFicticio())
  const origem = (await ler(bytes)).getWorksheet('Planilha1')
  const ordenada = (await ler(porItens)).getWorksheet('Planilha1')
  assert.deepEqual(ordenada.getColumn('C').values.slice(2), [
    'Kit 02 Timer', 'Kit 12 Presilhas', 'Kit 3 Timer', 'Pote Vidro 640 ml', 'Timer Unitário', 'Timer Unitário',
  ])
  assert.deepEqual(ordenada.getColumn('A').values.slice(2), [1, 2, 3, 4, 5, 6])
  const posicoesOriginais = [1, 2, 5, 7, 6, 3, 4]
  for (const [indice, posicao] of posicoesOriginais.entries()) {
    const linha = ordenada.getRow(indice + 1)
    assert.deepEqual(linha.values.slice(2), origem.getRow(posicao).values.slice(2))
    assert.equal(linha.height, origem.getRow(posicao).height)
  }
})

test('rejeita arquivo ilegível, aba ausente, coluna fiscal ausente e relatório sem dados', async () => {
  await assert.rejects(formatarPlanilha(new TextEncoder().encode('arquivo inválido').buffer), /ler o arquivo/)
  await assert.rejects(formatarPlanilha(await exemploFicticio({ outraAba: true })), /aba Vendas BR não foi encontrada/)
  await assert.rejects(formatarPlanilha(await exemploFicticio({ semEnderecoFiscal: true })), /Coluna obrigatória ausente.*Endereço/)
  await assert.rejects(formatarPlanilha(await exemploFicticio({ semPedidos: true })), /relatório não contém pedidos/)
})

test('mantém pacote com três pedidos junto mesmo quando os itens pertencem a grupos A–Z diferentes', async () => {
  const { bytes, porItens, linhas } = await formatarPlanilha(await exemploFicticio({ comPacote: true }))
  assert.equal(linhas, 10)
  const origem = (await ler(bytes)).getWorksheet('Planilha1')
  const ordenada = (await ler(porItens)).getWorksheet('Planilha1')
  for (const [aba, resumo] of [[origem, 8], [ordenada, 7]]) {
    assert.equal(aba.rowCount, 12)
    assert.equal(aba.getCell(resumo, 1).value, null)
    assert.equal(aba.getCell(resumo, 2).value, null)
    assert.equal(aba.getCell(resumo, 3).value, '')
    assert.equal(aba.getCell(resumo, 4).value, '(3 pedidos feitos)')
    assert.equal(aba.getCell(resumo, 5).value, 'Cliente do Pacote')
    assert.equal(aba.getCell(resumo, 6).value, 'Rua Fictícia, 2')
    assert.deepEqual([1, 2, 3].map((i) => aba.getCell(resumo + i, 3).value), [
      'Pote Vidro 640 ml', 'Kit 02 Timer', 'Timer Unitário',
    ])
    assert.deepEqual([1, 2, 3].map((i) => aba.getCell(resumo + i, 2).value), [2, 1, 1])
    assert.deepEqual(aba.getColumn('A').values.slice(2).filter(Number.isInteger), [1, 2, 3, 4, 5, 6, 7, 8, 9, 10])
  }
  assert.equal(origem.getCell('D12').value, 'Caneca Cerâmica')
  assert.equal(ordenada.getCell('D2').value, 'Caneca Cerâmica')
  const corrigida = await renumerarPlanilhaPorItens(porItens)
  assert.equal(corrigida, porItens)
})

test('gera uma etiqueta de 40 por 25 mm por envio na ordem original', async () => {
  const original = await ler(await exemploFicticio({ comPacote: true }))
  const vendas = original.getWorksheet('Vendas BR')
  vendas.getCell('AT11').value = vendas.getCell('AT8').value
  vendas.getCell('AH8').value = 'Marília de Assunção'
  vendas.getCell('C14').value = 'Venda cancelada. Não envie.'
  vendas.getCell('AH15').value = 'Cliente do Pacote com Nome Comprido para Testar a Impressão'
  vendas.getCell('X16').value = 'Pote de vidro quadrado empilhável com tampa para organização e conservação'
  vendas.getCell('V16').value = 'POTE-00016-VERSAO-AZUL-MODELO-LONGO-640ML-2026-1234567890'
  vendas.getCell('X17').value = 'Kit 2 temporizadores mecânicos de tomada'
  vendas.getCell('X18').value = 'Timer eletrônico digital de cozinha'
  const etiquetas = await lerEtiquetas(await original.xlsx.writeBuffer())
  assert.deepEqual(etiquetas.map(({ envio }) => envio), [
    '70000000008', '70000000009', '70000000012', '70000000013', '70000000015', '70000000019',
  ])
  assert.equal(etiquetas[0].comprador, 'Marília de Assunção')
  assert.deepEqual(etiquetas[0].produtos.map(({ sku, unidades }) => [sku, unidades]), [
    ['SKU-08', 2], ['SKU-11', 1],
  ])
  assert.equal(etiquetas[4].comprador, vendas.getCell('AH15').value)
  assert.deepEqual(etiquetas[4].produtos.map(({ sku }) => sku), [vendas.getCell('V16').value, 'SKU-17', 'SKU-18'])
  const pdf = await PDFDocument.load(await gerarPdfEtiquetas(etiquetas))
  assert.equal(pdf.getPageCount(), etiquetas.length)
  assert.equal(pdf.catalog.getViewerPreferences().getPrintScaling(), 'None')
  for (const pagina of pdf.getPages()) {
    assert.ok(Math.abs(pagina.getWidth() * 25.4 / 72 - 40) < 0.001)
    assert.ok(Math.abs(pagina.getHeight() * 25.4 / 72 - 25) < 0.001)
  }
})

const caminhoExemplo = process.env.MIXBONIX_EXEMPLO
test('confere o relatório real localmente sem exibir os dados das células', { skip: !caminhoExemplo }, async () => {
  const original = await readFile(caminhoExemplo)
  const referencia = await ler(original)
  const vendas = referencia.getWorksheet('Vendas BR')
  const modelo = referencia.getWorksheet('Planilha1')
  const { bytes, porItens, linhas } = await formatarPlanilha(original)
  const saida = (await ler(bytes)).getWorksheet('Planilha1')
  assert.equal(linhas, 30)
  assert.equal(saida.rowCount, 31)
  assert.equal(saida.columnCount, 6)
  assert.deepEqual(saida.columns.map(({ width }) => width), modelo.columns.slice(0, 6).map(({ width }) => width))
  const itens = saida.getColumn('C').values.slice(2)
  assert.equal(itens.filter((item) => item === 'Kit 02 Timer').length, 19)
  assert.equal(itens.filter((item) => item === 'Timer Unitário').length, 11)
  assert.equal(saida.getColumn('B').values.slice(2).includes(2), true)
  for (let ordem = 1; ordem <= linhas; ordem++) {
    assert.equal(saida.getCell('A' + (ordem + 1)).value, ordem)
    for (const [destino, origem] of [['B', 'G'], ['D', 'X'], ['E', 'AC'], ['F', 'AE']]) {
      const esperado = vendas.getCell(origem + (ordem + 6)).value
      const recebido = saida.getCell(destino + (ordem + 1)).value
      assert.equal(recebido === esperado, true, 'Valor divergente na ordem ' + ordem + ', coluna ' + destino)
    }
  }
  const ordenada = (await ler(porItens)).getWorksheet('Planilha1')
  assert.deepEqual(ordenada.getColumn('C').values.slice(2), [
    ...Array(19).fill('Kit 02 Timer'), ...Array(11).fill('Timer Unitário'),
  ])
  const ordens = ordenada.getColumn('A').values.slice(2)
  assert.deepEqual(ordens, saida.getColumn('A').values.slice(2))
  const originaisPorItem = ['Kit 02 Timer', 'Timer Unitário'].flatMap((item) =>
    itens.map((nome, indice) => nome === item ? saida.getRow(indice + 2) : null).filter(Boolean),
  )
  for (const [indice, original] of originaisPorItem.entries()) {
    assert.equal(ordenada.getRow(indice + 2).values.slice(2).every((valor, coluna) =>
      valor === original.getCell(coluna + 2).value), true, 'Dados divergentes na linha ordenada ' + (indice + 2))
  }
  if (process.env.MIXBONIX_VALIDACAO_XLSX) {
    await writeFile(process.env.MIXBONIX_VALIDACAO_XLSX, new Uint8Array(bytes))
  }
})

const caminhoPacote = process.env.MIXBONIX_EXEMPLO_PACOTE
test('confere o pacote do relatório real nas duas versões sem exibir dados pessoais', { skip: !caminhoPacote }, async () => {
  const original = await readFile(caminhoPacote)
  const vendas = (await ler(original)).getWorksheet('Vendas BR')
  const { bytes, porItens, linhas } = await formatarPlanilha(original)
  assert.equal(linhas, 38)
  for (const [indice, arquivo] of [bytes, porItens].entries()) {
    const aba = (await ler(arquivo)).getWorksheet('Planilha1')
    assert.equal(aba.rowCount, 40)
    const titulos = aba.getColumn('D').values
    const resumo = titulos.indexOf('(3 pedidos feitos)')
    assert.equal(titulos.lastIndexOf('(3 pedidos feitos)'), resumo)
    assert.equal(indice ? resumo > 0 : resumo === 13, true)
    assert.equal(aba.getCell(resumo, 1).value, null)
    assert.equal(aba.getCell(resumo, 5).value === vendas.getCell('AC18').value, true)
    assert.equal(aba.getCell(resumo, 6).value === vendas.getCell('AE18').value, true)
    for (let i = 1; i <= 3; i++) {
      assert.equal(aba.getCell(resumo + i, 4).value === vendas.getCell('X' + (18 + i)).value, true)
      assert.equal(aba.getCell(resumo + i, 2).value, vendas.getCell('G' + (18 + i)).value)
    }
    assert.deepEqual(aba.getColumn('A').values.slice(2).filter(Number.isInteger), Array.from({ length: 38 }, (_, i) => i + 1))
  }
})
