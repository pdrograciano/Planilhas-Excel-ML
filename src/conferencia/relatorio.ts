import { PageSizes, PDFDocument, rgb, StandardFonts } from 'pdf-lib'
import { quebrarLinhas, textoImprimivel } from '../pdf/texto.ts'
import type { ProdutoConferencia } from './pedidos.ts'

export interface DadosRelatorioConferencia {
  data: Date
  desconhecidos: number
  arquivos: { nome: string; paginasNaoReconhecidas: number[] }[]
  pedidos: { envio: string; destinatario: string; produtos: ProdutoConferencia[]; identificado: boolean }[]
}

const [largura, altura] = PageSizes.A4
const margem = 40
const larguraUtil = largura - margem * 2
const limite = 63
const tinta = rgb(.09, .12, .17)
const secundario = rgb(.30, .33, .38)
const borda = rgb(.80, .82, .84)
const fundo = rgb(.94, .95, .96)
const dataLocal = new Intl.DateTimeFormat('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric' })
const horaLocal = new Intl.DateTimeFormat('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' })
const numeros = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 20 })

export async function gerarRelatorioConferencia(dados: DadosRelatorioConferencia) {
  const pdf = await PDFDocument.create()
  pdf.setTitle('Relatório de conferência - MIXBONIX ORDER')
  pdf.setCreationDate(dados.data)
  const normal = await pdf.embedFont(StandardFonts.Helvetica)
  const negrito = await pdf.embedFont(StandardFonts.HelveticaBold)
  const caracteres = new Set(normal.getCharacterSet())
  const dia = dataLocal.format(dados.data)
  const hora = horaLocal.format(dados.data)
  const colunas = [
    { titulo: 'Envio', x: margem, largura: 106, fonte: normal },
    { titulo: 'Destinatário', x: margem + 106, largura: 140, fonte: normal },
    { titulo: 'Produto e quantidade', x: margem + 246, largura: 172, fonte: normal },
    { titulo: 'Situação', x: margem + 418, largura: larguraUtil - 418, fonte: negrito },
  ]

  function novaPagina() {
    const pagina = pdf.addPage(PageSizes.A4)
    pagina.drawText('MIXBONIX ORDER', { x: margem, y: altura - 47, font: negrito, size: 10, color: secundario })
    pagina.drawText('Relatório de conferência', { x: margem, y: altura - 75, font: negrito, size: 23, color: tinta })
    pagina.drawText(`Finalizada em ${dia} às ${hora}`, { x: margem, y: altura - 94, font: normal, size: 10, color: secundario })
    return pagina
  }

  let pagina = novaPagina()
  let y = altura - 114
  const identificados = dados.pedidos.filter(({ identificado }) => identificado).length
  const totais = [
    ['Envios no lote', dados.pedidos.length],
    ['Identificados', identificados],
    ['Pendentes', dados.pedidos.length - identificados],
    ['Leituras sem identificação', dados.desconhecidos],
  ] as const
  const larguraTotal = larguraUtil / totais.length
  pagina.drawRectangle({ x: margem, y: y - 52, width: larguraUtil, height: 52, color: fundo, borderColor: borda, borderWidth: .6 })
  for (const [indice, [titulo, valor]] of totais.entries()) {
    const x = margem + larguraTotal * indice
    if (indice) pagina.drawLine({ start: { x, y }, end: { x, y: y - 52 }, color: borda, thickness: .6 })
    pagina.drawText(titulo, { x: x + 10, y: y - 16, font: normal, size: 8.5, color: secundario })
    pagina.drawText(numeros.format(valor), { x: x + 10, y: y - 39, font: negrito, size: 18, color: tinta })
  }
  y -= 73

  function cabecalhoTabela() {
    y -= 24
    pagina.drawRectangle({ x: margem, y, width: larguraUtil, height: 24, color: fundo })
    for (const coluna of colunas) {
      pagina.drawText(coluna.titulo, { x: coluna.x + 7, y: y + 8, font: negrito, size: 9, color: tinta })
    }
  }

  cabecalhoTabela()
  const entrelinha = 13
  const espacoLinha = 12
  const capacidadePagina = Math.floor((altura - 114 - 24 - limite - espacoLinha) / entrelinha)
  for (const pedido of dados.pedidos) {
    const produtos = pedido.produtos.length ? pedido.produtos.map(({ nome, quantidade }) =>
      `${nome} - Qtd. ${quantidade === null ? 'não informada' : numeros.format(quantidade)}`) : ['Produto e quantidade indisponíveis']
    const valores = [[pedido.envio], [pedido.destinatario || 'Destinatário não informado'], produtos, [pedido.identificado ? 'Identificado' : 'Pendente']]
    const linhas = colunas.map(({ fonte, largura }, indice) => valores[indice].flatMap((texto) =>
      quebrarLinhas(textoImprimivel(texto, caracteres), fonte, 10, largura - 14)))
    const totalLinhas = Math.max(...linhas.map((celula) => celula.length))
    if (totalLinhas <= capacidadePagina && totalLinhas * entrelinha + espacoLinha > y - limite) {
      pagina = novaPagina()
      y = altura - 114
      cabecalhoTabela()
    }
    let inicio = 0
    while (inicio < totalLinhas) {
      if (y - limite < entrelinha + espacoLinha) {
        pagina = novaPagina()
        y = altura - 114
        cabecalhoTabela()
      }
      const fim = Math.min(totalLinhas, inicio + Math.floor((y - limite - espacoLinha) / entrelinha))
      for (const [indice, coluna] of colunas.entries()) {
        const trechos = inicio > 0 && indice === 0 ? [...linhas[0], '(continuação)'] : linhas[indice].slice(inicio, fim)
        for (const [linha, trecho] of trechos.slice(0, fim - inicio).entries()) {
          pagina.drawText(trecho, { x: coluna.x + 7, y: y - 15 - linha * entrelinha, font: coluna.fonte, size: 10, color: tinta })
        }
      }
      y -= (fim - inicio) * entrelinha + espacoLinha
      pagina.drawLine({ start: { x: margem, y }, end: { x: largura - margem, y }, color: borda, thickness: .5 })
      inicio = fim
    }
  }

  if (dados.arquivos.length) {
    if (y - limite < 54) { pagina = novaPagina(); y = altura - 94 }
    y -= 25
    pagina.drawText('PDFs do lote', { x: margem, y, font: negrito, size: 11, color: tinta })
    y -= 16
    for (const { nome, paginasNaoReconhecidas } of dados.arquivos) {
      const detalhe = paginasNaoReconhecidas.length ? ' - páginas não reconhecidas: ' + paginasNaoReconhecidas.join(', ') : ''
      for (const linha of quebrarLinhas(textoImprimivel(nome + detalhe, caracteres), normal, 9, larguraUtil)) {
        if (y < limite) { pagina = novaPagina(); y = altura - 114 }
        pagina.drawText(linha, { x: margem, y, font: normal, size: 9, color: secundario })
        y -= 13
      }
    }
  }

  const paginas = pdf.getPages()
  for (const [indice, folha] of paginas.entries()) {
    const texto = `Página ${indice + 1} de ${paginas.length}`
    folha.drawLine({ start: { x: margem, y: 45 }, end: { x: largura - margem, y: 45 }, color: borda, thickness: .6 })
    folha.drawText(texto, { x: largura - margem - normal.widthOfTextAtSize(texto, 9), y: 31, font: normal, size: 9, color: secundario })
  }
  const bytes = new Uint8Array(await pdf.save())
  return {
    arquivo: new Blob([bytes], { type: 'application/pdf' }),
    nome: `Conferencia_MIXBONIX_${dia.replace(/\//g, '-')}_${hora.replace(/:/g, '-')}.pdf`,
  }
}
