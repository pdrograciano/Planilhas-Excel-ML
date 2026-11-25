import { PDFDocument, PrintScaling, StandardFonts } from 'pdf-lib'
import type { PDFFont } from 'pdf-lib'

interface ProdutoEtiqueta {
  nome: string
  unidades: number | null
  sku: string
}

export interface EtiquetaEnvio {
  comprador: string
  envio: string
  produtos: ProdutoEtiqueta[]
}

interface Trecho {
  linhas: string[]
  fonte: PDFFont
  tamanho: number
  entrelinha: number
  intervalo: number
  divisoria: boolean
}

const largura = 100 * 72 / 25.4
const altura = 150 * 72 / 25.4
const margem = 17
const larguraTexto = largura - 2 * margem
const quantidades = new Intl.NumberFormat('pt-BR')

function textoImprimivel(texto: string, caracteres: Set<number>): string {
  return [...texto.normalize('NFC').replace(/\s+/g, ' ').trim()].map((letra) => {
    if (caracteres.has(letra.codePointAt(0)!)) return letra
    const equivalente = letra.normalize('NFKD').replace(/\p{M}/gu, '')
    return equivalente && [...equivalente].every((parte) => caracteres.has(parte.codePointAt(0)!)) ? equivalente : '?'
  }).join('')
}

function quebrarLinhas(texto: string, fonte: PDFFont, tamanho: number): string[] {
  const linhas: string[] = []
  let linha = ''
  const cabe = (valor: string) => fonte.widthOfTextAtSize(valor, tamanho) <= larguraTexto
  for (const palavra of texto.split(' ')) {
    const adicionada = linha ? linha + ' ' + palavra : palavra
    if (cabe(adicionada)) {
      linha = adicionada
      continue
    }
    if (linha) linhas.push(linha)
    linha = ''
    if (cabe(palavra)) {
      linha = palavra
      continue
    }
    for (const letra of palavra) {
      if (linha && !cabe(linha + letra)) {
        linhas.push(linha)
        linha = ''
      }
      linha += letra
    }
  }
  if (linha) linhas.push(linha)
  return linhas
}

function comporEtiqueta(etiqueta: EtiquetaEnvio, normal: PDFFont, negrito: PDFFont, caracteres: Set<number>, escala: number) {
  const trecho = (texto: string, fonte: PDFFont, pontos: number, intervalo = 0, divisoria = false): Trecho => {
    const tamanho = Math.max(7, pontos * escala)
    return {
      linhas: quebrarLinhas(textoImprimivel(texto, caracteres), fonte, tamanho),
      fonte, tamanho, entrelinha: tamanho * 1.23, intervalo: intervalo * escala, divisoria,
    }
  }
  const trechos = [
    trecho('COMPRADOR', negrito, 9, 3),
    trecho(etiqueta.comprador || 'Não informado', negrito, 19, 15, true),
    trecho(etiqueta.produtos.length === 1 ? 'PRODUTO' : 'PRODUTOS', negrito, 9, 5),
  ]
  for (const [indice, produto] of etiqueta.produtos.entries()) {
    const quantidade = produto.unidades === null ? 'não informada' : quantidades.format(produto.unidades)
    const separar = indice < etiqueta.produtos.length - 1
    trechos.push(
      trecho(produto.nome || 'Produto não informado', negrito, 15, 5),
      trecho('Quantidade: ' + quantidade, negrito, 12, 3),
      trecho('SKU: ' + (produto.sku || 'não informado'), normal, 12, separar ? 13 : 9, separar),
    )
  }
  if (etiqueta.envio) trechos.push(trecho('Envio: ' + etiqueta.envio, normal, 9))
  return trechos
}

export async function gerarPdfEtiquetas(etiquetas: EtiquetaEnvio[]): Promise<ArrayBuffer> {
  const pdf = await PDFDocument.create()
  pdf.setTitle('PDF nomes e SKU - MIXBONIX')
  pdf.catalog.getOrCreateViewerPreferences().setPrintScaling(PrintScaling.None)
  const normal = await pdf.embedFont(StandardFonts.Helvetica)
  const negrito = await pdf.embedFont(StandardFonts.HelveticaBold)
  const caracteres = new Set(normal.getCharacterSet())

  for (const [indice, etiqueta] of etiquetas.entries()) {
    let layout: Trecho[] | undefined
    for (const escala of [1, .9, .8, .7, .6, .5]) {
      const trechos = comporEtiqueta(etiqueta, normal, negrito, caracteres, escala)
      const espaco = trechos.reduce((total, trecho) => total + trecho.linhas.length * trecho.entrelinha + trecho.intervalo, 0)
      if (espaco <= altura - 2 * margem) {
        layout = trechos
        break
      }
    }
    if (!layout) throw new Error('O pedido ' + (etiqueta.envio || indice + 1) + ' tem itens demais para caber em uma etiqueta de 10 × 15 cm.')

    // O conteúdo é o mesmo no verso da etiqueta de envio e no verso da DANFE.
    for (let via = 0; via < 2; via++) {
      const pagina = pdf.addPage([largura, altura])
      let y = altura - margem
      for (const trecho of layout) {
        for (const linha of trecho.linhas) {
          pagina.drawText(linha, { x: margem, y: y - trecho.tamanho, size: trecho.tamanho, font: trecho.fonte })
          y -= trecho.entrelinha
        }
        if (trecho.divisoria) {
          const meio = y - trecho.intervalo / 2
          pagina.drawLine({ start: { x: margem, y: meio }, end: { x: largura - margem, y: meio }, thickness: .5 })
        }
        y -= trecho.intervalo
      }
    }
  }
  return new Uint8Array(await pdf.save()).buffer
}
