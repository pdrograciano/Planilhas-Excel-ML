import { PDFDocument, PrintScaling, StandardFonts } from 'pdf-lib'
import type { PDFFont } from 'pdf-lib'
import { quebrarLinhas, textoImprimivel } from '../pdf/texto.ts'

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

const largura = 40 * 72 / 25.4
const altura = 25 * 72 / 25.4
const margem = 3.5
const larguraTexto = largura - 2 * margem
const quantidades = new Intl.NumberFormat('pt-BR')

function comporEtiqueta(etiqueta: EtiquetaEnvio, normal: PDFFont, negrito: PDFFont, caracteres: Set<number>, escala: number) {
  const trecho = (texto: string, fonte: PDFFont, pontos: number, intervalo = 0, divisoria = false): Trecho => {
    const tamanho = Math.max(4.5, pontos * escala)
    return {
      linhas: quebrarLinhas(textoImprimivel(texto, caracteres), fonte, tamanho, larguraTexto),
      fonte, tamanho, entrelinha: tamanho * (1.05 + .07 * escala), intervalo: intervalo * escala, divisoria,
    }
  }
  const trechos = [trecho(etiqueta.comprador || 'Comprador não informado', negrito, 8.5, 3, true)]
  for (const produto of etiqueta.produtos) {
    const quantidade = produto.unidades === null ? 's/ info' : quantidades.format(produto.unidades)
    trechos.push(
      trecho(produto.nome || 'Produto não informado', negrito, 7.6),
      trecho('Qtd: ' + quantidade + ' | SKU: ' + (produto.sku || 's/ info'), normal, 6.8, 2),
    )
  }
  if (etiqueta.envio) trechos.push(trecho('Envio: ' + etiqueta.envio, normal, 6.2))
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
    for (const escala of [1, .95, .9, .85, .8, .75, .7, .65, .6, .55]) {
      const trechos = comporEtiqueta(etiqueta, normal, negrito, caracteres, escala)
      const espaco = trechos.reduce((total, trecho) => total + trecho.linhas.length * trecho.entrelinha + trecho.intervalo, 0)
      if (espaco <= altura - 2 * margem) {
        layout = trechos
        break
      }
    }
    if (!layout) throw new Error('Os dados do envio ' + (etiqueta.envio || indice + 1) + ' não cabem em uma etiqueta de 4 × 2,5 cm.')

    const pagina = pdf.addPage([largura, altura])
    let y = altura - margem
    for (const trecho of layout) {
      for (const linha of trecho.linhas) {
        pagina.drawText(linha, { x: margem, y: y - trecho.tamanho, size: trecho.tamanho, font: trecho.fonte })
        y -= trecho.entrelinha
      }
      if (trecho.divisoria) {
        const meio = y - trecho.intervalo / 2
        pagina.drawLine({ start: { x: margem, y: meio }, end: { x: largura - margem, y: meio }, thickness: .3 })
      }
      y -= trecho.intervalo
    }
  }
  return new Uint8Array(await pdf.save()).buffer
}
