import jsQR from 'jsqr'
import { getDocument, GlobalWorkerOptions } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFPageProxy } from 'pdfjs-dist'
import workerUrl from 'pdfjs-dist/legacy/build/pdf.worker.min.mjs?url'
import { chaveQr } from './pedidos.ts'
import type { LoteConferencia, PedidoConferencia } from './pedidos.ts'

GlobalWorkerOptions.workerSrc = workerUrl

async function linhasDaPagina(pagina: PDFPageProxy) {
  const { items } = await pagina.getTextContent()
  const linhas: { y: number; partes: { x: number; texto: string }[] }[] = []
  for (const item of items) {
    if (!('str' in item) || !item.str.trim()) continue
    const [, , , , x, y] = item.transform
    let linha = linhas.find((candidata) => Math.abs(candidata.y - y) < 2)
    if (!linha) {
      linha = { y, partes: [] }
      linhas.push(linha)
    }
    linha.partes.push({ x, texto: item.str })
  }
  return linhas.sort((a, b) => b.y - a.y).map(({ partes }) =>
    partes.sort((a, b) => a.x - b.x).map(({ texto }) => texto).join(' ').replace(/\s+/g, ' ').trim())
}

function nomeNaEtiqueta(linhas: string[]): string {
  const explicito = linhas.find((linha) => /^destinat[áa]rio\s*:/i.test(linha))?.split(':').slice(1).join(':').trim()
  if (explicito) return explicito
  const endereco = linhas.findIndex((linha) => /^endere[cç]o\s*:/i.test(linha))
  const anterior = linhas[endereco - 1] ?? ''
  const candidato = anterior.startsWith('(') ? linhas[endereco - 2] ?? '' : anterior.split('(')[0].trim()
  return /\p{L}/u.test(candidato) && !/[:\d]/.test(candidato) ? candidato : ''
}

async function lerQrDaPagina(pagina: PDFPageProxy, contexto: CanvasRenderingContext2D, tamanho: number) {
  const original = pagina.getViewport({ scale: 1 })
  const viewport = pagina.getViewport({ scale: tamanho / Math.max(original.width, original.height) })
  const canvas = contexto.canvas
  canvas.width = Math.ceil(viewport.width)
  canvas.height = Math.ceil(viewport.height)
  await pagina.render({ canvas, canvasContext: contexto, viewport, background: '#fff' }).promise
  // As etiquetas do Mercado Livre posicionam o QR no canto inferior direito.
  const x = Math.floor(canvas.width * .54)
  const y = Math.floor(canvas.height * .69)
  const recorte = contexto.getImageData(x, y, canvas.width - x, canvas.height - y)
  const qrNoCanto = jsQR(recorte.data, recorte.width, recorte.height, { inversionAttempts: 'attemptBoth' })?.data
  if (qrNoCanto) return qrNoCanto
  const imagem = contexto.getImageData(0, 0, canvas.width, canvas.height)
  return jsQR(imagem.data, imagem.width, imagem.height, { inversionAttempts: 'attemptBoth' })?.data ?? ''
}

export async function lerPdfEtiquetas(arquivo: File, progredir: (pagina: number, total: number) => void): Promise<LoteConferencia> {
  const canvas = document.createElement('canvas')
  const contexto = canvas.getContext('2d', { willReadFrequently: true })
  if (!contexto) throw new Error('O navegador não conseguiu preparar a leitura dos QR codes.')
  const tarefa = getDocument({ data: await arquivo.arrayBuffer() })
  const pedidos = new Map<string, PedidoConferencia>()
  const paginasNaoReconhecidas: number[] = []
  let etiquetaAnterior: { pedido: PedidoConferencia; nota: string } | undefined

  try {
    const pdf = await tarefa.promise
    for (let numero = 1; numero <= pdf.numPages; numero++) {
      progredir(numero, pdf.numPages)
      const pagina = await pdf.getPage(numero)
      try {
        const linhas = await linhasDaPagina(pagina)
        const texto = linhas.join('\n')
        const envioImpresso = texto.match(/\benvio\s*:\s*(\d(?:[ \t]*\d){5,19})\b/i)?.[1].replace(/\s/g, '') ?? ''
        if (!envioImpresso && /\bdanfe\b/i.test(texto)) {
          const destinatario = nomeNaEtiqueta(linhas)
          const chaveFiscal = texto.match(/\b\d{44}\b/)?.[0]
          const nf = chaveFiscal?.slice(25, 34).replace(/^0+/, '')
          if (destinatario && etiquetaAnterior && !etiquetaAnterior.pedido.destinatario &&
            (etiquetaAnterior.nota ? nf === etiquetaAnterior.nota : etiquetaAnterior.pedido.pagina + 1 === numero)) {
            etiquetaAnterior.pedido.destinatario = destinatario
          }
          etiquetaAnterior = undefined
          continue
        }

        const qr = await lerQrDaPagina(pagina, contexto, 800) || await lerQrDaPagina(pagina, contexto, 1450)
        const codigo = chaveQr(qr)
        const envioQr = codigo.match(/^envio:(\d+)$/)?.[1] ?? ''
        const envio = envioImpresso || envioQr
        if (!envio) {
          paginasNaoReconhecidas.push(numero)
          etiquetaAnterior = undefined
          continue
        }
        const qrConfirmado = codigo && (envioQr ? envioQr === envio : codigo === envio) ? codigo : null
        const destinatario = nomeNaEtiqueta(linhas)
        const pedido = pedidos.get(envio) ?? { pagina: numero, envio, destinatario, codigo: qrConfirmado }
        pedido.codigo ??= qrConfirmado
        pedido.destinatario ||= destinatario
        pedidos.set(envio, pedido)
        const nota = texto.match(/\bNF\s*:\s*([\d.]+)/i)?.[1].replace(/\D/g, '').replace(/^0+/, '') ?? ''
        etiquetaAnterior = { pedido, nota }
      } finally {
        pagina.cleanup()
      }
    }
    if (!pedidos.size) throw new Error('Nenhum envio foi encontrado. Selecione o PDF de etiquetas do Mercado Livre.')
    return { pedidos: [...pedidos.values()], paginasNaoReconhecidas }
  } finally {
    await tarefa.destroy()
    canvas.width = canvas.height = 0
  }
}
