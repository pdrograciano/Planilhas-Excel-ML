import { nomeCurto } from '../formatador/item.ts'
import type { EtiquetaEnvio } from '../formatador/etiquetas.ts'

export interface PedidoConferencia {
  pagina: number
  envio: string
  destinatario: string
  codigo: string | null
}

export interface LoteConferencia {
  pedidos: PedidoConferencia[]
  paginasNaoReconhecidas: number[]
}

export interface ProdutoConferencia {
  nome: string
  quantidade: number | null
}

export function chaveQr(texto: string): string {
  const codigo = texto.trim()
  try {
    const dados = JSON.parse(codigo)
    if (dados?.t === 'lm' && /^\d{6,20}$/.test(String(dados.id))) return 'envio:' + dados.id
  } catch {
    // Um QR também pode conter texto simples.
  }
  return codigo
}

export function combinarProdutos(pedidos: PedidoConferencia[], envios: EtiquetaEnvio[]): Map<string, ProdutoConferencia[]> {
  const ids = new Set(pedidos.map(({ envio }) => envio).filter(Boolean))
  const encontrados = new Map<string, ProdutoConferencia[]>()
  for (const { envio, produtos } of envios) {
    if (!ids.has(envio)) continue
    encontrados.set(envio, produtos.filter(({ nome }) => nome.trim()).map(({ nome, unidades }) => ({
      nome: nomeCurto(nome), quantidade: unidades,
    })))
  }
  return encontrados
}
