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

export function reunirPedidos(lotes: LoteConferencia[]): PedidoConferencia[] {
  const pedidos = new Map<string, PedidoConferencia>()
  for (const lote of lotes) {
    for (const pedido of lote.pedidos) {
      const anterior = pedidos.get(pedido.envio)
      pedidos.set(pedido.envio, anterior ? {
        ...anterior,
        destinatario: anterior.destinatario || pedido.destinatario,
        codigo: anterior.codigo ?? pedido.codigo,
      } : pedido)
    }
  }
  return [...pedidos.values()]
}

export function produtosPorEnvio(envios: EtiquetaEnvio[]): Map<string, ProdutoConferencia[]> {
  const encontrados = new Map<string, ProdutoConferencia[]>()
  for (const { envio, produtos } of envios) {
    const itens = produtos.filter(({ nome }) => nome.trim()).map(({ nome, unidades }) => ({
      nome: nomeCurto(nome), quantidade: unidades,
    }))
    if (envio && itens.length) encontrados.set(envio, itens)
  }
  return encontrados
}
