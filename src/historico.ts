import { tipoExcel } from './formatador/arquivo.ts'

const prefixo = 'mixbonix:pedidos:v1:'
const marcaDaGuia = 'mixbonixPedidos'

type Entrada<T> = {
  id: string
  nome: string
  nomeGerado?: string
  data: number
  linhas: number
  original: T
  formatada: T
  porItens?: T
  porItensNumerado?: boolean
}

export type EntradaHistorico = Entrada<Blob> & { salva: boolean }

function blobParaBase64(arquivo: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const leitor = new FileReader()
    leitor.onerror = () => reject(leitor.error)
    leitor.onload = () => resolve(String(leitor.result).split(',')[1])
    leitor.readAsDataURL(arquivo)
  })
}

function base64ParaBlob(base64: string): Blob {
  const bytes = Uint8Array.from(atob(base64), (letra) => letra.charCodeAt(0))
  return new Blob([bytes], { type: tipoExcel })
}

export function identificadorHistorico(): string {
  return [...crypto.getRandomValues(new Uint32Array(4))].map((valor) => valor.toString(16).padStart(8, '0')).join('')
}

function chavesDoHistorico(armazenamento: Storage): string[] {
  return Object.keys(armazenamento).filter((chave) => chave.startsWith(prefixo))
}

export function carregarHistorico(): { entradas: EntradaHistorico[]; aviso: string } {
  const entradas: EntradaHistorico[] = []
  let aviso = ''
  try {
    const armazenamento = window.sessionStorage
    const estado = window.history.state
    if (window.opener && !estado?.[marcaDaGuia]) {
      // Guias abertas por script podem herdar uma cópia da sessão da guia de origem.
      chavesDoHistorico(armazenamento).forEach((chave) => armazenamento.removeItem(chave))
    }
    window.history.replaceState({ ...estado, [marcaDaGuia]: true }, '')
    for (const chave of chavesDoHistorico(armazenamento)) {
      try {
        const entrada: Entrada<string> = JSON.parse(armazenamento.getItem(chave) ?? '')
        entradas.push({
          ...entrada,
          original: base64ParaBlob(entrada.original),
          formatada: base64ParaBlob(entrada.formatada),
          porItens: entrada.porItens ? base64ParaBlob(entrada.porItens) : undefined,
          salva: true,
        })
      } catch {
        aviso = 'Uma entrada do histórico não pôde ser recuperada.'
      }
    }
  } catch {
    aviso = 'O armazenamento da guia está indisponível. Os arquivos gerados ficarão disponíveis até recarregar a página.'
  }
  return { entradas: entradas.sort((a, b) => b.data - a.data), aviso }
}

export async function guardarNoHistorico(entrada: Entrada<Blob>): Promise<boolean> {
  try {
    const [original, formatada, porItens] = await Promise.all([
      blobParaBase64(entrada.original),
      blobParaBase64(entrada.formatada),
      entrada.porItens && blobParaBase64(entrada.porItens),
    ])
    const dados: Entrada<string> = { ...entrada, original, formatada, porItens }
    window.sessionStorage.setItem(prefixo + entrada.id, JSON.stringify(dados))
    return true
  } catch {
    return false
  }
}

export async function guardarNumeracaoPorItens(id: string, arquivo: Blob): Promise<void> {
  try {
    const porItens = await blobParaBase64(arquivo)
    const armazenamento = window.sessionStorage
    const chave = prefixo + id
    const registro = armazenamento.getItem(chave)
    if (registro === null) return
    const dados: Entrada<string> = JSON.parse(registro)
    armazenamento.setItem(chave, JSON.stringify({ ...dados, porItens, porItensNumerado: true }))
  } catch {
    // O arquivo corrigido continua disponível na página, mesmo se o histórico não puder ser atualizado.
  }
}

export function excluirDoHistorico(id: string): boolean {
  try {
    window.sessionStorage.removeItem(prefixo + id)
    return true
  } catch {
    return false
  }
}

export function limparHistorico(): boolean {
  try {
    const armazenamento = window.sessionStorage
    chavesDoHistorico(armazenamento).forEach((chave) => armazenamento.removeItem(chave))
    return true
  } catch {
    return false
  }
}
