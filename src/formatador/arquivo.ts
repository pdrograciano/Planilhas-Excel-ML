import type { EtiquetaEnvio } from './etiquetas.ts'

export const tipoExcel = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

export type TarefaWorker = { tarefa: 'formatar' | 'renumerar' | 'ler-envios'; bytes: ArrayBuffer }

type ResultadoWorker =
  | { ok: true; bytes?: ArrayBuffer; porItens?: ArrayBuffer; linhas?: number; envios?: EtiquetaEnvio[] }
  | { ok: false; mensagem: string }

export function nomeBaseGerado(data: number): string {
  const horarioLocal = new Date(data)
  const doisDigitos = (valor: number) => String(valor).padStart(2, '0')
  const dia = [horarioLocal.getDate(), horarioLocal.getMonth() + 1].map(doisDigitos).join('-')
  const hora = [horarioLocal.getHours(), horarioLocal.getMinutes(), horarioLocal.getSeconds()].map(doisDigitos).join('-')
  return dia + '-' + horarioLocal.getFullYear() + '_' + hora + '_MercadoLivre_mixbonixgerator.xlsx'
}

export function nomeDaVersao(nomeBase: string, versao: 'Formatado' | 'PorItens'): string {
  return nomeBase.replace('_MercadoLivre_', '_' + versao + '_MercadoLivre_')
}

async function executarWorker(arquivo: Blob, tarefa: TarefaWorker['tarefa']) {
  let bytes: ArrayBuffer
  try {
    bytes = await arquivo.arrayBuffer()
  } catch {
    throw new Error('Não foi possível ler o arquivo selecionado. Selecione o arquivo novamente.')
  }
  const worker = new Worker(new URL('./formatador.worker.ts', import.meta.url), { type: 'module' })
  try {
    return await new Promise<Extract<ResultadoWorker, { ok: true }>>((resolve, reject) => {
      worker.onmessage = ({ data }: MessageEvent<ResultadoWorker>) => {
        if (!data.ok) {
          reject(new Error(data.mensagem))
          return
        }
        resolve(data)
      }
      worker.onerror = (evento) => {
        evento.preventDefault()
        reject(new Error('Não foi possível iniciar o processamento do Excel. Recarregue a página e tente novamente.'))
      }
      const mensagem: TarefaWorker = { tarefa, bytes }
      worker.postMessage(mensagem, [bytes])
    })
  } finally {
    worker.terminate()
  }
}

export async function processarArquivo(arquivo: File) {
  const { bytes, porItens, linhas } = await executarWorker(arquivo, 'formatar')
  if (!bytes || !porItens || linhas === undefined) throw new Error('Não foi possível gerar as planilhas formatadas.')
  return {
    formatada: new Blob([bytes], { type: tipoExcel }),
    porItens: new Blob([porItens], { type: tipoExcel }),
    linhas,
  }
}

export async function renumerarArquivoPorItens(arquivo: Blob): Promise<Blob> {
  const { bytes } = await executarWorker(arquivo, 'renumerar')
  if (!bytes) throw new Error('Não foi possível renumerar a planilha.')
  return new Blob([bytes], { type: tipoExcel })
}

export async function lerEnviosDaPlanilha(arquivo: Blob): Promise<EtiquetaEnvio[]> {
  const { envios } = await executarWorker(arquivo, 'ler-envios')
  if (!envios) throw new Error('Não foi possível ler os envios da planilha.')
  return envios
}

export function baixarArquivo(arquivo: Blob, nome: string): void {
  const url = URL.createObjectURL(arquivo)
  const link = document.createElement('a')
  link.href = url
  link.download = nome
  document.body.append(link)
  link.click()
  link.remove()
  window.setTimeout(() => URL.revokeObjectURL(url), 1000)
}
