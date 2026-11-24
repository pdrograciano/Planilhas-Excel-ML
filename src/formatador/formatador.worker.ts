import type { TarefaWorker } from './arquivo.ts'
import { ErroPlanilha, formatarPlanilha, renumerarPlanilhaPorItens } from './planilha.ts'

const worker = self as unknown as DedicatedWorkerGlobalScope

worker.onmessage = async ({ data }: MessageEvent<TarefaWorker>) => {
  try {
    if (data.tarefa === 'renumerar') {
      const bytes = await renumerarPlanilhaPorItens(data.bytes)
      worker.postMessage({ ok: true, bytes }, [bytes])
      return
    }
    const { bytes, porItens, linhas } = await formatarPlanilha(data.bytes)
    worker.postMessage({ ok: true, bytes, porItens, linhas }, [bytes, porItens])
  } catch (erro) {
    const mensagem = erro instanceof ErroPlanilha ? erro.message : 'Falha ao processar a planilha. Tente novamente.'
    worker.postMessage({ ok: false, mensagem })
  }
}
