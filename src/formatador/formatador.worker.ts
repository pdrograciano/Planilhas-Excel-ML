import type { TarefaWorker } from './arquivo.ts'
import { ErroPlanilha, formatarPlanilha, lerEtiquetas, renumerarPlanilhaPorItens } from './planilha.ts'

const worker = self as unknown as DedicatedWorkerGlobalScope

worker.onmessage = async ({ data }: MessageEvent<TarefaWorker>) => {
  try {
    if (data.tarefa === 'ler-envios') {
      worker.postMessage({ ok: true, envios: await lerEtiquetas(data.bytes) })
      return
    }
    if (data.tarefa === 'etiquetas') {
      const etiquetas = await lerEtiquetas(data.bytes)
      const { gerarPdfEtiquetas } = await import('./etiquetas.ts')
      const bytes = await gerarPdfEtiquetas(etiquetas)
      worker.postMessage({ ok: true, bytes }, [bytes])
      return
    }
    if (data.tarefa === 'renumerar') {
      const bytes = await renumerarPlanilhaPorItens(data.bytes)
      worker.postMessage({ ok: true, bytes }, [bytes])
      return
    }
    const { bytes, porItens, linhas } = await formatarPlanilha(data.bytes)
    worker.postMessage({ ok: true, bytes, porItens, linhas }, [bytes, porItens])
  } catch (erro) {
    const mensagem = erro instanceof ErroPlanilha || (data.tarefa === 'etiquetas' && erro instanceof Error)
      ? erro.message : 'Falha ao processar a planilha. Tente novamente.'
    worker.postMessage({ ok: false, mensagem })
  }
}
