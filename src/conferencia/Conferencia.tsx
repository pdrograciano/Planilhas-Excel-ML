import { Camera, CheckCircle2, CircleAlert, ClipboardCheck, FileText, FolderOpen, LoaderCircle, ScanLine, UploadCloud } from 'lucide-react'
import { useEffect, useMemo, useRef, useState } from 'react'
import type { DragEvent } from 'react'
import { lerEnviosDaPlanilha, tipoExcel } from '../formatador/arquivo.ts'
import type { EntradaHistorico } from '../historico.ts'
import { CameraConferencia } from './CameraConferencia.tsx'
import { chaveQr, combinarProdutos } from './pedidos.ts'
import type { LoteConferencia, PedidoConferencia, ProdutoConferencia } from './pedidos.ts'

interface Bipagem {
  tipo: 'identificado' | 'repetido' | 'desconhecido'
  pedido?: PedidoConferencia
}

function DadosPedido({ pedido, produtos }: { pedido: PedidoConferencia; produtos?: ProdutoConferencia[] }) {
  return (
    <div className="dados-pedido">
      <strong>{pedido.destinatario || 'Destinatário não informado'}</strong>
      <span>Envio {pedido.envio}</span>
      {produtos?.length ? (
        <ul>
          {produtos.map(({ nome, quantidade }, indice) => (
            <li key={indice}><span>{nome}</span><span>{quantidade === null ? 'Quantidade não informada' : `Qtd. ${quantidade}`}</span></li>
          ))}
        </ul>
      ) : <span>Produto e quantidade indisponíveis. Carregue a planilha de vendas correspondente.</span>}
    </div>
  )
}

export default function Conferencia({ ativo, entradas }: { ativo: boolean; entradas: EntradaHistorico[] }) {
  const inputPdf = useRef<HTMLInputElement>(null)
  const inputExcel = useRef<HTMLInputElement>(null)
  const areaCamera = useRef<HTMLElement>(null)
  const areaResumo = useRef<HTMLHeadingElement>(null)
  const profundidade = useRef(0)
  const validos = useRef(new Set<string>())
  const [arrastando, setArrastando] = useState(false)
  const [arquivo, setArquivo] = useState('')
  const [lote, setLote] = useState<LoteConferencia | null>(null)
  const [produtos, setProdutos] = useState(new Map<string, ProdutoConferencia[]>())
  const [processando, setProcessando] = useState('')
  const [erro, setErro] = useState('')
  const [camera, setCamera] = useState(false)
  const [bipagem, setBipagem] = useState<Bipagem | null>(null)
  const [quantidade, setQuantidade] = useState(0)
  const [desconhecidos, setDesconhecidos] = useState(0)
  const [finalizado, setFinalizado] = useState(false)
  const codigos = useMemo(() => new Map(lote?.pedidos.flatMap((pedido) => pedido.codigo ? [[pedido.codigo, pedido] as const] : [])), [lote])
  const faltamProdutos = lote?.pedidos.filter(({ envio }) => {
    const itens = produtos.get(envio)
    return !itens?.length || itens.some(({ quantidade }) => quantidade === null)
  }).length ?? 0
  const semQr = (lote?.pedidos.length ?? 0) - codigos.size

  useEffect(() => { if (!ativo) setCamera(false) }, [ativo])
  useEffect(() => { if (camera) areaCamera.current?.scrollIntoView({ block: 'nearest' }) }, [camera])
  useEffect(() => { if (finalizado) areaResumo.current?.focus() }, [finalizado])

  async function importarPdf(arquivos: FileList | null) {
    if (!arquivos?.length || processando) return
    if (arquivos.length !== 1 || !/\.pdf$/i.test(arquivos[0].name)) {
      setErro('Selecione um único arquivo no formato .pdf.')
      return
    }
    const arquivoPdf = arquivos[0]
    setErro('')
    setCamera(false)
    setProcessando('Abrindo PDF…')
    try {
      const { lerPdfEtiquetas } = await import('./leituraPdf.ts')
      const novoLote = await lerPdfEtiquetas(arquivoPdf, (atual, total) => setProcessando(`Lendo página ${atual} de ${total}…`))
      const encontrados = new Map<string, ProdutoConferencia[]>()
      if (entradas.length) setProcessando('Buscando produtos nas planilhas do histórico…')
      for (const entrada of entradas) {
        try {
          const dados = combinarProdutos(novoLote.pedidos, await lerEnviosDaPlanilha(entrada.original))
          for (const [envio, lista] of dados) {
            if (lista.length && !encontrados.has(envio)) encontrados.set(envio, lista)
          }
          if (encontrados.size === novoLote.pedidos.length) break
        } catch {
          // Uma planilha antiga pode estar ilegível. Ainda é possível carregar outra.
        }
      }
      setArquivo(arquivoPdf.name)
      setLote(novoLote)
      setProdutos(encontrados)
      validos.current.clear()
      setQuantidade(0)
      setBipagem(null)
      setDesconhecidos(0)
      setFinalizado(false)
    } catch (falha) {
      setErro(falha instanceof Error && falha.message.startsWith('Nenhum envio')
        ? falha.message : 'Não foi possível ler o PDF. Exporte as etiquetas novamente e tente de novo.')
    } finally {
      setProcessando('')
    }
  }

  async function importarExcel(planilha: File | undefined) {
    if (!planilha || !lote || processando) return
    setErro('')
    setProcessando('Cruzando produtos por código de envio…')
    try {
      if (!/\.xlsx$/i.test(planilha.name)) throw new Error('Selecione a planilha de vendas original no formato .xlsx.')
      const encontrados = combinarProdutos(lote.pedidos, await lerEnviosDaPlanilha(planilha))
      if (![...encontrados.values()].some((lista) => lista.length)) throw new Error('A planilha não contém produtos dos envios deste PDF. Confira se os dois arquivos são do mesmo lote.')
      setProdutos((anteriores) => new Map([...anteriores, ...[...encontrados].filter(([, lista]) => lista.length)]))
    } catch (falha) {
      setErro(falha instanceof Error ? falha.message : 'Não foi possível ler a planilha.')
    } finally {
      setProcessando('')
    }
  }

  function receberPdf(evento: DragEvent) {
    evento.preventDefault()
    profundidade.current = 0
    setArrastando(false)
    void importarPdf(evento.dataTransfer.files)
  }

  function registrar(texto: string) {
    const pedido = codigos.get(chaveQr(texto))
    if (!pedido || !pedido.destinatario) {
      setBipagem({ tipo: 'desconhecido' })
      setDesconhecidos((total) => total + 1)
      return
    }
    if (validos.current.has(pedido.envio)) {
      setBipagem({ tipo: 'repetido', pedido })
      return
    }
    validos.current.add(pedido.envio)
    setQuantidade(validos.current.size)
    setBipagem({ tipo: 'identificado', pedido })
    navigator.vibrate?.(45)
  }

  const pendentes = (lote?.pedidos.length ?? 0) - quantidade
  const totalPaginasIgnoradas = lote?.paginasNaoReconhecidas.length ?? 0
  const semNome = lote?.pedidos.filter(({ destinatario }) => !destinatario).length ?? 0
  const mensagens = { identificado: 'Pedido identificado', repetido: 'Pedido já identificado', desconhecido: 'Pedido não identificado' }

  return (
    <div className="conferencia">
      <section className="panel" aria-labelledby="titulo-pdf-conferencia">
        <div className="panel-heading">
          <span className="section-icon"><FileText size={22} aria-hidden="true" /></span>
          <div><h2 id="titulo-pdf-conferencia">Etiquetas do dia</h2><p>Envie o PDF de etiquetas do Mercado Livre para identificar os destinatários.</p></div>
        </div>
        <div
          className={'dropzone conferencia-dropzone' + (arrastando ? ' dropzone-active' : '')}
          onDragEnter={(evento) => {
            evento.preventDefault()
            if (!processando && evento.dataTransfer.types.includes('Files')) {
              profundidade.current++
              setArrastando(true)
            }
          }}
          onDragOver={(evento) => evento.preventDefault()}
          onDragLeave={(evento) => {
            evento.preventDefault()
            profundidade.current = Math.max(0, profundidade.current - 1)
            if (!profundidade.current) setArrastando(false)
          }}
          onDrop={receberPdf}
        >
          <span className="upload-icon"><UploadCloud size={25} aria-hidden="true" /></span>
          <p className="dropzone-title">{arrastando ? 'Solte o PDF aqui' : 'Arraste e solte o PDF aqui'}</p>
          <p className="dropzone-caption">Formato aceito .pdf</p>
          <input type="file" accept=".pdf,application/pdf" hidden ref={inputPdf} onChange={(evento) => {
            void importarPdf(evento.currentTarget.files)
            evento.currentTarget.value = ''
          }} />
          <button type="button" className="button button-outline" disabled={Boolean(processando)} onClick={() => inputPdf.current?.click()}>
            <FolderOpen size={17} aria-hidden="true" /> {lote ? 'Trocar PDF' : 'Selecionar PDF de etiquetas'}
          </button>
        </div>
        {arquivo && <div className="selected-file"><FileText size={20} aria-hidden="true" /><div className="file-copy"><span className="file-name">{arquivo}</span><span className="file-meta">{lote?.pedidos.length} {lote?.pedidos.length === 1 ? 'envio encontrado' : 'envios encontrados'}</span></div></div>}
        {processando && <p className="notice notice-info" role="status"><LoaderCircle className="spinner" size={18} aria-hidden="true" /> {processando}</p>}
        {erro && <p className="notice notice-error" role="alert"><CircleAlert size={18} aria-hidden="true" /> {erro}</p>}
        {Boolean(semQr || totalPaginasIgnoradas || semNome) && (
          <div className="notice notice-warning" role="status"><CircleAlert size={18} aria-hidden="true" /><div>
            {semQr > 0 && <p>{semQr} envio(s) sem QR confirmado no PDF.</p>}
            {semNome > 0 && <p>{semNome} envio(s) sem destinatário identificado no PDF.</p>}
            {totalPaginasIgnoradas > 0 && <p>Páginas que não puderam ser reconhecidas: {lote?.paginasNaoReconhecidas.join(', ')}.</p>}
          </div></div>
        )}
        {lote && faltamProdutos > 0 && (
          <div className="conferencia-planilha">
            <p>Há dados de produto ou quantidade pendentes em {faltamProdutos} {faltamProdutos === 1 ? 'envio' : 'envios'}. Para conferir os itens, carregue a planilha de vendas do mesmo lote.</p>
            <input type="file" accept={'.xlsx,' + tipoExcel} hidden ref={inputExcel} onChange={(evento) => {
              void importarExcel(evento.currentTarget.files?.[0])
              evento.currentTarget.value = ''
            }} />
            <button type="button" className="button button-outline" disabled={Boolean(processando)} onClick={() => inputExcel.current?.click()}>Carregar planilha de vendas</button>
          </div>
        )}
      </section>

      {lote && (
        <section className="panel" aria-labelledby="titulo-bipagem" ref={areaCamera}>
          <div className="panel-heading">
            <span className="section-icon"><ScanLine size={22} aria-hidden="true" /></span>
            <div><h2 id="titulo-bipagem">Validar pedidos</h2><p>{quantidade} de {lote.pedidos.length} envios bipados</p></div>
          </div>
          <progress className="progresso-bipagem" value={quantidade} max={lote.pedidos.length} aria-label="Progresso da conferência" />
          {ativo && camera && <CameraConferencia aoLer={registrar} />}
          {bipagem && !finalizado && (
            <div className={'notice aviso-bipagem ' + ({ identificado: 'notice-success', repetido: 'notice-info', desconhecido: 'notice-warning' }[bipagem.tipo])} role="status">
              {bipagem.tipo === 'identificado' ? <CheckCircle2 size={21} aria-hidden="true" /> : <CircleAlert size={21} aria-hidden="true" />}
              <div><strong>{mensagens[bipagem.tipo]}</strong>
                {bipagem.pedido ? <DadosPedido pedido={bipagem.pedido} produtos={produtos.get(bipagem.pedido.envio)} /> : <p>Este QR não pôde ser associado a um destinatário do PDF. Você pode seguir bipando.</p>}
              </div>
            </div>
          )}
          <div className="conferencia-acoes">
            <button type="button" className={'button ' + (camera ? 'button-outline' : 'button-primary')} disabled={!codigos.size || Boolean(processando)} onClick={() => {
              setFinalizado(false)
              setCamera(!camera)
            }}><Camera size={19} aria-hidden="true" /> {camera ? 'Fechar câmera' : 'Abrir camera para validar pedidos'}</button>
            {!finalizado && <button type="button" className="button button-outline" disabled={Boolean(processando)} onClick={() => { setCamera(false); setFinalizado(true) }}><ClipboardCheck size={19} aria-hidden="true" /> Finalizar</button>}
          </div>
        </section>
      )}

      {finalizado && lote && (
        <section className="panel" aria-labelledby="titulo-resumo-conferencia">
          <div className="panel-heading">
            <span className="section-icon"><ClipboardCheck size={22} aria-hidden="true" /></span>
            <div><h2 id="titulo-resumo-conferencia" tabIndex={-1} ref={areaResumo}>Resultado da conferência</h2>
              <p>{!pendentes && !totalPaginasIgnoradas ? 'Todos os envios do PDF foram bipados.' : 'Confira os envios pendentes e os avisos de leitura do PDF.'}</p></div>
          </div>
          <dl className="conferencia-totais">
            <div><dt>Identificados</dt><dd>{quantidade} / {lote.pedidos.length}</dd></div>
            <div><dt>Pendentes</dt><dd>{pendentes}</dd></div>
            <div><dt>Leituras sem identificação</dt><dd>{desconhecidos}</dd></div>
          </dl>
          <ul className="lista-conferencia">
            {lote.pedidos.map((pedido) => {
              const confirmado = validos.current.has(pedido.envio)
              return <li key={pedido.envio}><DadosPedido pedido={pedido} produtos={produtos.get(pedido.envio)} /><span className={'estado-pedido ' + (confirmado ? 'estado-identificado' : 'estado-pendente')}>{confirmado ? 'Identificado' : 'Pendente'}</span></li>
            })}
          </ul>
        </section>
      )}
    </div>
  )
}
