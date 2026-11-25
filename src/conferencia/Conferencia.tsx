import { Camera, CheckCircle2, ChevronDown, CircleAlert, ClipboardCheck, Download, FileText, FolderOpen, LoaderCircle, ScanLine, Trash2, UploadCloud } from 'lucide-react'
import { useEffect, useMemo, useRef, useState } from 'react'
import type { DragEvent } from 'react'
import { baixarArquivo, lerEnviosDaPlanilha, tipoExcel } from '../formatador/arquivo.ts'
import type { EntradaHistorico } from '../historico.ts'
import { CameraConferencia } from './CameraConferencia.tsx'
import { prepararSom, tocarBipe } from './bipes.ts'
import type { TipoBipe } from './bipes.ts'
import { chaveQr, produtosPorEnvio, reunirPedidos } from './pedidos.ts'
import type { LoteConferencia, PedidoConferencia, ProdutoConferencia } from './pedidos.ts'
import type { DadosRelatorioConferencia } from './relatorio.ts'

interface Bipagem {
  tipo: TipoBipe
  pedido?: PedidoConferencia
}

interface PdfConferencia extends LoteConferencia {
  id: number
  nome: string
}

interface PdfRelatorio {
  relatorio: DadosRelatorioConferencia
  download: { arquivo: Blob; nome: string } | null
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
  const proximoPdf = useRef(0)
  const emProcessamento = useRef(false)
  const validos = useRef(new Set<string>())
  const audio = useRef<AudioContext | null>(null)
  const [arrastando, setArrastando] = useState(false)
  const [arquivos, setArquivos] = useState<PdfConferencia[]>([])
  const [produtos, setProdutos] = useState(new Map<string, ProdutoConferencia[]>())
  const [processando, setProcessando] = useState('')
  const [erro, setErro] = useState('')
  const [camera, setCamera] = useState(false)
  const [bipagem, setBipagem] = useState<Bipagem | null>(null)
  const [quantidade, setQuantidade] = useState(0)
  const [desconhecidos, setDesconhecidos] = useState(0)
  const [relatorio, setRelatorio] = useState<DadosRelatorioConferencia | null>(null)
  const [pdfRelatorio, setPdfRelatorio] = useState<PdfRelatorio | null>(null)
  const finalizado = relatorio !== null
  const pdfAtual = pdfRelatorio?.relatorio === relatorio ? pdfRelatorio : null
  const pedidos = useMemo(() => reunirPedidos(arquivos), [arquivos])
  const codigos = useMemo(() => {
    const porEnvio = new Map(pedidos.map((pedido) => [pedido.envio, pedido]))
    return new Map(arquivos.flatMap((arquivo) => arquivo.pedidos.flatMap(({ codigo, envio }) =>
      codigo ? [[codigo, porEnvio.get(envio)!] as const] : [])))
  }, [arquivos, pedidos])
  const faltamProdutos = pedidos.filter(({ envio }) => {
    const itens = produtos.get(envio)
    return !itens?.length || itens.some(({ quantidade }) => quantidade === null)
  }).length
  const semQr = pedidos.filter(({ codigo }) => !codigo).length

  useEffect(() => { if (!ativo) setCamera(false) }, [ativo])
  useEffect(() => { if (camera) areaCamera.current?.scrollIntoView({ block: 'nearest' }) }, [camera])
  useEffect(() => { if (finalizado) areaResumo.current?.focus() }, [finalizado])
  useEffect(() => () => {
    if (audio.current) void audio.current.close().catch(() => {})
    audio.current = null
  }, [])

  useEffect(() => {
    if (!relatorio) {
      setPdfRelatorio(null)
      return
    }
    let cancelado = false
    const dados = relatorio
    async function gerar() {
      try {
        const { gerarRelatorioConferencia } = await import('./relatorio.ts')
        const download = await gerarRelatorioConferencia(dados)
        if (cancelado) return
        setPdfRelatorio({ relatorio: dados, download })
        baixarArquivo(download.arquivo, download.nome)
      } catch {
        if (!cancelado) setPdfRelatorio({ relatorio: dados, download: null })
      }
    }
    void gerar()
    return () => { cancelado = true }
  }, [relatorio])

  async function importarPdfs(selecao: FileList | null) {
    if (!selecao?.length || emProcessamento.current) return
    const selecionados = Array.from(selecao)
    const novos: PdfConferencia[] = []
    const falhas: string[] = []
    emProcessamento.current = true
    setErro('')
    setCamera(false)
    setProcessando('Abrindo etiquetas…')
    try {
      const { lerPdfEtiquetas } = await import('./leituraPdf.ts')
      for (const [indice, pdf] of selecionados.entries()) {
        if (!/\.pdf$/i.test(pdf.name)) {
          falhas.push(`${pdf.name} precisa estar no formato PDF.`)
          continue
        }
        try {
          const lote = await lerPdfEtiquetas(pdf, (atual, total) => {
            setProcessando(`PDF ${indice + 1} de ${selecionados.length} · Página ${atual} de ${total}`)
          })
          novos.push({ ...lote, id: proximoPdf.current++, nome: pdf.name })
        } catch (falha) {
          falhas.push(falha instanceof Error && falha.message.startsWith('Nenhum envio')
            ? `Nenhum envio encontrado em ${pdf.name}.` : `Não foi possível ler ${pdf.name}.`)
        }
      }
      if (novos.length) {
        const pendentes = new Set(reunirPedidos([...arquivos, ...novos]).map(({ envio }) => envio).filter((envio) => !produtos.has(envio)))
        const encontrados = new Map(produtos)
        if (pendentes.size && entradas.length) setProcessando('Buscando produtos no histórico…')
        for (const entrada of entradas) {
          if (!pendentes.size) break
          try {
            const dados = produtosPorEnvio(await lerEnviosDaPlanilha(entrada.original))
            for (const [envio, lista] of dados) {
              if (pendentes.delete(envio)) encontrados.set(envio, lista)
            }
          } catch {
            // Uma planilha antiga ilegível não impede a conferência.
          }
        }
        setArquivos((anteriores) => [...anteriores, ...novos])
        setProdutos(encontrados)
        setRelatorio(null)
      }
      setErro(falhas.join(' '))
    } catch {
      setErro('Não foi possível iniciar a leitura dos PDFs. Tente novamente.')
    } finally {
      emProcessamento.current = false
      setProcessando('')
    }
  }

  async function importarExcel(planilha: File | undefined) {
    if (!planilha || !pedidos.length || emProcessamento.current) return
    emProcessamento.current = true
    setErro('')
    setProcessando('Cruzando produtos por código de envio…')
    try {
      if (!/\.xlsx$/i.test(planilha.name)) throw new Error('Selecione a planilha de vendas original no formato .xlsx.')
      const encontrados = produtosPorEnvio(await lerEnviosDaPlanilha(planilha))
      if (!pedidos.some(({ envio }) => encontrados.has(envio))) throw new Error('A planilha não contém produtos dos envios adicionados.')
      setProdutos((anteriores) => new Map([...anteriores, ...encontrados]))
      setRelatorio(null)
    } catch (falha) {
      setErro(falha instanceof Error ? falha.message : 'Não foi possível ler a planilha.')
    } finally {
      emProcessamento.current = false
      setProcessando('')
    }
  }

  function receberPdf(evento: DragEvent) {
    evento.preventDefault()
    profundidade.current = 0
    setArrastando(false)
    void importarPdfs(evento.dataTransfer.files)
  }

  function removerPdf(id: number) {
    if (emProcessamento.current) return
    const restantes = arquivos.filter((arquivo) => arquivo.id !== id)
    const envios = new Set(reunirPedidos(restantes).map(({ envio }) => envio))
    validos.current = new Set([...validos.current].filter((envio) => envios.has(envio)))
    setQuantidade(validos.current.size)
    setArquivos(restantes)
    setRelatorio(null)
    if (bipagem?.pedido && !envios.has(bipagem.pedido.envio)) setBipagem(null)
    if (!restantes.length) {
      setProdutos(new Map())
      setDesconhecidos(0)
      setBipagem(null)
      setCamera(false)
      setErro('')
    }
  }

  function registrar(texto: string) {
    const pedido = codigos.get(chaveQr(texto))
    if (!pedido || !pedido.destinatario) {
      setBipagem({ tipo: 'desconhecido' })
      setDesconhecidos((total) => total + 1)
      tocarBipe(audio.current, 'desconhecido')
      return
    }
    if (validos.current.has(pedido.envio)) {
      setBipagem({ tipo: 'repetido', pedido })
      tocarBipe(audio.current, 'repetido')
      return
    }
    validos.current.add(pedido.envio)
    setQuantidade(validos.current.size)
    setBipagem({ tipo: 'identificado', pedido })
    tocarBipe(audio.current, 'identificado')
    navigator.vibrate?.(45)
  }

  function finalizar() {
    setCamera(false)
    setRelatorio({
      data: new Date(),
      desconhecidos,
      arquivos: arquivos.map(({ nome, paginasNaoReconhecidas }) => ({ nome, paginasNaoReconhecidas })),
      pedidos: pedidos.map(({ envio, destinatario }) => ({
        envio, destinatario, produtos: produtos.get(envio) ?? [], identificado: validos.current.has(envio),
      })),
    })
  }

  const pendentes = pedidos.length - quantidade
  const arquivosComAvisos = arquivos.filter(({ paginasNaoReconhecidas }) => paginasNaoReconhecidas.length)
  const semNome = pedidos.filter(({ destinatario }) => !destinatario).length
  const mensagens = { identificado: 'Pedido identificado', repetido: 'Pedido já identificado', desconhecido: 'Pedido não identificado' }

  return (
    <div className={'conferencia' + (camera ? ' conferencia-em-leitura' : '')}>
      <section className="panel conferencia-arquivos" aria-labelledby="titulo-pdf-conferencia">
        <div className="panel-heading">
          <span className="section-icon"><FileText size={22} aria-hidden="true" /></span>
          <div><h2 id="titulo-pdf-conferencia">Etiquetas do dia</h2><p>{pedidos.length ? `${pedidos.length} ${pedidos.length === 1 ? 'envio encontrado' : 'envios encontrados'}` : 'Reúna os PDFs do dia em uma só conferência.'}</p></div>
        </div>
        <div
          className={'dropzone conferencia-dropzone' + (arquivos.length ? ' conferencia-dropzone-compacta' : '') + (arrastando ? ' dropzone-active' : '')}
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
          <p className="dropzone-title">{arrastando ? 'Solte os PDFs aqui' : 'Arraste os PDFs de etiquetas aqui'}</p>
          <input type="file" accept=".pdf,application/pdf" multiple hidden ref={inputPdf} onChange={(evento) => {
            void importarPdfs(evento.currentTarget.files)
            evento.currentTarget.value = ''
          }} />
          <button type="button" className="button button-outline" disabled={Boolean(processando)} onClick={() => inputPdf.current?.click()}>
            <FolderOpen size={17} aria-hidden="true" /> {arquivos.length ? 'Adicionar PDFs' : 'Selecionar PDFs de etiquetas'}
          </button>
        </div>
        {arquivos.length > 0 && (
          <details className="conferencia-detalhes">
            <summary>{arquivos.length} {arquivos.length === 1 ? 'PDF adicionado' : 'PDFs adicionados'} <ChevronDown size={17} aria-hidden="true" /></summary>
            <ul className="pdfs-conferencia">
              {arquivos.map(({ id, nome, pedidos: etiquetas, paginasNaoReconhecidas }) => (
                <li key={id}>
                  <div className="file-copy">
                    <span className="file-name">{nome}</span>
                    <span className="file-meta">{etiquetas.length} {etiquetas.length === 1 ? 'envio' : 'envios'}</span>
                    {paginasNaoReconhecidas.length > 0 && <span className="file-meta">Páginas não reconhecidas: {paginasNaoReconhecidas.join(', ')}.</span>}
                  </div>
                  <button className="icon-button" type="button" aria-label={'Remover ' + nome} disabled={Boolean(processando)} onClick={() => removerPdf(id)}><Trash2 size={18} aria-hidden="true" /></button>
                </li>
              ))}
            </ul>
          </details>
        )}
        {processando && <p className="notice notice-info" role="status"><LoaderCircle className="spinner" size={18} aria-hidden="true" /> {processando}</p>}
        {erro && <p className="notice notice-error" role="alert"><CircleAlert size={18} aria-hidden="true" /> {erro}</p>}
        {Boolean(semQr || arquivosComAvisos.length || semNome) && (
          <div className="notice notice-warning" role="status"><CircleAlert size={18} aria-hidden="true" /><div>
            {semQr > 0 && <p>{semQr} envio(s) sem QR confirmado no PDF.</p>}
            {semNome > 0 && <p>{semNome} envio(s) sem destinatário identificado no PDF.</p>}
            {arquivosComAvisos.length > 0 && <p>Há páginas não reconhecidas. Veja os detalhes na lista de PDFs.</p>}
          </div></div>
        )}
        {faltamProdutos > 0 && (
          <details className="conferencia-detalhes">
            <summary>Vincular planilha de vendas <ChevronDown size={17} aria-hidden="true" /></summary>
            <div className="conferencia-planilha">
              <p>Faltam dados de produto ou quantidade em {faltamProdutos} {faltamProdutos === 1 ? 'envio' : 'envios'}.</p>
              <input type="file" accept={'.xlsx,' + tipoExcel} hidden ref={inputExcel} onChange={(evento) => {
                void importarExcel(evento.currentTarget.files?.[0])
                evento.currentTarget.value = ''
              }} />
              <button type="button" className="button button-outline" disabled={Boolean(processando)} onClick={() => inputExcel.current?.click()}>Carregar planilha de vendas</button>
            </div>
          </details>
        )}
      </section>

      {pedidos.length > 0 && (
        <section className="panel" aria-labelledby="titulo-bipagem" ref={areaCamera}>
          <div className="panel-heading">
            <span className="section-icon"><ScanLine size={22} aria-hidden="true" /></span>
            <div><h2 id="titulo-bipagem">Validar pedidos</h2><p>{quantidade} de {pedidos.length} envios bipados</p></div>
          </div>
          <progress className="progresso-bipagem" value={quantidade} max={pedidos.length} aria-label="Progresso da conferência" />
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
              if (!camera) {
                audio.current = prepararSom(audio.current)
                setRelatorio(null)
              }
              setCamera(!camera)
            }}><Camera size={19} aria-hidden="true" /> {camera ? 'Fechar câmera' : 'Abrir camera para validar pedidos'}</button>
            {!finalizado && <button type="button" className="button button-outline" disabled={Boolean(processando)} onClick={finalizar}><ClipboardCheck size={19} aria-hidden="true" /> Finalizar</button>}
          </div>
        </section>
      )}

      {relatorio && pedidos.length > 0 && (
        <section className="panel" aria-labelledby="titulo-resumo-conferencia">
          <div className="panel-heading">
            <span className="section-icon"><ClipboardCheck size={22} aria-hidden="true" /></span>
            <div><h2 id="titulo-resumo-conferencia" tabIndex={-1} ref={areaResumo}>Resultado da conferência</h2>
              <p>{!pendentes && !arquivosComAvisos.length ? 'Todos os envios do lote foram bipados.' : 'Confira os envios pendentes e os avisos de leitura.'}</p></div>
          </div>
          <dl className="conferencia-totais">
            <div><dt>Identificados</dt><dd>{quantidade} / {pedidos.length}</dd></div>
            <div><dt>Pendentes</dt><dd>{pendentes}</dd></div>
            <div><dt>Leituras sem identificação</dt><dd>{desconhecidos}</dd></div>
          </dl>
          <div className="acoes-relatorio">
            <button type="button" className="button button-primary" disabled={!pdfAtual}
              onClick={() => {
                if (pdfAtual?.download) baixarArquivo(pdfAtual.download.arquivo, pdfAtual.download.nome)
                else setRelatorio({ ...relatorio })
              }}>
              {!pdfAtual ? <LoaderCircle className="spinner" size={18} aria-hidden="true" /> : <Download size={18} aria-hidden="true" />}
              {!pdfAtual ? 'Gerando relatório…' : pdfAtual.download ? 'Baixar relatório em PDF' : 'Tentar gerar PDF novamente'}
            </button>
            {pdfAtual && !pdfAtual.download && <p className="notice notice-error" role="alert">Não foi possível gerar o relatório. Tente novamente.</p>}
          </div>
          <ul className="lista-conferencia">
            {pedidos.map((pedido) => {
              const confirmado = validos.current.has(pedido.envio)
              return <li key={pedido.envio}><DadosPedido pedido={pedido} produtos={produtos.get(pedido.envio)} /><span className={'estado-pedido ' + (confirmado ? 'estado-identificado' : 'estado-pendente')}>{confirmado ? 'Identificado' : 'Pendente'}</span></li>
            })}
          </ul>
        </section>
      )}
    </div>
  )
}
