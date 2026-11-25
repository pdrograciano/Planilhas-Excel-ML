import {
  CheckCircle2, CircleAlert, Clock3, Download, FileSpreadsheet, FolderOpen,
  History, LoaderCircle, Menu, Moon, ScanLine, ShieldCheck, Sun, Trash2, UploadCloud, WandSparkles, X,
} from 'lucide-react'
import { lazy, Suspense, useLayoutEffect, useRef, useState } from 'react'
import type { DragEvent } from 'react'
import { baixarArquivo, nomeBaseGerado, nomeDaVersao, processarArquivo, renumerarArquivoPorItens, tipoExcel } from './formatador/arquivo.ts'
import { excluirDoHistorico, guardarNoHistorico, guardarNumeracaoPorItens, identificadorHistorico, limparHistorico } from './historico.ts'
import type { carregarHistorico, EntradaHistorico } from './historico.ts'

const dataHora = new Intl.DateTimeFormat('pt-BR', { dateStyle: 'short', timeStyle: 'short' })
const numeros = new Intl.NumberFormat('pt-BR')
const chaveTema = 'mixbonix:tema'
const logoClaro = new URL('./assets/mixbonix-claro.png', import.meta.url).href
const logoEscuro = new URL('./assets/mixbonix-escuro.png', import.meta.url).href
const Conferencia = lazy(() => import('./conferencia/Conferencia.tsx'))
const abas = [
  { id: 'planilhas', titulo: 'Organização de pedidos e planilhas', tituloMobile: 'Pedidos e planilhas', Icone: FileSpreadsheet },
  { id: 'conferencia', titulo: 'Conferência de pedidos do dia', tituloMobile: 'Conferência do dia', Icone: ScanLine },
] as const

function totalLinhas(total: number) {
  return numeros.format(total) + (total === 1 ? ' linha' : ' linhas')
}

function tamanhoArquivo(bytes: number) {
  if (bytes < 1024) return numeros.format(bytes) + ' B'
  const megabyte = 1024 ** 2
  const divisor = bytes < megabyte ? 1024 : megabyte
  return new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 1 }).format(bytes / divisor) +
    (divisor === 1024 ? ' KB' : ' MB')
}

function DownloadsGerados({ entrada, resultado = false }: { entrada: EntradaHistorico; resultado?: boolean }) {
  const [corrigido, setCorrigido] = useState<Blob | null>(null)
  const [corrigindo, setCorrigindo] = useState(false)
  const [erro, setErro] = useState('')
  const nomeBase = entrada.nomeGerado ?? nomeBaseGerado(entrada.data)
  const porItens = entrada.porItens

  async function baixarPorItens() {
    if (!porItens || corrigindo) return
    const nome = nomeDaVersao(nomeBase, 'PorItens')
    if (corrigido || entrada.porItensNumerado) {
      baixarArquivo(corrigido ?? porItens, nome)
      return
    }
    setCorrigindo(true)
    setErro('')
    try {
      const arquivo = await renumerarArquivoPorItens(porItens)
      setCorrigido(arquivo)
      if (entrada.salva) void guardarNumeracaoPorItens(entrada.id, arquivo)
      baixarArquivo(arquivo, nome)
    } catch (falha) {
      setErro(falha instanceof Error ? falha.message : 'Não foi possível baixar o Arquivo Ordenado por Itens.')
    } finally {
      setCorrigindo(false)
    }
  }

  return (
    <>
      <button className="button button-outline" type="button" onClick={() => baixarArquivo(entrada.original, entrada.nome)}>
        <Download size={17} aria-hidden="true" /> Arquivo original ML
      </button>
      <button
        className={'button ' + (resultado ? 'button-primary' : 'button-download')}
        type="button"
        onClick={() => baixarArquivo(entrada.formatada, nomeDaVersao(nomeBase, 'Formatado'))}
      >
        <Download size={17} aria-hidden="true" /> Arquivo Formatado
      </button>
      {porItens && (
        <button
          className="button button-download"
          type="button"
          onClick={baixarPorItens}
          disabled={corrigindo}
          aria-busy={corrigindo}
        >
          {corrigindo ? <LoaderCircle className="spinner" size={17} aria-hidden="true" /> : <Download size={17} aria-hidden="true" />}
          Arquivo Ordenado por Itens
        </button>
      )}
      {erro && <p className="notice notice-error" role="alert">{erro}</p>}
    </>
  )
}

export function App({ historicoInicial }: { historicoInicial: ReturnType<typeof carregarHistorico> }) {
  const [aba, setAba] = useState(0)
  const [menuAberto, setMenuAberto] = useState(false)
  const [abriuConferencia, setAbriuConferencia] = useState(false)
  const [escuro, setEscuro] = useState(() => {
    try {
      return window.localStorage.getItem(chaveTema) === 'escuro'
    } catch {
      return false
    }
  })
  const [arquivo, setArquivo] = useState<File | null>(null)
  const [entradas, setEntradas] = useState(historicoInicial.entradas)
  const [resultado, setResultado] = useState<EntradaHistorico | null>(null)
  const [erro, setErro] = useState('')
  const [erroHistorico, setErroHistorico] = useState('')
  const [processando, setProcessando] = useState(false)
  const [arrastando, setArrastando] = useState(false)
  const input = useRef<HTMLInputElement>(null)
  const botaoMenu = useRef<HTMLButtonElement>(null)
  const profundidade = useRef(0)
  const emExecucao = useRef(false)

  function abrirAba(indice: number) {
    setAba(indice)
    if (indice === 1) setAbriuConferencia(true)
  }

  function fecharMenu() {
    setMenuAberto(false)
    if (botaoMenu.current?.offsetParent) botaoMenu.current.focus()
  }

  useLayoutEffect(() => {
    const tema = escuro ? 'escuro' : 'claro'
    const raiz = document.documentElement
    raiz.dataset.tema = tema
    const corDeFundo = getComputedStyle(raiz).getPropertyValue('--canvas').trim()
    document.querySelector('meta[name="theme-color"]')?.setAttribute('content', corDeFundo)
    try {
      window.localStorage.setItem(chaveTema, tema)
    } catch {
      // A preferência ainda vale para a página atual.
    }
  }, [escuro])

  function selecionar(arquivos: FileList | null) {
    if (!arquivos?.length || emExecucao.current) return
    setResultado(null)
    if (arquivos.length !== 1 || !/\.xlsx$/i.test(arquivos[0].name)) {
      setArquivo(null)
      setErro('Selecione um único arquivo no formato .xlsx.')
      return
    }
    setErro('')
    setArquivo(arquivos[0])
  }

  function receber(evento: DragEvent) {
    evento.preventDefault()
    profundidade.current = 0
    setArrastando(false)
    selecionar(evento.dataTransfer.files)
  }

  function limparArquivo() {
    setArquivo(null)
    setResultado(null)
    setErro('')
    document.getElementById('selecionar-arquivo')?.focus()
  }

  async function gerar() {
    if (!arquivo || emExecucao.current) return
    emExecucao.current = true
    setProcessando(true)
    setErro('')
    setResultado(null)
    try {
      const { formatada, porItens, linhas } = await processarArquivo(arquivo)
      const data = Date.now()
      const entrada = {
        id: identificadorHistorico(),
        nome: arquivo.name,
        nomeGerado: nomeBaseGerado(data),
        data,
        linhas,
        original: arquivo,
        formatada,
        porItens,
        porItensNumerado: true,
      }
      const salva = await guardarNoHistorico(entrada)
      const concluida = { ...entrada, salva }
      setEntradas((anteriores) => [concluida, ...anteriores])
      setResultado(concluida)
    } catch (falha) {
      setErro(falha instanceof Error ? falha.message : 'Não foi possível gerar a planilha. Tente novamente.')
    } finally {
      emExecucao.current = false
      setProcessando(false)
    }
  }

  function apagarEntrada(entrada: EntradaHistorico) {
    if (emExecucao.current) return
    if (entrada.salva && !excluirDoHistorico(entrada.id)) {
      setErroHistorico('Não foi possível apagar esta entrada do histórico.')
      return
    }
    setErroHistorico('')
    setEntradas((anteriores) => anteriores.filter(({ id }) => id !== entrada.id))
    if (resultado?.id === entrada.id) setResultado(null)
  }

  function apagarTudo() {
    if (emExecucao.current || !window.confirm('Limpar todo o histórico desta guia?')) return
    if (!limparHistorico() && entradas.some(({ salva }) => salva)) {
      setErroHistorico('Não foi possível limpar o histórico desta guia.')
      return
    }
    setErroHistorico('')
    setEntradas([])
    setResultado(null)
  }

  return (
    <>
      <header className="topbar" onKeyDown={(evento) => {
        if (evento.key === 'Escape' && menuAberto) fecharMenu()
      }}>
        <div className="topbar-inner">
          <img className="brand-logo" src={escuro ? logoEscuro : logoClaro} width={444} height={204} alt="MIXBONIX" />
          <div className="topbar-actions">
            <span className="privacy"><ShieldCheck size={17} aria-hidden="true" /> Processamento local</span>
            <button
              className="icon-button theme-toggle"
              type="button"
              aria-label="Modo escuro"
              aria-pressed={escuro}
              title={escuro ? 'Ativar modo claro' : 'Ativar modo escuro'}
              onClick={() => setEscuro(!escuro)}
            >
              {escuro ? <Moon size={19} aria-hidden="true" /> : <Sun size={20} aria-hidden="true" />}
            </button>
            <button className="icon-button botao-menu" type="button" ref={botaoMenu}
              aria-label={menuAberto ? 'Fechar menu' : 'Abrir menu'} aria-expanded={menuAberto} aria-controls="abas-mixbonix"
              onClick={() => setMenuAberto(!menuAberto)}>
              {menuAberto ? <X size={22} aria-hidden="true" /> : <Menu size={22} aria-hidden="true" />}
            </button>
          </div>
        </div>
        <div id="abas-mixbonix" className="topbar-tabs" data-menu-aberto={menuAberto} role="tablist" aria-label="Ferramentas MIXBONIX" onKeyDown={(evento) => {
          if (!['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End'].includes(evento.key)) return
          evento.preventDefault()
          const proxima = evento.key === 'Home' ? 0 : evento.key === 'End' ? abas.length - 1 : 1 - aba
          abrirAba(proxima)
          document.getElementById('aba-' + abas[proxima].id)?.focus()
        }}>
          {abas.map(({ id, titulo, Icone }, indice) => (
            <button className="topbar-tab" type="button" role="tab" key={id} id={'aba-' + id}
              aria-selected={aba === indice} aria-controls={'painel-' + id} tabIndex={aba === indice ? 0 : -1}
              onClick={() => {
                abrirAba(indice)
                if (menuAberto) fecharMenu()
              }}>
              <Icone size={18} aria-hidden="true" /> <span>{titulo}</span>
            </button>
          ))}
        </div>
      </header>

      <main className="workspace">
        <div className="intro">
          <h1><span className="titulo-desktop">MIXBONIX ORDER</span><span className="titulo-mobile">{abas[aba].tituloMobile}</span></h1>
          {aba === 0 && <p>Conversor de Excel de vendas do Mercado Livre para o modelo da MIXBONIX.</p>}
        </div>

        <div role="tabpanel" id="painel-planilhas" aria-labelledby="aba-planilhas" hidden={aba !== 0}>
        <section className="panel" aria-labelledby="titulo-formatar">
          <div className="panel-heading">
            <span className="section-icon"><FileSpreadsheet size={21} aria-hidden="true" /></span>
            <div>
              <h2 id="titulo-formatar">Formatar planilha</h2>
              <p>Selecione o relatório exportado do Mercado Livre.</p>
            </div>
          </div>

          <div
            className={'dropzone' + (arrastando ? ' dropzone-active' : '')}
            onDragEnter={(evento) => {
              evento.preventDefault()
              if (!emExecucao.current && evento.dataTransfer.types.includes('Files')) {
                profundidade.current++
                setArrastando(true)
              }
            }}
            onDragOver={(evento) => {
              evento.preventDefault()
              evento.dataTransfer.dropEffect = emExecucao.current ? 'none' : 'copy'
            }}
            onDragLeave={(evento) => {
              evento.preventDefault()
              profundidade.current = Math.max(0, profundidade.current - 1)
              if (!profundidade.current) setArrastando(false)
            }}
            onDrop={receber}
          >
            <span className="upload-icon"><UploadCloud size={25} strokeWidth={1.9} aria-hidden="true" /></span>
            <p className="dropzone-title">{arrastando ? 'Solte a planilha aqui' : 'Arraste e solte a planilha aqui'}</p>
            <p className="dropzone-caption">Formato aceito .xlsx</p>
            <input
              type="file"
              accept={'.xlsx,' + tipoExcel}
              hidden
              ref={input}
              onChange={(evento) => {
                selecionar(evento.currentTarget.files)
                evento.currentTarget.value = ''
              }}
            />
            <button
              className="button button-outline"
              id="selecionar-arquivo"
              type="button"
              disabled={processando}
              onClick={() => input.current?.click()}
            >
              <FolderOpen size={17} aria-hidden="true" />
              {arquivo ? 'Trocar arquivo' : 'Selecionar arquivo'}
            </button>
          </div>

          {arquivo && (
            <div className="selected-file">
              <FileSpreadsheet size={20} aria-hidden="true" />
              <div className="file-copy">
                <span className="file-name">{arquivo.name}</span>
                <span className="file-meta">{tamanhoArquivo(arquivo.size)}</span>
              </div>
              <button className="icon-button" type="button" onClick={limparArquivo} disabled={processando} aria-label="Remover arquivo selecionado">
                <X size={18} aria-hidden="true" />
              </button>
            </div>
          )}

          {erro && (
            <div className="notice notice-error" role="alert">
              <CircleAlert size={20} aria-hidden="true" />
              <span>{erro}</span>
            </div>
          )}

          <div className="process-actions">
            <button className={'button ' + (resultado ? 'button-outline' : 'button-primary')} type="button" disabled={!arquivo || processando} onClick={gerar}>
              {processando ? <LoaderCircle className="spinner" size={18} aria-hidden="true" /> : <WandSparkles size={18} aria-hidden="true" />}
              {processando ? 'Formatando planilhas…' : 'Gerar planilhas'}
            </button>
          </div>
          <div className="sr-only" role="status">{processando ? 'Processando o arquivo selecionado.' : ''}</div>

          {resultado && (
            <div className="result">
              <div className="result-copy" role="status">
                <CheckCircle2 size={21} aria-hidden="true" />
                <div>
                  <strong>Planilhas prontas</strong>
                  <p>{totalLinhas(resultado.linhas)} de {resultado.linhas === 1 ? 'pedido' : 'pedidos'}.</p>
                </div>
              </div>
              <div className="download-actions">
                <DownloadsGerados entrada={resultado} resultado />
              </div>
            </div>
          )}
        </section>

        <section className="panel history-panel" aria-labelledby="titulo-historico">
          <div className="panel-heading history-heading">
            <span className="section-icon"><History size={22} aria-hidden="true" /></span>
            <div>
              <h2 id="titulo-historico">Histórico</h2>
              <p>Os arquivos ficam disponíveis nesta guia.</p>
            </div>
            {entradas.length > 0 && (
              <button className="button button-danger" type="button" disabled={processando} onClick={apagarTudo}>
                Limpar histórico
              </button>
            )}
          </div>
          {erroHistorico && (
            <div className="notice notice-error" role="alert">
              <CircleAlert size={19} aria-hidden="true" /> <span>{erroHistorico}</span>
            </div>
          )}
          {historicoInicial.aviso && (
            <div className="notice notice-warning" role="status">
              <CircleAlert size={19} aria-hidden="true" /> <span>{historicoInicial.aviso}</span>
            </div>
          )}
          {entradas.length === 0 ? (
            <div className="empty">
              <span className="empty-icon"><Clock3 size={22} strokeWidth={1.8} aria-hidden="true" /></span>
              <strong>Nenhuma planilha formatada</strong>
              <p>As planilhas que você formatar aparecerão aqui.</p>
            </div>
          ) : (
            <ul className="history-list">
              {entradas.map((entrada) => (
                <li className="history-row" key={entrada.id}>
                  <div className="history-file">
                    <span className="history-file-icon"><FileSpreadsheet size={19} aria-hidden="true" /></span>
                    <div className="file-copy">
                      <span className="file-name">{entrada.nome}</span>
                      <span className="file-meta">
                        <time dateTime={new Date(entrada.data).toISOString()}>{dataHora.format(entrada.data)}</time>
                        <span aria-hidden="true"> · </span>{totalLinhas(entrada.linhas)}
                      </span>
                    </div>
                  </div>
                  <div className="download-actions history-actions">
                    <DownloadsGerados entrada={entrada} />
                  </div>
                  <button
                    className="icon-button history-delete"
                    type="button"
                    aria-label={'Apagar do histórico ' + entrada.nome}
                    title="Apagar do histórico"
                    disabled={processando}
                    onClick={() => apagarEntrada(entrada)}
                  >
                    <Trash2 size={18} aria-hidden="true" />
                  </button>
                  {!entrada.salva && (
                    <p className="notice notice-warning history-warning" role="status">
                      <CircleAlert size={18} aria-hidden="true" />
                      <span>O navegador não conseguiu guardar estes arquivos. Baixe-os antes de recarregar a página.</span>
                    </p>
                  )}
                </li>
              ))}
            </ul>
          )}
        </section>
        </div>
        <div role="tabpanel" id="painel-conferencia" aria-labelledby="aba-conferencia" hidden={aba !== 1}>
          {abriuConferencia && (
            <Suspense fallback={<p className="notice notice-info" role="status">Abrindo conferência…</p>}>
              <Conferencia ativo={aba === 1} entradas={entradas} />
            </Suspense>
          )}
        </div>
      </main>
      <footer className="site-footer">
        <div className="site-footer-inner">
          <small>© {new Date().getFullYear()} MIXBONIX. Todos os direitos reservados.</small>
          <small>Desenvolvido por Pedro Graciano.</small>
        </div>
      </footer>
    </>
  )
}
