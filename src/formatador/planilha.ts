import ExcelJS from 'exceljs'
import type { Border, Borders, CellValue, Fill, Worksheet } from 'exceljs'
import { larguraTexto, nomeCurto, normalizar } from './item.ts'
import type { EtiquetaEnvio } from './etiquetas.ts'

const campos = {
  unidades: { nome: 'Unidades', secao: 'vendas', posicao: 7 },
  titulo: { nome: 'Título do anúncio', secao: 'anuncios', posicao: 24 },
  dados: { nome: 'Dados pessoais ou da empresa', secao: 'faturamento ao comprador', posicao: 29 },
  endereco: { nome: 'Endereço', secao: 'faturamento ao comprador', posicao: 31 },
  venda: { nome: 'N.º de venda', secao: 'vendas', posicao: 1 },
  data: { nome: 'Data de venda', secao: 'vendas', posicao: 2 },
  estado: { nome: 'Estado', secao: 'vendas', posicao: 3 },
  pacote: { nome: 'Pacote de diversos produtos', secao: 'vendas', posicao: 5 },
  comprador: { nome: 'Comprador', secao: 'compradores', posicao: 34 },
  sku: { nome: 'SKU', secao: 'anuncios', posicao: 22 },
  envio: { nome: 'Identificador do envio', secao: 'envios', posicao: 46 },
} as const

type Campo = keyof typeof campos
type Cabecalhos = Map<string, number[]>

const obrigatorios = ['unidades', 'titulo', 'dados', 'endereco'] as const
const secoes = new Set([
  'vendas', 'publicidade', 'anuncios', 'faturamento ao comprador', 'compradores',
  'envios', 'devolucoes', 'reclamacoes',
])

const cabecalhoSaida = ['ORDEM', campos.unidades.nome, 'ITEM', campos.titulo.nome, campos.dados.nome, campos.endereco.nome]
const larguras = [12.6640625, 12.6640625, 18.6640625, 68.5546875, 36.21875, 82.6640625]
const coresCabecalho = ['FFF1C232', 'FFF1C232', 'FFF1C232', 'FFF6B26B', 'FFFFD966', 'FFFFD966']
const linhaBase = 22.5
const alfabeto = new Intl.Collator('pt-BR', { sensitivity: 'base' })

interface Pedido {
  unidades: number | null
  titulo: string
  dados: string
  endereco: string
}

type PedidoFormatado = Pedido & { item: string }

interface ReferenciaEnvio {
  envio: string
  venda: string
  comprador: string
}

type PedidoRelatorio = PedidoFormatado & { sku: string; referencia: ReferenciaEnvio; cancelado: boolean }

interface BlocoPedidos {
  resumo?: Pedido
  referencia: ReferenciaEnvio
  pedidos: PedidoRelatorio[]
}

export class ErroPlanilha extends Error {}

function textoCelula(valor: CellValue): string {
  if (valor == null) return ''
  if (typeof valor !== 'object') return String(valor)
  if (valor instanceof Date) return valor.toISOString()
  if ('richText' in valor) return valor.richText.map((parte) => parte.text).join('')
  if ('text' in valor) return valor.text ?? ''
  if ('result' in valor) return textoCelula(valor.result)
  if ('error' in valor) return valor.error
  return ''
}

function chaveCabecalho(texto: string): string {
  return normalizar(texto).replace(/^n\s*\.?\s*[º°o]?\s*\.?\s+de venda$/, 'numero de venda')
}

function encontrarCabecalho(aba: Worksheet) {
  let melhor: { linha: number; colunas: Cabecalhos; pontos: number } | undefined
  const nomes = Object.values(campos).map(({ nome }) => chaveCabecalho(nome))
  aba.eachRow((linha) => {
    if (melhor?.pontos === nomes.length) return
    const colunas: Cabecalhos = new Map()
    linha.eachCell((celula, numero) => {
      if (celula.isMerged && celula.address !== celula.master.address) return
      const nome = chaveCabecalho(textoCelula(celula.value))
      if (!nomes.includes(nome)) return
      const ocorrencias = colunas.get(nome) ?? []
      ocorrencias.push(numero)
      colunas.set(nome, ocorrencias)
    })
    if (colunas.size > (melhor?.pontos ?? 0)) {
      melhor = { linha: linha.number, colunas, pontos: colunas.size }
    }
  })
  if (!melhor) throw new ErroPlanilha('Colunas obrigatórias ausentes. Confira o cabeçalho da aba Vendas BR.')
  return melhor
}

function mapearColunas(aba: Worksheet, linha: number, colunas: Cabecalhos) {
  const inicios: { coluna: number; nome: string }[] = []
  if (linha > 1) {
    aba.getRow(linha - 1).eachCell((celula, coluna) => {
      if (celula.isMerged && celula.address !== celula.master.address) return
      const nome = normalizar(textoCelula(celula.value))
      if (secoes.has(nome)) inicios.push({ coluna, nome })
    })
  }
  inicios.reverse()
  const localizar = (campo: Campo) => {
    const { nome, secao, posicao } = campos[campo]
    const candidatas = colunas.get(chaveCabecalho(nome)) ?? []
    if (!inicios.length) {
      return candidatas.find((coluna) => coluna === posicao) ?? (candidatas.length === 1 ? candidatas[0] : undefined)
    }
    const naSecao = candidatas.filter((coluna) => inicios.find((inicio) => inicio.coluna <= coluna)?.nome === secao)
    return naSecao.find((coluna) => coluna === posicao) ?? (naSecao.length === 1 ? naSecao[0] : undefined)
  }
  const resultado = {
    unidades: localizar('unidades'),
    titulo: localizar('titulo'),
    dados: localizar('dados'),
    endereco: localizar('endereco'),
    venda: localizar('venda'),
    data: localizar('data'),
    estado: localizar('estado'),
    pacote: localizar('pacote'),
    comprador: localizar('comprador'),
    sku: localizar('sku'),
    envio: localizar('envio'),
  }
  const faltantes = obrigatorios.filter((campo) => resultado[campo] === undefined)
  if (faltantes.length) {
    const nomes = faltantes.map((campo) => campos[campo].nome).join(', ')
    throw new ErroPlanilha('Coluna obrigatória ausente ou em seção incorreta. Confira ' + nomes + ' na aba Vendas BR.')
  }
  return resultado as typeof resultado & Record<(typeof obrigatorios)[number], number>
}

function numeroUnidades(texto: string): number | null {
  const valor = texto.trim()
  if (!valor) return null
  if (/^\d{1,3}(?:\.\d{3})+(?:,\d+)?$/.test(valor)) {
    return Number(valor.replaceAll('.', '').replace(',', '.'))
  }
  return /^\d+(?:[.,]\d+)?$/.test(valor) ? Number(valor.replace(',', '.')) : NaN
}

const rodape = /^(?:(?:sub)?total|resumo|rodape|fim (?:do|da) |relatorio gerado)(?:\b|$)/
const tituloRodape = /^(?:total(?: geral| de vendas| de pedidos)?|subtotal|resumo|fim do relatorio)$/
const estadoPacote = /^pacote de (\d+) produtos?$/
const tituloPacote = /^\(\d+ pedidos? feitos?\)$/

function lerBlocos(aba: Worksheet): BlocoPedidos[] {
  const cabecalho = encontrarCabecalho(aba)
  const colunas = mapearColunas(aba, cabecalho.linha, cabecalho.colunas)
  const blocos: BlocoPedidos[] = []
  let pacote: { bloco: BlocoPedidos; total: number; data: string; comprador: string } | undefined

  aba.eachRow((linha) => {
    if (linha.number <= cabecalho.linha) return
    const texto = (coluna?: number) => coluna === undefined ? '' : textoCelula(linha.getCell(coluna).value)
    const valorUnidades = linha.getCell(colunas.unidades).value
    const unidades = typeof valorUnidades === 'number' ? valorUnidades : numeroUnidades(textoCelula(valorUnidades))
    const titulo = texto(colunas.titulo)
    const dados = texto(colunas.dados)
    const endereco = texto(colunas.endereco)
    const identificador = texto(colunas.venda).trim()
    const identificadorDeVenda = /\d/.test(identificador) && /^#?[\p{L}\p{N}./_-]+$/u.test(identificador)
    const temProduto = /\S/.test(titulo)
    const temCliente = /\S/.test(dados) || /\S/.test(endereco)
    if (rodape.test(normalizar(identificador)) || (!identificador && !temCliente && tituloRodape.test(normalizar(titulo)))) return

    const data = texto(colunas.data).trim()
    const nomeComprador = texto(colunas.comprador).trim()
    const nome = dados.trim() ? dados : nomeComprador
    const comprador = normalizar(nome)
    const estado = normalizar(texto(colunas.estado))
    const codigoEnvio = texto(colunas.envio).trim()
    const referencia: ReferenciaEnvio = {
      envio: /^\d+$/.test(codigoEnvio) ? codigoEnvio : '',
      venda: identificadorDeVenda ? identificador : '',
      comprador: nomeComprador || dados.trim(),
    }
    const total = Number(estado.match(estadoPacote)?.[1] ?? 0)
    if (total > 1 && !temProduto) {
      const bloco: BlocoPedidos = { resumo: { unidades: null, titulo, dados: nome, endereco }, referencia, pedidos: [] }
      blocos.push(bloco)
      pacote = { bloco, total, data, comprador }
      return
    }

    const temQuantidade = unidades != null && Number.isFinite(unidades)
    if (!identificadorDeVenda && !(temProduto && (temQuantidade || temCliente)) &&
        !(identificador && (temProduto || temCliente || temQuantidade))) return
    if (unidades != null && (!Number.isFinite(unidades) || unidades < 0)) {
      throw new ErroPlanilha('A quantidade na linha ' + linha.number + ' da aba Vendas BR é inválida.')
    }
    const pedido: PedidoRelatorio = {
      unidades, titulo, dados, endereco, referencia,
      item: nomeCurto(titulo),
      sku: texto(colunas.sku).trim(),
      cancelado: /\bcancelad[ao]\b|\bnao envie\b/.test(estado),
    }
    if (pacote && (temProduto || temQuantidade)) {
      const marca = normalizar(texto(colunas.pacote))
      const dataIgual = Boolean(data && data === pacote.data)
      const dataCompativel = !data || !pacote.data || dataIgual
      const compradorCompativel = !comprador || !pacote.comprador || comprador === pacote.comprador
      if ((marca === 'sim' || (!marca && dataIgual)) && dataCompativel && compradorCompativel) {
        pacote.bloco.pedidos.push(pedido)
        if (pacote.bloco.pedidos.length === pacote.total) pacote = undefined
        return
      }
    }
    pacote = undefined
    blocos.push({ referencia, pedidos: [pedido] })
  })
  if (!blocos.some(({ pedidos }) => pedidos.length > 0)) {
    throw new ErroPlanilha('O relatório não contém pedidos para formatar.')
  }
  return blocos
}

const bordaFina: Partial<Border> = { style: 'thin', color: { argb: 'FF000000' } }
const bordas: Partial<Borders> = { top: bordaFina, right: bordaFina, bottom: bordaFina, left: bordaFina }

function preenchimento(argb: string): Fill {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } }
}

function linhasVisiveis(texto: string, largura: number): number {
  const limite = largura * 7 - 5
  return texto.split(/\r\n|\r|\n/).reduce((total, trecho) => {
    let linhas = 1
    let usada = 0
    for (const palavra of trecho.split(/\s+/)) {
      const medida = larguraTexto(palavra)
      if (usada && usada + 3.5 + medida > limite) {
        linhas++
        usada = 0
      }
      if (medida > limite) linhas += Math.ceil(medida / limite) - 1
      usada = medida > limite ? medida % limite : usada + (usada ? 3.5 : 0) + medida
    }
    return total + linhas
  }, 0)
}

function gerarPlanilha(blocos: BlocoPedidos[]): ExcelJS.Workbook {
  const arquivo = new ExcelJS.Workbook()
  const aba = arquivo.addWorksheet('Planilha1', {
    properties: { defaultRowHeight: linhaBase },
    views: [{ zoomScale: 100 }],
  })
  larguras.forEach((largura, indice) => { aba.getColumn(indice + 1).width = largura })

  const topo = aba.addRow(cabecalhoSaida)
  topo.height = 60
  topo.eachCell((celula, coluna) => {
    celula.font = { name: 'Calibri', size: 9, bold: true }
    celula.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
    celula.fill = preenchimento(coresCabecalho[coluna - 1])
    celula.border = bordas
  })

  function adicionarLinha({ unidades, item, titulo, dados, endereco }: PedidoFormatado, ordem: number | null) {
    const linha = aba.addRow([ordem, unidades, item, titulo, dados, endereco])
    const linhas = Math.max(...[titulo, dados, endereco].map((texto, i) => linhasVisiveis(texto, larguras[i + 3])))
    linha.height = Math.min(409, Math.max(linhaBase, linhas > 1 ? linhas * 13 + 4 : linhaBase))
    for (let coluna = 1; coluna <= cabecalhoSaida.length; coluna++) {
      const celula = linha.getCell(coluna)
      celula.font = { name: 'Calibri', size: 10, bold: coluna <= 2 }
      celula.alignment = {
        horizontal: coluna <= 2 ? 'center' : 'left',
        vertical: 'middle',
        wrapText: coluna >= 4,
        shrinkToFit: coluna === 3,
      }
      celula.fill = preenchimento(coluna === 1 ? 'FFC0E6F5' : 'FFFFFFFF')
      celula.border = bordas
      if (coluna === 1) {
        celula.font = { ...celula.font, color: { argb: 'FFFF0000' } }
        celula.numFmt = '0'
      }
    }
  }

  let ordem = 0
  for (const { resumo, pedidos } of blocos) {
    if (resumo) {
      const total = pedidos.length
      const titulo = total === 0 ? resumo.titulo : '(' + total + (total === 1 ? ' pedido feito)' : ' pedidos feitos)')
      adicionarLinha({ ...resumo, item: '', titulo }, null)
    }
    for (const pedido of pedidos) {
      ordem++
      adicionarLinha(pedido, ordem)
    }
  }
  return arquivo
}

async function lerRelatorio(original: ArrayBuffer): Promise<BlocoPedidos[]> {
  const arquivo = new ExcelJS.Workbook()
  try {
    await arquivo.xlsx.load(original)
  } catch {
    throw new ErroPlanilha('Não foi possível ler o arquivo. Baixe novamente o relatório em formato .xlsx.')
  }
  const aba = arquivo.worksheets.find((planilha) => normalizar(planilha.name) === 'vendas br')
  if (!aba) throw new ErroPlanilha('A aba Vendas BR não foi encontrada no arquivo.')
  return lerBlocos(aba)
}

export async function lerEtiquetas(original: ArrayBuffer): Promise<EtiquetaEnvio[]> {
  const blocos = await lerRelatorio(original)
  const temEnvios = blocos.some(({ referencia, pedidos }) =>
    referencia.envio || pedidos.some((pedido) => pedido.referencia.envio))
  const etiquetas = new Map<string, EtiquetaEnvio>()

  for (const [indice, bloco] of blocos.entries()) {
    for (const pedido of bloco.pedidos) {
      const envio = pedido.referencia.envio || bloco.referencia.envio
      if (pedido.cancelado || (temEnvios && !envio)) continue

      const venda = bloco.referencia.venda
      const chave = envio ? 'envio:' + envio : venda ? 'venda:' + venda : 'linha:' + indice
      const comprador = bloco.referencia.comprador || pedido.referencia.comprador
      const etiqueta = etiquetas.get(chave) ?? { comprador, envio, produtos: [] }
      etiqueta.comprador ||= comprador
      etiqueta.produtos.push({ nome: pedido.titulo, unidades: pedido.unidades, sku: pedido.sku })
      etiquetas.set(chave, etiqueta)
    }
  }
  if (!etiquetas.size) throw new ErroPlanilha('O relatório não contém pedidos para imprimir etiquetas.')
  return [...etiquetas.values()]
}

export async function formatarPlanilha(original: ArrayBuffer) {
  const blocos = await lerRelatorio(original)
  const ordenados = blocos.map((bloco) => ({ bloco, item: bloco.pedidos.find(({ item }) => item)?.item ?? '' }))
    .sort((a, b) => Number(!a.item) - Number(!b.item) || alfabeto.compare(a.item, b.item))
    .map(({ bloco }) => bloco)
  try {
    const buffer = await gerarPlanilha(blocos).xlsx.writeBuffer()
    const bufferPorItens = await gerarPlanilha(ordenados).xlsx.writeBuffer()
    return {
      bytes: new Uint8Array(buffer).buffer,
      porItens: new Uint8Array(bufferPorItens).buffer,
      linhas: blocos.reduce((total, { pedidos }) => total + pedidos.length, 0),
    }
  } catch {
    throw new ErroPlanilha('Não foi possível gerar as planilhas formatadas. Tente novamente.')
  }
}

export async function renumerarPlanilhaPorItens(original: ArrayBuffer): Promise<ArrayBuffer> {
  try {
    const arquivo = new ExcelJS.Workbook()
    await arquivo.xlsx.load(original)
    const aba = arquivo.getWorksheet('Planilha1')
    if (!aba || aba.getCell('A1').value !== 'ORDEM') throw new Error('Cabeçalho ORDEM ausente.')

    let ordem = 0
    let alterada = false
    aba.eachRow((linha, numero) => {
      if (numero === 1) return
      const resumo = linha.getCell(2).value == null && !linha.getCell(3).value &&
        tituloPacote.test(textoCelula(linha.getCell(4).value))
      const numeroEsperado = resumo ? null : ++ordem
      const celula = linha.getCell(1)
      if (celula.value !== numeroEsperado) {
        celula.value = numeroEsperado
        alterada = true
      }
    })
    if (!alterada) return original

    return new Uint8Array(await arquivo.xlsx.writeBuffer()).buffer
  } catch {
    throw new ErroPlanilha('Não foi possível renumerar esta planilha. Gere as planilhas novamente a partir do arquivo original ML.')
  }
}
