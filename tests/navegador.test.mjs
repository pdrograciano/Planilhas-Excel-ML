import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import { test } from 'node:test'
import ExcelJS from 'exceljs'
import { chromium } from 'playwright'
import { preview } from 'vite'
import { exemploFicticio } from './exemplo-ficticio.mjs'

const abrirHistorico = (pagina) => pagina.getByRole('region', { name: 'Histórico' })

async function baixar(pagina, botao) {
  const [download] = await Promise.all([pagina.waitForEvent('download'), botao.click()])
  return { nome: download.suggestedFilename(), bytes: await readFile(await download.path()) }
}

async function lerPlanilha(bytes) {
  const arquivo = new ExcelJS.Workbook()
  await arquivo.xlsx.load(bytes)
  return arquivo.getWorksheet('Planilha1')
}

async function comNavegador(executar) {
  const servidor = await preview({ preview: { host: '127.0.0.1', port: 0 } })
  let navegador
  try {
    const url = 'http://127.0.0.1:' + servidor.httpServer.address().port
    navegador = await chromium.launch({ channel: process.env.MIXBONIX_CANAL_BROWSER || 'chrome' })
    await executar(navegador, url)
  } finally {
    await navegador?.close()
    await new Promise((resolve, reject) => servidor.httpServer.close((erro) => erro ? reject(erro) : resolve()))
  }
}

test('formata, recarrega, baixa os arquivos e isola o histórico por guia', { timeout: 80_000 }, () => comNavegador(async (navegador, url) => {
  const contexto = await navegador.newContext({ viewport: { width: 1280, height: 900 }, timezoneId: 'America/Sao_Paulo' })
  const pagina = await contexto.newPage()
  const externas = new Set()
  pagina.on('request', (pedido) => {
    if (/^https?:/.test(pedido.url()) && new URL(pedido.url()).origin !== url) externas.add(pedido.url())
  })
  await pagina.clock.setFixedTime(new Date('2026-11-24T12:34:56-03:00'))
  await pagina.goto(url)
  await pagina.getByText('Nenhuma planilha formatada').waitFor()
  assert.equal(await pagina.title(), 'MIXBONIX GERATOR')
  const titulo = pagina.getByRole('heading', { name: 'MIXBONIX GERATOR', level: 1 })
  await titulo.waitFor()
  await pagina.getByText('Conversor de Excel de vendas do Mercado Livre para o modelo da MIXBONIX.', { exact: true }).waitFor()
  const rodape = pagina.getByRole('contentinfo')
  assert.match(await rodape.innerText(), /MIXBONIX\. Todos os direitos reservados\./)
  await rodape.getByText('Desenvolvido por Pedro Graciano.', { exact: true }).waitFor()
  assert.equal(await titulo.evaluate(async (elemento) => {
    await document.fonts.ready
    const estilo = getComputedStyle(elemento)
    return estilo.fontFamily.startsWith('Inter,') && estilo.fontWeight === '750' &&
      getComputedStyle(document.body).fontFamily.includes('Manrope Variable') &&
      Array.from(document.fonts).some((fonte) => fonte.family.includes('Manrope Variable') && fonte.status === 'loaded')
  }), true)
  const gerar = pagina.getByRole('button', { name: 'Gerar planilhas' })
  assert.equal(await gerar.isDisabled(), true)
  const tema = pagina.getByRole('button', { name: 'Modo escuro' })
  const logo = pagina.getByRole('img', { name: 'MIXBONIX', exact: true })
  assert.match(await logo.getAttribute('src'), /mixbonix-claro.*\.png$/)
  await logo.evaluate((imagem) => imagem.decode())
  assert.equal(await tema.getAttribute('aria-pressed'), 'false')
  assert.equal(await tema.locator('.lucide-sun').count(), 1)
  await pagina.keyboard.press('Tab')
  assert.equal(await tema.evaluate((botao) => botao === document.activeElement), true)
  await pagina.keyboard.press('Enter')
  await tema.locator('.lucide-moon').waitFor()
  assert.equal(await tema.getAttribute('aria-pressed'), 'true')
  assert.equal(await pagina.locator('meta[name="theme-color"]').getAttribute('content'), '#282a36')
  assert.match(await logo.getAttribute('src'), /mixbonix-escuro.*\.png$/)
  await logo.evaluate((imagem) => imagem.decode())
  await pagina.reload()
  await tema.locator('.lucide-moon').waitFor()
  assert.equal(await pagina.locator('html').getAttribute('data-tema'), 'escuro')
  assert.match(await logo.getAttribute('src'), /mixbonix-escuro.*\.png$/)
  await pagina.keyboard.press('Tab')
  await pagina.keyboard.press('Enter')
  await tema.locator('.lucide-sun').waitFor()
  assert.match(await logo.getAttribute('src'), /mixbonix-claro.*\.png$/)
  await pagina.keyboard.press('Tab')
  assert.equal(await pagina.evaluate(() => document.activeElement.textContent.trim()), 'Selecionar arquivo')
  assert.equal(await pagina.locator('button:focus-visible').count(), 1)

  const original = await exemploFicticio()
  const arquivo = { name: 'pedidos-teste.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: original }
  const selecionar = pagina.getByRole('button', { name: 'Selecionar arquivo' })
  const [escolha] = await Promise.all([pagina.waitForEvent('filechooser'), selecionar.press('Enter')])
  await escolha.setFiles(arquivo)
  let liberarWorker
  const pausa = new Promise((resolver) => { liberarWorker = resolver })
  await pagina.route('**/assets/formatador.worker-*.js', async (rota) => {
    await pausa
    await rota.continue()
  })
  await gerar.click()
  try {
    const ocupado = pagina.getByRole('button', { name: 'Formatando planilhas…' })
    await ocupado.waitFor()
    assert.equal(await ocupado.isDisabled(), true)
  } finally {
    liberarWorker()
  }
  const resultado = pagina.locator('.result')
  const formato = { name: 'Arquivo Formatado', exact: true }
  const porItem = { name: 'Arquivo Formatado por Itens', exact: true }
  await pagina.getByText('6 linhas de pedidos.').waitFor()
  const primeira = await baixar(pagina, resultado.getByRole('button', formato))
  const primeiraPorItem = await baixar(pagina, resultado.getByRole('button', porItem))
  assert.equal(primeira.nome, '24-11-2026_12-34-56_Formatado_MercadoLivre_mixbonixgerator.xlsx')
  assert.equal(primeiraPorItem.nome, '24-11-2026_12-34-56_PorItens_MercadoLivre_mixbonixgerator.xlsx')
  assert.equal(primeiraPorItem.bytes.equals(primeira.bytes), false)
  const planilhaFormatada = await lerPlanilha(primeira.bytes)
  const planilhaPorItens = await lerPlanilha(primeiraPorItem.bytes)
  assert.deepEqual(planilhaPorItens.getColumn('A').values.slice(2), [1, 2, 3, 4, 5, 6])
  assert.deepEqual(planilhaPorItens.getColumn('A').values, planilhaFormatada.getColumn('A').values)

  await pagina.clock.setFixedTime(new Date('2026-11-24T12:35:01-03:00'))
  await gerar.click()
  await pagina.getByText('6 linhas de pedidos.').waitFor()
  const segunda = await baixar(pagina, resultado.getByRole('button', formato))
  const segundaPorItem = await baixar(pagina, resultado.getByRole('button', porItem))
  const historico = abrirHistorico(pagina)
  const entradas = historico.getByRole('listitem')
  assert.equal(await entradas.count(), 2)
  await pagina.reload()
  await entradas.nth(1).waitFor()
  assert.equal(await entradas.count(), 2)
  for (const [indice, [esperado, esperadoPorItem]] of [[segunda, segundaPorItem], [primeira, primeiraPorItem]].entries()) {
    const entrada = entradas.nth(indice)
    assert.match(await entrada.innerText(), /pedidos-teste\.xlsx.*\n.*6 linhas/s)
    const retornoOriginal = await baixar(pagina, entrada.getByRole('button', { name: 'Arquivo original ML', exact: true }))
    const retornoFormatada = await baixar(pagina, entrada.getByRole('button', formato))
    const retornoPorItem = await baixar(pagina, entrada.getByRole('button', porItem))
    assert.equal(retornoOriginal.nome, arquivo.name)
    assert.equal(retornoOriginal.bytes.equals(original), true)
    assert.equal(retornoFormatada.nome, esperado.nome)
    assert.equal(retornoFormatada.bytes.equals(esperado.bytes), true)
    assert.equal(retornoPorItem.nome, esperadoPorItem.nome)
    assert.equal(retornoPorItem.bytes.equals(esperadoPorItem.bytes), true)
  }
  assert.equal(externas.size, 0)

  const independente = await contexto.newPage()
  await independente.goto(url)
  await independente.getByText('Nenhuma planilha formatada').waitFor()
  const [aberta] = await Promise.all([pagina.waitForEvent('popup'), pagina.evaluate(() => { window.open(location.href) })])
  await aberta.getByText('Nenhuma planilha formatada').waitFor()
  await aberta.close()
  await entradas.first().getByRole('button', { name: 'Apagar do histórico pedidos-teste.xlsx' }).click()
  await entradas.nth(1).waitFor({ state: 'detached' })
  await pagina.reload()
  await entradas.first().waitFor()
  assert.equal(await entradas.count(), 1)
  const mantida = await baixar(pagina, entradas.first().getByRole('button', formato))
  assert.equal(mantida.nome, primeira.nome)
  const limpar = historico.getByRole('button', { name: 'Limpar histórico' })
  pagina.once('dialog', (dialogo) => dialogo.dismiss())
  await limpar.click()
  assert.equal(await entradas.count(), 1)
  pagina.once('dialog', (dialogo) => dialogo.accept())
  await limpar.click()
  await pagina.getByText('Nenhuma planilha formatada').waitFor()
  await pagina.reload()
  await pagina.getByText('Nenhuma planilha formatada').waitFor()
  await pagina.evaluate(({ original, formatada }) => {
    sessionStorage.setItem('mixbonix:pedidos:v1:antiga', JSON.stringify({
      id: 'antiga',
      nome: 'anterior.xlsx',
      nomeGerado: '23-11-2026_17-40-00_MercadoLivre_mixbonixgerator.xlsx',
      data: Date.now(),
      linhas: 6,
      original,
      formatada,
    }))
  }, { original: original.toString('base64'), formatada: primeira.bytes.toString('base64') })
  await pagina.reload()
  const anterior = entradas.filter({ hasText: 'anterior.xlsx' })
  await anterior.waitFor()
  const baixada = await baixar(pagina, anterior.getByRole('button', formato))
  assert.equal(baixada.nome, '23-11-2026_17-40-00_Formatado_MercadoLivre_mixbonixgerator.xlsx')
  assert.equal(baixada.bytes.equals(primeira.bytes), true)
  assert.equal(await anterior.getByRole('button', porItem).count(), 0)

  const planilhaDesordenada = await lerPlanilha(primeiraPorItem.bytes)
  for (const [indice, ordem] of [4, 6, 2, 1, 5, 3].entries()) {
    planilhaDesordenada.getCell(indice + 2, 1).value = ordem
  }
  const copiaAntiga = Buffer.from(await planilhaDesordenada.workbook.xlsx.writeBuffer())
  const chaveAntiga = 'mixbonix:pedidos:v1:antiga'
  await pagina.evaluate(({ chave, bytes }) => {
    const registro = JSON.parse(sessionStorage.getItem(chave))
    sessionStorage.setItem(chave, JSON.stringify({ ...registro, porItens: bytes }))
  }, { chave: chaveAntiga, bytes: copiaAntiga.toString('base64') })
  await pagina.reload()
  const downloadCorrigido = await baixar(pagina, anterior.getByRole('button', porItem))
  const planilhaCorrigida = await lerPlanilha(downloadCorrigido.bytes)
  assert.equal(downloadCorrigido.nome, '23-11-2026_17-40-00_PorItens_MercadoLivre_mixbonixgerator.xlsx')
  assert.deepEqual(planilhaCorrigida.getSheetValues(), planilhaPorItens.getSheetValues())
  assert.deepEqual(planilhaCorrigida.getCell('A2').style, planilhaPorItens.getCell('A2').style)
  await pagina.waitForFunction((chave) => JSON.parse(sessionStorage.getItem(chave)).porItensNumerado, chaveAntiga)
  await pagina.reload()
  const recuperado = await baixar(pagina, anterior.getByRole('button', porItem))
  assert.equal(recuperado.bytes.equals(downloadCorrigido.bytes), true)
  const formatadoDaEntradaAntiga = await baixar(pagina, anterior.getByRole('button', formato))
  assert.equal(formatadoDaEntradaAntiga.bytes.equals(primeira.bytes), true)
  await pagina.close()
  const nova = await contexto.newPage()
  await nova.goto(url)
  await nova.getByText('Nenhuma planilha formatada').waitFor()
}))

test('funciona no celular e conserva os downloads quando o armazenamento falha', { timeout: 80_000 }, () => comNavegador(async (navegador, url) => {
  const pagina = await navegador.newPage({ viewport: { width: 360, height: 800 }, hasTouch: true, isMobile: true })
  await pagina.addInitScript(() => {
    const setItem = Storage.prototype.setItem
    Storage.prototype.setItem = function (chave, valor) {
      if (chave.startsWith('mixbonix:pedidos:')) throw new DOMException('Quota de teste', 'QuotaExceededError')
      return setItem.call(this, chave, valor)
    }
  })
  await pagina.goto(url)
  await pagina.getByRole('button', { name: 'Modo escuro' }).tap()
  assert.equal(await pagina.getByText('Arraste e solte a planilha aqui', { exact: true }).isVisible(), false)
  const entrada = pagina.locator('input[type=file]')
  assert.match(await entrada.getAttribute('accept'), /spreadsheetml\.sheet/)
  await entrada.setInputFiles({ name: 'arquivo.csv', mimeType: 'text/csv', buffer: Buffer.from('inválido') })
  await pagina.getByRole('alert').filter({ hasText: 'formato .xlsx' }).waitFor()

  const gerar = pagina.getByRole('button', { name: 'Gerar planilhas' })
  await entrada.setInputFiles({ name: 'inválido.xlsx', mimeType: 'text/plain', buffer: Buffer.from('inválido') })
  await gerar.click()
  await pagina.getByRole('alert').filter({ hasText: 'Não foi possível ler o arquivo' }).waitFor()

  await entrada.setInputFiles({ name: 'sem-pedidos.xlsx', mimeType: 'application/octet-stream', buffer: await exemploFicticio({ semPedidos: true }) })
  await gerar.click()
  await pagina.getByRole('alert').filter({ hasText: 'não contém pedidos' }).waitFor()

  const original = await exemploFicticio()
  const nome = 'Vendas-Mercado-Livre-relatorio-ficticio-com-nome-comprido-2026-11-23.xlsx'
  const [seletor] = await Promise.all([
    pagina.waitForEvent('filechooser'),
    pagina.getByRole('button', { name: 'Trocar arquivo', exact: true }).tap(),
  ])
  await seletor.setFiles({ name: nome, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: original })
  await gerar.tap()
  await pagina.getByText('O navegador não conseguiu guardar estes arquivos.', { exact: false }).waitFor()
  const historico = abrirHistorico(pagina)
  assert.equal(await historico.getByRole('listitem').count(), 1)
  assert.equal(await pagina.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth), true)
  assert.equal((await baixar(pagina, historico.getByRole('button', { name: 'Arquivo original ML', exact: true }))).bytes.equals(original), true)
  assert.equal((await baixar(pagina, historico.getByRole('button', { name: 'Arquivo Formatado', exact: true }))).bytes.length > 0, true)
  assert.equal((await baixar(pagina, historico.getByRole('button', { name: 'Arquivo Formatado por Itens', exact: true }))).bytes.length > 0, true)
  await entrada.setInputFiles({ name: 'outro-inválido.xlsx', mimeType: 'text/plain', buffer: Buffer.from('inválido') })
  await gerar.click()
  await pagina.getByRole('alert').waitFor()
  assert.equal(await historico.getByRole('listitem').count(), 1)
  await pagina.reload()
  await pagina.getByText('Nenhuma planilha formatada').waitFor()
}))
