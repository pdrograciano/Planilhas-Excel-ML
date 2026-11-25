import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import { test } from 'node:test'
import ExcelJS from 'exceljs'
import { PDFDocument } from 'pdf-lib'
import { chromium } from 'playwright'
import QRCode from 'qrcode'
import { preview } from 'vite'
import { exemploFicticio } from './exemplo-ficticio.mjs'

async function etiquetasFicticias(indices) {
  const numeros = [8, 9, 11, 99]
  const compradores = ['Cliente Fictício', 'Cliente da Silva', 'Cliente do Segundo Lote']
  const codigos = numeros.map((numero) => JSON.stringify({ id: String(70000000000 + numero), t: 'lm' }))
  const imagens = await Promise.all(codigos.map((codigo) => QRCode.toDataURL(codigo, { width: 260 })))
  const pdf = await PDFDocument.create()
  for (const indice of indices) {
    const comprador = compradores[indice]
    const pagina = pdf.addPage([283, 425])
    pagina.drawText('Envio: 700000', { x: 20, y: 365, size: 12 })
    pagina.drawText(String(numeros[indice]).padStart(5, '0'), { x: 102, y: 365, size: 14 })
    const qr = await pdf.embedPng(imagens[indice])
    pagina.drawImage(qr, { x: 178, y: 30, width: 80, height: 80 })
    if (!indice) pagina.drawText(comprador + ' (CLIENTE)', { x: 20, y: 95, size: 10 })
    pagina.drawText('Endereço: Rua Exemplo, 10', { x: 20, y: 76, size: 8 })
    const danfe = pdf.addPage([283, 425])
    danfe.drawText('DESTINATÁRIO: ' + comprador, { x: 20, y: 98, size: 10 })
    danfe.drawText('DANFE SIMPLIFICADO', { x: 20, y: 80, size: 10 })
  }
  return { imagens, buffer: Buffer.from(await pdf.save()) }
}

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

test('reúne PDFs, preserva os bipes ao adicionar etiquetas e usa o menu no celular', { timeout: 80_000 }, () => comNavegador(async (navegador, url) => {
  const pagina = await navegador.newPage({ viewport: { width: 375, height: 750 }, isMobile: true, hasTouch: true })
  await pagina.addInitScript(() => {
    window.frequenciasBipadas = []
    window.AudioContext = class extends window.AudioContext {
      createOscillator() {
        const oscilador = super.createOscillator()
        const definir = oscilador.frequency.setValueAtTime.bind(oscilador.frequency)
        oscilador.frequency.setValueAtTime = (frequencia, tempo) => {
          window.frequenciasBipadas.push(frequencia)
          return definir(frequencia, tempo)
        }
        return oscilador
      }
    }
    navigator.mediaDevices.getUserMedia = async () => {
      const canvas = document.createElement('canvas')
      canvas.width = canvas.height = 640
      const contexto = canvas.getContext('2d')
      const stream = canvas.captureStream(20)
      const faixa = stream.getVideoTracks()[0]
      window.faixaTeste = faixa
      let qr
      window.mostrarQr = async (imagem) => {
        const proximo = new Image()
        proximo.src = imagem
        await proximo.decode()
        qr = proximo
      }
      const quadros = setInterval(() => {
        if (faixa.readyState === 'ended') {
          clearInterval(quadros)
          return
        }
        contexto.fillStyle = 'white'
        contexto.fillRect(0, 0, 640, 640)
        if (qr) contexto.drawImage(qr, 190, 190)
      }, 55)
      return stream
    }
  })
  const { imagens, buffer } = await etiquetasFicticias([0])
  const manha = { name: 'manha.pdf', mimeType: 'application/pdf', buffer }
  const tarde = { name: 'tarde.pdf', mimeType: 'application/pdf', buffer: (await etiquetasFicticias([0, 1])).buffer }
  const extra = (await etiquetasFicticias([2])).buffer
  const invalido = { name: 'invalido.pdf', mimeType: 'application/pdf', buffer: Buffer.from('inválido') }
  const planilha = { name: 'pedidos-ficticios.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await exemploFicticio() }

  await pagina.goto(url)
  const aba = pagina.getByRole('tab', { name: 'Conferência de pedidos do dia' })
  const menu = pagina.getByRole('button', { name: 'Abrir menu' })
  assert.equal(await aba.count(), 0)
  await menu.click()
  await pagina.getByRole('button', { name: 'Fechar menu' }).press('Escape')
  assert.equal(await aba.count(), 0)
  await menu.click()
  await aba.click()
  assert.equal(await menu.getAttribute('aria-expanded'), 'false')
  const painel = pagina.getByRole('tabpanel', { name: 'Conferência de pedidos do dia' })
  const camera = painel.getByRole('button', { name: 'Abrir camera para validar pedidos' })
  assert.equal(await camera.count(), 0)
  const campoPdfs = painel.locator('input[accept$="application/pdf"]')
  await campoPdfs.setInputFiles([manha, invalido, tarde])
  await painel.getByText('2 envios encontrados').waitFor()
  await painel.getByText('2 PDFs adicionados').waitFor()
  await painel.getByRole('alert').getByText('Não foi possível ler invalido.pdf.').waitFor()
  assert.deepEqual(await painel.locator('.notice-warning').allTextContents(), [])
  const anexarPlanilha = painel.getByRole('button', { name: 'Carregar planilha de vendas' })
  await painel.getByText('Vincular planilha de vendas', { exact: true }).click()
  await anexarPlanilha.waitFor()
  await painel.locator('input[accept^=".xlsx"]').setInputFiles(planilha)
  await anexarPlanilha.waitFor({ state: 'hidden' })
  await camera.click()
  assert.equal(await painel.locator('.conferencia-arquivos').isVisible(), false)
  await painel.getByText('Centralize o QR code da etiqueta no quadro.').waitFor()

  await pagina.evaluate((imagem) => window.mostrarQr(imagem), imagens[0])
  await painel.getByText('Pedido identificado', { exact: true }).waitFor()
  await painel.getByText('Cliente Fictício', { exact: true }).waitFor()
  await painel.getByText('Kit 02 Timer', { exact: true }).waitFor()
  await painel.getByText('Qtd. 2', { exact: true }).waitFor()
  await painel.getByText('1 de 2 envios bipados', { exact: true }).waitFor()
  await pagina.evaluate((imagem) => window.mostrarQr(imagem), imagens[3])
  await painel.getByText('Pedido não identificado', { exact: true }).waitFor()
  await pagina.evaluate((imagem) => window.mostrarQr(imagem), imagens[0])
  await painel.getByText('Pedido já identificado', { exact: true }).waitFor()
  await painel.getByText('1 de 2 envios bipados', { exact: true }).waitFor()
  assert.deepEqual(await pagina.evaluate(() => window.frequenciasBipadas), [880, 1175, 330, 660, 660])

  await painel.getByRole('button', { name: 'Fechar câmera' }).click()
  const dataTransfer = await pagina.evaluateHandle((bytes) => {
    const dados = new DataTransfer()
    dados.items.add(new File([Uint8Array.from(bytes)], 'extra.pdf', { type: 'application/pdf' }))
    return dados
  }, [...extra])
  await painel.locator('.conferencia-dropzone').dispatchEvent('drop', { dataTransfer })
  await dataTransfer.dispose()
  await painel.getByText('1 de 3 envios bipados', { exact: true }).waitFor()
  assert.equal(await painel.getByText('Vincular planilha de vendas', { exact: true }).count(), 0)
  await painel.getByText('3 PDFs adicionados').click()
  await painel.getByRole('button', { name: 'Remover manha.pdf' }).click()
  await painel.getByText('2 PDFs adicionados').waitFor()
  await painel.getByText('1 de 3 envios bipados', { exact: true }).waitFor()

  const finalizar = painel.getByRole('button', { name: 'Finalizar' })
  const pdfParcial = await baixar(pagina, finalizar)
  assert.match(pdfParcial.nome, /^Conferencia_MIXBONIX_\d{2}-\d{2}-\d{4}_\d{2}-\d{2}-\d{2}\.pdf$/)
  const relatorio = await PDFDocument.load(pdfParcial.bytes)
  assert.equal(relatorio.getTitle(), 'Relatório de conferência - MIXBONIX ORDER')
  const { width, height } = relatorio.getPage(0).getSize()
  assert.ok(Math.abs(width * 25.4 / 72 - 210) < .01)
  assert.ok(Math.abs(height * 25.4 / 72 - 297) < .01)
  const resumo = painel.getByRole('region', { name: 'Resultado da conferência' })
  await resumo.getByText('Cliente da Silva').waitFor()
  assert.equal(await resumo.getByText('Pendente', { exact: true }).count(), 2)
  assert.equal(await pagina.evaluate(() => window.faixaTeste.readyState), 'ended')

  await camera.click()
  await painel.getByText('Centralize o QR code da etiqueta no quadro.').waitFor()
  await pagina.evaluate((imagem) => window.mostrarQr(imagem), imagens[1])
  await painel.getByText('2 de 3 envios bipados', { exact: true }).waitFor({ timeout: 8000 })
  await pagina.evaluate((imagem) => window.mostrarQr(imagem), imagens[2])
  await painel.getByText('3 de 3 envios bipados', { exact: true }).waitFor({ timeout: 8000 })
  const pdfCompleto = await baixar(pagina, finalizar)
  await resumo.getByText('Todos os envios do lote foram bipados.').waitFor()
  assert.equal(await resumo.getByText('Identificado', { exact: true }).count(), 3)
  assert.equal(pdfCompleto.bytes.equals(pdfParcial.bytes), false)
  const repetido = await baixar(pagina, resumo.getByRole('button', { name: 'Baixar relatório em PDF' }))
  assert.equal(repetido.nome, pdfCompleto.nome)
  assert.equal(repetido.bytes.equals(pdfCompleto.bytes), true)
  assert.equal(await pagina.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true)

  const abaPlanilhas = pagina.getByRole('tab', { name: 'Organização de pedidos e planilhas' })
  await menu.click()
  await abaPlanilhas.click()
  await pagina.locator('#painel-planilhas input[type="file"]').setInputFiles(planilha)
  await pagina.getByRole('button', { name: 'Gerar planilhas' }).click()
  await pagina.getByText('Planilhas prontas').waitFor()
  await menu.click()
  await aba.click()
  await resumo.getByText('Todos os envios do lote foram bipados.').waitFor()
  await pagina.reload()
  await menu.click()
  await aba.click()
  await campoPdfs.setInputFiles([manha, tarde])
  await painel.getByText('0 de 2 envios bipados', { exact: true }).waitFor()
  assert.equal(await anexarPlanilha.count(), 0)
}))

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
  assert.equal(await pagina.title(), 'MIXBONIX ORDER')
  const titulo = pagina.getByRole('heading', { name: 'MIXBONIX ORDER', level: 1 })
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
  await resultado.waitFor()
  const botoesDownload = await resultado.locator('.download-actions > button').evaluateAll((botoes) => botoes.map((botao) => botao.getBoundingClientRect().top))
  assert.equal(botoesDownload.length, 3)
  assert.equal(botoesDownload[0], botoesDownload[1])
  assert.ok(botoesDownload[0] < botoesDownload[2])
  const formato = { name: 'Arquivo Formatado', exact: true }
  const porItem = { name: 'Arquivo Ordenado por Itens', exact: true }
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
    assert.equal(await entrada.locator('.download-actions > button').count(), 3)
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
  assert.equal((await baixar(pagina, historico.getByRole('button', { name: 'Arquivo Ordenado por Itens', exact: true }))).bytes.length > 0, true)
  assert.equal(await historico.locator('.download-actions > button').count(), 3)
  await entrada.setInputFiles({ name: 'outro-inválido.xlsx', mimeType: 'text/plain', buffer: Buffer.from('inválido') })
  await gerar.click()
  await pagina.getByRole('alert').waitFor()
  assert.equal(await historico.getByRole('listitem').count(), 1)
  await pagina.reload()
  await pagina.getByText('Nenhuma planilha formatada').waitFor()
}))
