export function normalizar(texto: string): string {
  return texto.normalize('NFD').replace(/\p{M}/gu, '').toLowerCase().trim().replace(/\s+/g, ' ')
}

const kit = /^(?:kit|combo|conjunto|jogo|pack|pacote|par)(?:\s+(?:de|com))?\b\s*/i
const kitNoTitulo = /\b(?:kit|combo|conjunto|jogo|pack|pacote|par de)\b(?:\s+(?:de|com))?\s*(\d+)?/
const timer = /^(?:(?:tomada)\s+)?(?:timer|temporizador)(?:es|s)?\b/
const medida = /^(\d+(?:[.,]\d+)?)(ml|mm|cm|kg|mg|l|m|g|w|v|h)$/i
const unidade = /^(?:ml|mm|cm|kg|mg|l|m|g|w|v|h)$/
const nomeUnidade: Record<string, string> = { l: 'L', w: 'W', v: 'V' }
const ignorar = new Set([
  'a', 'as', 'ao', 'com', 'da', 'das', 'de', 'do', 'dos', 'e', 'em', 'o', 'os', 'para', 'por', 'sem',
  'novo', 'nova', 'original', 'premium', 'promocao', 'oferta', 'especial', 'profissional',
  'exclusivo', 'incrivel', 'gratis', 'frete', 'imperdivel', 'lancamento', 'qualidade', 'unidade',
  'unidades', 'peca', 'pecas', 'un', 'pcs', 'kit', 'combo', 'conjunto', 'jogo', 'pack', 'pacote',
])
const abreviacoes: Record<string, string> = {
  multifuncional: 'Multifunc.', organizador: 'Organiz.', organizadora: 'Organiz.',
  reutilizavel: 'Reutiliz.', transparente: 'Transp.', retangular: 'Retang.',
}

export function larguraTexto(texto: string): number {
  return [...texto.normalize('NFC')].reduce((largura, caractere) => {
    if (/[ .,:;iílI!|'"’]/.test(caractere)) return largura + 3.5
    if (/[mwMW@%]/.test(caractere)) return largura + 10
    return largura + (/\p{Lu}/u.test(caractere) ? 8 : 6.8)
  }, 0)
}

function apresentar(palavra: string): string {
  if (/^(?:USB|LED|PVC|ABS|HDMI|RGB|[A-Z]\d+)$/i.test(palavra)) return palavra.toUpperCase()
  const texto = palavra.toLocaleLowerCase('pt-BR')
  return texto.charAt(0).toLocaleUpperCase('pt-BR') + texto.slice(1)
}

function compactar(palavras: string[], detalhe: string, prefixo: string): string {
  const cabem = (partes: string[]) => larguraTexto(partes.filter(Boolean).join(' ')) <= 119
  const nome: string[] = []
  for (const palavra of palavras.slice(0, 3)) {
    const curta = abreviacoes[normalizar(palavra)] ?? palavra
    if (cabem([prefixo, ...nome, palavra, detalhe])) nome.push(palavra)
    else if (cabem([prefixo, ...nome, curta, detalhe])) nome.push(curta)
    else break
  }
  if (!nome.length && palavras.length) {
    nome.push(abreviacoes[normalizar(palavras[0])] ?? palavras[0])
  }
  const partes = [prefixo, ...nome]
  if (detalhe && cabem([...partes, detalhe])) partes.push(detalhe)
  return partes.filter(Boolean).join(' ')
}

export function nomeCurto(titulo: string): string {
  const texto = titulo.trim()
  const chave = normalizar(texto)
  if (!chave) return ''
  if (/^(?:kit com )?4 mini potes\b/.test(chave)) return 'Mini Potes'
  if (/^trident chiclet/.test(chave)) return 'Trident'
  if (/^porta 4 ampolas\b/.test(chave)) return 'Porta Ampola Unitário'
  if (/^kit 2 porta ampolas?\b/.test(chave)) return 'Kit 2 Porta Ampola'
  if (/^kit 2 temporizador(?:es)?\b/.test(chave)) return 'Kit 02 Timer'

  const marcador = texto.match(kit)
  const indicador = chave.match(kitNoTitulo)
  const quantidadeAntes = !marcador ? chave.match(/^(\d+)\s+(?=(?:timer|temporizador))/)?.[1] : undefined
  const quantidadeDepois = chave.match(/\b(\d+)\s*(?:unidades?|pecas?|un|pcs)\b/)?.[1]
  const quantidade = quantidadeAntes ?? quantidadeDepois
  const temKit = Boolean(marcador || indicador || Number(quantidade) > 1)

  if (!temKit && timer.test(chave)) return 'Timer Unitário'

  let produto = marcador ? texto.slice(marcador[0].length) : texto
  const quantidadePrefixo = marcador ? produto.match(/^(\d+)\s+(?=\p{L})/u)?.[1] : quantidadeAntes
  if (quantidadePrefixo) {
    produto = produto.slice(quantidadePrefixo.length).trimStart().replace(/^(?:unidades?|peças?|un|pcs|de)\b\s*/i, '')
    produto = produto.replace(/^de\b\s*/i, '')
  }
  const ehPar = normalizar(marcador?.[0] ?? '').startsWith('par')
  const numeroKit = quantidadePrefixo ?? quantidadeDepois ?? indicador?.[1] ?? (ehPar ? '2' : '')
  const prefixo = temKit ? ['Kit', numeroKit].filter(Boolean).join(' ') : ''
  if (temKit && timer.test(normalizar(produto))) return [prefixo, 'Timer'].join(' ')

  const tokens = produto.match(/[\p{L}\p{N}]+(?:[-.,/][\p{L}\p{N}]+)*/gu) ?? []
  const palavras: string[] = []
  const vistos = new Set<string>()
  let detalhe = ''
  for (const [indice, palavra] of tokens.entries()) {
    const token = normalizar(palavra)
    const proximo = normalizar(tokens[indice + 1] ?? '')
    const justa = token.match(medida)
    const separada = /^\d+(?:[.,]\d+)?$/.test(token) && unidade.test(proximo)
    if (justa || separada) {
      const numero = justa?.[1] ?? token
      const simbolo = justa?.[2] ?? proximo
      detalhe ||= [numero, nomeUnidade[simbolo] ?? simbolo].join(' ')
    } else if (!unidade.test(token) && !ignorar.has(token) && !vistos.has(token) && /\p{L}/u.test(token)) {
      palavras.push(apresentar(palavra))
      vistos.add(token)
    }
  }
  return compactar(palavras, detalhe, prefixo) || apresentar(tokens[0] ?? texto)
}
