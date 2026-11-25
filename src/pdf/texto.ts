import type { PDFFont } from 'pdf-lib'

export function textoImprimivel(texto: string, caracteres: Set<number>): string {
  return [...texto.normalize('NFC').replace(/\s+/g, ' ').trim()].map((letra) => {
    if (caracteres.has(letra.codePointAt(0)!)) return letra
    const equivalente = letra.normalize('NFKD').replace(/\p{M}/gu, '')
    return equivalente && [...equivalente].every((parte) => caracteres.has(parte.codePointAt(0)!)) ? equivalente : '?'
  }).join('')
}

export function quebrarLinhas(texto: string, fonte: PDFFont, tamanho: number, largura: number): string[] {
  const linhas: string[] = []
  let linha = ''
  const cabe = (valor: string) => fonte.widthOfTextAtSize(valor, tamanho) <= largura
  for (const palavra of texto.split(' ')) {
    const adicionada = linha ? linha + ' ' + palavra : palavra
    if (cabe(adicionada)) {
      linha = adicionada
      continue
    }
    if (cabe(palavra)) {
      if (linha) linhas.push(linha)
      linha = palavra
      continue
    }
    if (linha) linha += ' '
    for (const letra of palavra) {
      if (linha && !cabe(linha + letra)) {
        linhas.push(linha.trimEnd())
        linha = ''
      }
      linha += letra
    }
  }
  if (linha) linhas.push(linha)
  return linhas
}
