export type TipoBipe = 'identificado' | 'repetido' | 'desconhecido'

const padroes = {
  identificado: { frequencias: [880, 1175], duracao: .085, intervalo: .018 },
  desconhecido: { frequencias: [330], duracao: .25, intervalo: 0 },
  repetido: { frequencias: [660, 660], duracao: .06, intervalo: .075 },
} satisfies Record<TipoBipe, { frequencias: number[]; duracao: number; intervalo: number }>

export function prepararSom(contexto: AudioContext | null): AudioContext | null {
  try {
    const audio = contexto ?? new AudioContext()
    if (audio.state !== 'running') void audio.resume().catch(() => {})
    return audio
  } catch {
    return null
  }
}

export function tocarBipe(audio: AudioContext | null, tipo: TipoBipe): void {
  if (!audio || audio.state !== 'running') return
  const { frequencias, duracao, intervalo } = padroes[tipo]
  const agora = audio.currentTime
  for (const [indice, frequencia] of frequencias.entries()) {
    const inicio = agora + indice * (duracao + intervalo)
    const fim = inicio + duracao
    const oscilador = audio.createOscillator()
    const volume = audio.createGain()
    oscilador.type = tipo === 'desconhecido' ? 'triangle' : 'sine'
    oscilador.frequency.setValueAtTime(frequencia, inicio)
    volume.gain.setValueAtTime(0, inicio)
    volume.gain.linearRampToValueAtTime(.19, inicio + .008)
    volume.gain.linearRampToValueAtTime(0, fim)
    oscilador.connect(volume).connect(audio.destination)
    oscilador.onended = () => { oscilador.disconnect(); volume.disconnect() }
    oscilador.start(inicio)
    oscilador.stop(fim)
  }
}
