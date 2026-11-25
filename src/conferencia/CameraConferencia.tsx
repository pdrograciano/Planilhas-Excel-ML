import { useEffect, useEffectEvent, useRef, useState } from 'react'
import jsQR from 'jsqr'

export function CameraConferencia({ aoLer }: { aoLer: (codigo: string) => void }) {
  const video = useRef<HTMLVideoElement>(null)
  const [estado, setEstado] = useState('Solicitando acesso à câmera…')
  const [erro, setErro] = useState(false)
  const ler = useEffectEvent(aoLer)

  useEffect(() => {
    const elemento = video.current!
    const canvas = document.createElement('canvas')
    const contexto = canvas.getContext('2d', { willReadFrequently: true })
    let stream: MediaStream | undefined
    let quadro = 0
    let encerrada = false
    let ultimoQuadro = 0
    let ultimaLeitura = 0
    let codigoVisivel = ''

    function examinar(instante: number) {
      if (encerrada || !contexto) return
      quadro = requestAnimationFrame(examinar)
      if (instante - ultimoQuadro < 130 || elemento.readyState < 2 || !elemento.videoWidth) return
      ultimoQuadro = instante
      const lado = Math.min(elemento.videoWidth, elemento.videoHeight)
      contexto.drawImage(elemento, (elemento.videoWidth - lado) / 2, (elemento.videoHeight - lado) / 2, lado, lado, 0, 0, 640, 640)
      const amostra = contexto.getImageData(0, 0, 640, 640)
      const codigo = jsQR(amostra.data, 640, 640, { inversionAttempts: 'dontInvert' })?.data
      if (!codigo) {
        if (instante - ultimaLeitura > 650) codigoVisivel = ''
        return
      }
      ultimaLeitura = instante
      if (codigo === codigoVisivel) return
      codigoVisivel = codigo
      ler(codigo)
    }

    async function abrir() {
      try {
        if (!contexto || !navigator.mediaDevices?.getUserMedia) throw new Error('camera indisponível')
        canvas.width = canvas.height = 640
        stream = await navigator.mediaDevices.getUserMedia({
          audio: false,
          video: { facingMode: { ideal: 'environment' }, width: { ideal: 1280 }, height: { ideal: 720 } },
        })
        if (encerrada) {
          stream.getTracks().forEach((track) => track.stop())
          return
        }
        elemento.srcObject = stream
        await elemento.play()
        if (encerrada) return
        setEstado('Centralize o QR code da etiqueta no quadro.')
        quadro = requestAnimationFrame(examinar)
      } catch (falha) {
        if (encerrada) return
        stream?.getTracks().forEach((track) => track.stop())
        setErro(true)
        setEstado(falha instanceof DOMException && falha.name === 'NotAllowedError'
          ? 'Permita o acesso à câmera nas configurações do navegador e tente de novo.'
          : 'Não foi possível abrir a câmera. Confira se ela está disponível e tente de novo.')
      }
    }

    void abrir()
    return () => {
      encerrada = true
      cancelAnimationFrame(quadro)
      stream?.getTracks().forEach((track) => track.stop())
      elemento.srcObject = null
    }
  }, [])

  return (
    <div className="camera-conferencia">
      <div className="visor-camera" hidden={erro}>
        <video ref={video} playsInline muted autoPlay aria-label="Prévia da câmera" />
        <span className="mira-camera" aria-hidden="true" />
      </div>
      <p className={erro ? 'notice notice-error' : 'instrucao-camera'} role={erro ? 'alert' : 'status'}>{estado}</p>
    </div>
  )
}
