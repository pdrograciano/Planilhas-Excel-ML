import { createRoot } from 'react-dom/client'
import { App } from './App.tsx'
import { carregarHistorico } from './historico.ts'
import '@fontsource-variable/manrope'
import './styles.css'

createRoot(document.getElementById('root')!).render(<App historicoInicial={carregarHistorico()} />)
