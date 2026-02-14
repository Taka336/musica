import './style.css'

type KeySpec = {
  keyboard: string
  note: string
  midi: number
  isBlack: boolean
}

const keySpecs: KeySpec[] = [
  { keyboard: 'a', note: 'C4', midi: 60, isBlack: false },
  { keyboard: 'w', note: 'C#4', midi: 61, isBlack: true },
  { keyboard: 's', note: 'D4', midi: 62, isBlack: false },
  { keyboard: 'e', note: 'D#4', midi: 63, isBlack: true },
  { keyboard: 'd', note: 'E4', midi: 64, isBlack: false },
  { keyboard: 'f', note: 'F4', midi: 65, isBlack: false },
  { keyboard: 't', note: 'F#4', midi: 66, isBlack: true },
  { keyboard: 'g', note: 'G4', midi: 67, isBlack: false },
  { keyboard: 'y', note: 'G#4', midi: 68, isBlack: true },
  { keyboard: 'h', note: 'A4', midi: 69, isBlack: false },
  { keyboard: 'u', note: 'A#4', midi: 70, isBlack: true },
  { keyboard: 'j', note: 'B4', midi: 71, isBlack: false },
  { keyboard: 'k', note: 'C5', midi: 72, isBlack: false }
]

const keyMap = new Map(keySpecs.map((spec) => [spec.keyboard, spec]))
const activeOscillators = new Map<string, OscillatorNode>()

let audioContext: AudioContext | null = null

const app = document.querySelector<HTMLDivElement>('#app')
if (!app) {
  throw new Error('App root not found')
}

app.innerHTML = `
  <main class="container">
    <h1>Keyboard Piano</h1>
    <p>PCキーボードで演奏: <strong>A W S E D F T G Y H U J K</strong></p>
    <div class="keyboard" aria-label="Piano keyboard"></div>
  </main>
`

const keyboardEl = document.querySelector<HTMLDivElement>('.keyboard')
if (!keyboardEl) {
  throw new Error('Keyboard container not found')
}

for (const spec of keySpecs) {
  const keyEl = document.createElement('button')
  keyEl.type = 'button'
  keyEl.className = `piano-key ${spec.isBlack ? 'black' : 'white'}`
  keyEl.dataset.key = spec.keyboard
  keyEl.dataset.note = spec.note
  keyEl.innerHTML = `<span>${spec.keyboard.toUpperCase()}</span><small>${spec.note}</small>`
  keyboardEl.appendChild(keyEl)
}

function midiToFrequency(midi: number): number {
  return 440 * Math.pow(2, (midi - 69) / 12)
}

function ensureAudioContext(): AudioContext {
  if (!audioContext) {
    audioContext = new AudioContext()
  }
  return audioContext
}

function highlightKey(key: string, isActive: boolean): void {
  const el = document.querySelector<HTMLElement>(`[data-key="${key}"]`)
  if (!el) {
    return
  }
  el.classList.toggle('active', isActive)
}

function noteOn(key: string): void {
  const spec = keyMap.get(key)
  if (!spec || activeOscillators.has(key)) {
    return
  }

  const ctx = ensureAudioContext()
  if (ctx.state === 'suspended') {
    void ctx.resume()
  }

  const oscillator = ctx.createOscillator()
  const gain = ctx.createGain()

  oscillator.type = 'triangle'
  oscillator.frequency.setValueAtTime(midiToFrequency(spec.midi), ctx.currentTime)

  gain.gain.setValueAtTime(0.0001, ctx.currentTime)
  gain.gain.exponentialRampToValueAtTime(0.2, ctx.currentTime + 0.01)

  oscillator.connect(gain)
  gain.connect(ctx.destination)

  oscillator.start()
  activeOscillators.set(key, oscillator)
  highlightKey(key, true)
}

function noteOff(key: string): void {
  const oscillator = activeOscillators.get(key)
  if (!oscillator || !audioContext) {
    return
  }

  const now = audioContext.currentTime
  const gain = oscillator.context.createGain()
  oscillator.disconnect()
  oscillator.connect(gain)
  gain.connect(audioContext.destination)
  gain.gain.setValueAtTime(0.2, now)
  gain.gain.exponentialRampToValueAtTime(0.0001, now + 0.05)

  oscillator.stop(now + 0.06)
  activeOscillators.delete(key)
  highlightKey(key, false)
}

window.addEventListener('keydown', (event) => {
  if (event.repeat) {
    return
  }

  const key = event.key.toLowerCase()
  if (!keyMap.has(key)) {
    return
  }

  event.preventDefault()
  noteOn(key)
})

window.addEventListener('keyup', (event) => {
  const key = event.key.toLowerCase()
  if (!keyMap.has(key)) {
    return
  }

  event.preventDefault()
  noteOff(key)
})
