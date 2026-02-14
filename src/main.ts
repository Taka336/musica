import './style.css'

type KeySpec = {
  keyboard: string
  note: string
  midi: number
  isBlack: boolean
}

const toneOptions = [
  { value: 'piano', label: 'Piano (リアル寄り)' },
  { value: 'sine', label: 'Sine (柔らかい)' },
  { value: 'triangle', label: 'Triangle (標準)' },
  { value: 'square', label: 'Square (硬め)' },
  { value: 'sawtooth', label: 'Sawtooth (明るめ)' }
] as const

type ToneType = (typeof toneOptions)[number]['value']
type WaveToneType = Exclude<ToneType, 'piano'>
type ActiveVoice = {
  stop: (when: number) => void
}
type ChordStep = {
  midiNotes: number[]
}
type PlayMode = 'chord' | 'arpeggio'
type RhythmMode = 'straight' | 'eighth' | 'syncopated'
type ArticulationMode = 'normal' | 'staccato' | 'legato'
type ParsedProgression = {
  steps: ChordStep[]
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
const midiToKeyboardKey = new Map(keySpecs.map((spec) => [spec.midi, spec.keyboard]))
const activeVoices = new Map<string, ActiveVoice>()
const activeHighlights = new Map<string, number>()

let audioContext: AudioContext | null = null
let noiseBuffer: AudioBuffer | null = null
let selectedTone: ToneType = 'piano'
let selectedPlayMode: PlayMode = 'chord'
let selectedRhythmMode: RhythmMode = 'straight'
let selectedArticulation: ArticulationMode = 'normal'
let tempoBpm = 100
let isPlayingProgression = false
let progressionTimeouts: number[] = []
let progressionVoices: ActiveVoice[] = []
const defaultProgressionInput = 'A – E – F#m – C#m – D – A – D – E'

const app = document.querySelector<HTMLDivElement>('#app')
if (!app) {
  throw new Error('App root not found')
}

app.innerHTML = `
  <main class="container">
    <h1>Keyboard Piano</h1>
    <p>PCキーボードで演奏: <strong>A W S E D F T G Y H U J K</strong></p>
    <label class="sound-control">
      音の種類
      <select id="tone-select" aria-label="音の種類">
        ${toneOptions
          .map(
            (option) =>
              `<option value="${option.value}" ${option.value === selectedTone ? 'selected' : ''}>${option.label}</option>`
          )
          .join('')}
      </select>
    </label>
    <label class="sound-control">
      再生モード
      <select id="play-mode-select" aria-label="再生モード">
        <option value="chord" selected>コード</option>
        <option value="arpeggio">アルペジオ</option>
      </select>
    </label>
    <label class="sound-control">
      リズム
      <select id="rhythm-select" aria-label="リズム">
        <option value="straight" selected>ストレート</option>
        <option value="eighth">8ビート</option>
        <option value="syncopated">シンコペーション</option>
      </select>
    </label>
    <label class="sound-control">
      演奏法
      <select id="articulation-select" aria-label="演奏法">
        <option value="normal" selected>ノーマル</option>
        <option value="staccato">スタッカート</option>
        <option value="legato">レガート</option>
      </select>
    </label>
    <label class="sound-control">
      テンポ (BPM)
      <input id="tempo-input" class="tempo-input" type="number" min="40" max="220" step="1" value="100" aria-label="テンポ" />
    </label>
    <label class="progression-control">
      コード進行入力
      <input id="progression-input" type="text" value="${defaultProgressionInput}" aria-label="コード進行入力" />
    </label>
    <p id="progression-error" class="progression-error" aria-live="polite"></p>
    <button id="play-progression" class="play-btn" type="button" aria-label="コード進行を再生">
      コード進行を再生
    </button>
    <div class="keyboard" aria-label="Piano keyboard"></div>
  </main>
`

const toneSelectEl = document.querySelector<HTMLSelectElement>('#tone-select')
if (toneSelectEl) {
  toneSelectEl.addEventListener('change', (event) => {
    const value = (event.target as HTMLSelectElement).value
    const isValid = toneOptions.some((option) => option.value === value)
    if (isValid) {
      selectedTone = value as ToneType
    }
  })
}

const playModeSelectEl = document.querySelector<HTMLSelectElement>('#play-mode-select')
if (playModeSelectEl) {
  playModeSelectEl.addEventListener('change', (event) => {
    const value = (event.target as HTMLSelectElement).value
    if (value === 'chord' || value === 'arpeggio') {
      selectedPlayMode = value
    }
  })
}

const rhythmSelectEl = document.querySelector<HTMLSelectElement>('#rhythm-select')
if (rhythmSelectEl) {
  rhythmSelectEl.addEventListener('change', (event) => {
    const value = (event.target as HTMLSelectElement).value
    if (value === 'straight' || value === 'eighth' || value === 'syncopated') {
      selectedRhythmMode = value
    }
  })
}

const articulationSelectEl = document.querySelector<HTMLSelectElement>('#articulation-select')
if (articulationSelectEl) {
  articulationSelectEl.addEventListener('change', (event) => {
    const value = (event.target as HTMLSelectElement).value
    if (value === 'normal' || value === 'staccato' || value === 'legato') {
      selectedArticulation = value
    }
  })
}

const tempoInputEl = document.querySelector<HTMLInputElement>('#tempo-input')
if (tempoInputEl) {
  tempoInputEl.addEventListener('change', (event) => {
    const raw = Number((event.target as HTMLInputElement).value)
    if (!Number.isFinite(raw)) {
      return
    }
    tempoBpm = Math.max(40, Math.min(220, Math.round(raw)))
    tempoInputEl.value = String(tempoBpm)
  })
}

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

function parseChordSymbolToMidiNotes(symbol: string): number[] | null {
  const match = symbol.trim().match(/^([A-Ga-g])([#b]?)(m|minor|maj)?$/)
  if (!match) {
    return null
  }

  const [, letterRaw, accidentalRaw, qualityRaw] = match
  const letter = letterRaw.toUpperCase()
  const accidental = accidentalRaw ?? ''
  const quality = (qualityRaw ?? '').toLowerCase()

  const baseNoteClass: Record<string, number> = {
    C: 0,
    D: 2,
    E: 4,
    F: 5,
    G: 7,
    A: 9,
    B: 11
  }

  let noteClass = baseNoteClass[letter]
  if (accidental === '#') {
    noteClass += 1
  } else if (accidental === 'b') {
    noteClass -= 1
  }
  noteClass = (noteClass + 12) % 12

  const isMinor = quality === 'm' || quality === 'minor'
  const intervals = isMinor ? [0, 3, 7] : [0, 4, 7]
  const rootMidi = 60 + noteClass
  return intervals.map((interval) => rootMidi + interval)
}

function parseProgression(input: string): ParsedProgression | null {
  const symbols = input
    .split(/[\-\u2010-\u2015,|]+/)
    .map((s) => s.trim())
    .filter((s) => s.length > 0)

  if (symbols.length === 0) {
    return null
  }

  const steps: ChordStep[] = []
  for (const symbol of symbols) {
    const midiNotes = parseChordSymbolToMidiNotes(symbol)
    if (!midiNotes) {
      return null
    }
    steps.push({ midiNotes })
  }

  return { steps }
}

function ensureAudioContext(): AudioContext {
  if (!audioContext) {
    audioContext = new AudioContext()
  }
  return audioContext
}

function getNoiseBuffer(ctx: AudioContext): AudioBuffer {
  if (noiseBuffer && noiseBuffer.sampleRate === ctx.sampleRate) {
    return noiseBuffer
  }

  const length = Math.floor(ctx.sampleRate * 0.08)
  const buffer = ctx.createBuffer(1, length, ctx.sampleRate)
  const channel = buffer.getChannelData(0)
  for (let i = 0; i < channel.length; i += 1) {
    channel[i] = Math.random() * 2 - 1
  }
  noiseBuffer = buffer
  return buffer
}

function highlightKey(key: string, isActive: boolean): void {
  const el = document.querySelector<HTMLElement>(`[data-key="${key}"]`)
  if (!el) {
    return
  }
  el.classList.toggle('active', isActive)
}

function activateKeyHighlight(key: string): void {
  const count = activeHighlights.get(key) ?? 0
  activeHighlights.set(key, count + 1)
  if (count === 0) {
    highlightKey(key, true)
  }
}

function deactivateKeyHighlight(key: string): void {
  const count = activeHighlights.get(key) ?? 0
  if (count <= 1) {
    activeHighlights.delete(key)
    highlightKey(key, false)
    return
  }
  activeHighlights.set(key, count - 1)
}

function createWaveVoice(ctx: AudioContext, frequency: number, tone: WaveToneType): ActiveVoice {
  const oscillator = ctx.createOscillator()
  const gain = ctx.createGain()

  oscillator.type = tone
  oscillator.frequency.setValueAtTime(frequency, ctx.currentTime)
  gain.gain.setValueAtTime(0.0001, ctx.currentTime)
  gain.gain.exponentialRampToValueAtTime(0.2, ctx.currentTime + 0.01)

  oscillator.connect(gain)
  gain.connect(ctx.destination)
  oscillator.start()

  return {
    stop: (when: number) => {
      gain.gain.cancelScheduledValues(when)
      gain.gain.setValueAtTime(Math.max(gain.gain.value, 0.0001), when)
      gain.gain.exponentialRampToValueAtTime(0.0001, when + 0.06)
      oscillator.stop(when + 0.07)
    }
  }
}

function createPianoVoice(ctx: AudioContext, frequency: number): ActiveVoice {
  const output = ctx.createGain()
  output.gain.setValueAtTime(0.0001, ctx.currentTime)
  output.connect(ctx.destination)

  const bodyFilter = ctx.createBiquadFilter()
  bodyFilter.type = 'lowpass'
  bodyFilter.frequency.setValueAtTime(4200, ctx.currentTime)
  bodyFilter.Q.setValueAtTime(0.9, ctx.currentTime)
  bodyFilter.connect(output)

  const partials: Array<{ multiple: number; gain: number; type: OscillatorType; detune: number }> = [
    { multiple: 1, gain: 0.8, type: 'triangle', detune: 0 },
    { multiple: 2, gain: 0.22, type: 'sine', detune: -3 },
    { multiple: 3, gain: 0.14, type: 'sine', detune: 2.5 }
  ]

  const oscillators: OscillatorNode[] = []
  for (const partial of partials) {
    const osc = ctx.createOscillator()
    const partialGain = ctx.createGain()
    osc.type = partial.type
    osc.frequency.setValueAtTime(frequency * partial.multiple, ctx.currentTime)
    osc.detune.setValueAtTime(partial.detune, ctx.currentTime)
    partialGain.gain.setValueAtTime(partial.gain, ctx.currentTime)
    osc.connect(partialGain)
    partialGain.connect(bodyFilter)
    osc.start()
    oscillators.push(osc)
  }

  const hammerNoise = ctx.createBufferSource()
  hammerNoise.buffer = getNoiseBuffer(ctx)
  const hammerFilter = ctx.createBiquadFilter()
  hammerFilter.type = 'bandpass'
  hammerFilter.frequency.setValueAtTime(2200, ctx.currentTime)
  hammerFilter.Q.setValueAtTime(0.7, ctx.currentTime)
  const hammerGain = ctx.createGain()
  hammerGain.gain.setValueAtTime(0.08, ctx.currentTime)
  hammerGain.gain.exponentialRampToValueAtTime(0.0001, ctx.currentTime + 0.03)
  hammerNoise.connect(hammerFilter)
  hammerFilter.connect(hammerGain)
  hammerGain.connect(bodyFilter)
  hammerNoise.start()

  const now = ctx.currentTime
  output.gain.exponentialRampToValueAtTime(0.24, now + 0.003)
  output.gain.exponentialRampToValueAtTime(0.11, now + 0.17)
  output.gain.exponentialRampToValueAtTime(0.0001, now + 2.6)

  return {
    stop: (when: number) => {
      output.gain.cancelScheduledValues(when)
      output.gain.setValueAtTime(Math.max(output.gain.value, 0.0001), when)
      output.gain.exponentialRampToValueAtTime(0.0001, when + 0.18)
      for (const osc of oscillators) {
        osc.stop(when + 0.2)
      }
    }
  }
}

function startVoice(frequency: number): ActiveVoice {
  const ctx = ensureAudioContext()
  if (ctx.state === 'suspended') {
    void ctx.resume()
  }

  return selectedTone === 'piano'
    ? createPianoVoice(ctx, frequency)
    : createWaveVoice(ctx, frequency, selectedTone)
}

function noteOn(key: string): void {
  const spec = keyMap.get(key)
  if (!spec || activeVoices.has(key)) {
    return
  }

  const frequency = midiToFrequency(spec.midi)
  const voice = startVoice(frequency)
  activeVoices.set(key, voice)
  activateKeyHighlight(key)
}

function noteOff(key: string): void {
  const voice = activeVoices.get(key)
  if (!voice || !audioContext) {
    return
  }

  const now = audioContext.currentTime
  voice.stop(now)
  activeVoices.delete(key)
  deactivateKeyHighlight(key)
}

function stopProgression(): void {
  for (const timeoutId of progressionTimeouts) {
    window.clearTimeout(timeoutId)
  }
  progressionTimeouts = []

  if (audioContext) {
    const now = audioContext.currentTime
    for (const voice of progressionVoices) {
      voice.stop(now)
    }
  }
  progressionVoices = []
  isPlayingProgression = false
}

function getRhythmPattern(): number[] {
  if (selectedRhythmMode === 'eighth') {
    return [0.5, 0.5]
  }
  if (selectedRhythmMode === 'syncopated') {
    return [0.75, 0.25]
  }
  return [1]
}

function getArticulationGate(): number {
  if (selectedArticulation === 'staccato') {
    return 0.45
  }
  if (selectedArticulation === 'legato') {
    return 0.95
  }
  return 0.75
}

function playProgression(): void {
  const progressionInputEl = document.querySelector<HTMLInputElement>('#progression-input')
  const errorEl = document.querySelector<HTMLElement>('#progression-error')
  const playBtnEl = document.querySelector<HTMLButtonElement>('#play-progression')
  if (!progressionInputEl || !errorEl || !playBtnEl) {
    return
  }

  const parsed = parseProgression(progressionInputEl.value)
  if (!parsed) {
    errorEl.textContent = 'コード進行の形式が不正です。例: A - E - F#m - C#m - D - A - D - E'
    return
  }
  if (!Number.isFinite(tempoBpm) || tempoBpm < 40 || tempoBpm > 220) {
    errorEl.textContent = 'テンポは 40〜220 の範囲で指定してください。'
    return
  }
  errorEl.textContent = ''

  stopProgression()
  isPlayingProgression = true

  playBtnEl.disabled = true
  playBtnEl.textContent = selectedPlayMode === 'arpeggio' ? '再生中... (アルペジオ)' : '再生中... (コード)'

  const beatMs = 60000 / tempoBpm
  const rhythmPattern = getRhythmPattern()
  const articulationGate = getArticulationGate()
  let elapsedMs = 0

  for (const step of parsed.steps) {
    for (const subdivision of rhythmPattern) {
      const slotMs = beatMs * subdivision
      const noteMs = Math.max(80, slotMs * articulationGate)
      const startAtMs = elapsedMs

      if (selectedPlayMode === 'chord') {
        const startTimeout = window.setTimeout(() => {
          const keysToHighlight = step.midiNotes
            .map((midi) => midiToKeyboardKey.get(midi))
            .filter((key): key is string => typeof key === 'string')
          for (const key of keysToHighlight) {
            activateKeyHighlight(key)
          }

          const startedVoices = step.midiNotes.map((midi) => startVoice(midiToFrequency(midi)))
          progressionVoices.push(...startedVoices)

          const stopTimeout = window.setTimeout(() => {
            if (!audioContext) {
              return
            }
            const now = audioContext.currentTime
            for (const voice of startedVoices) {
              voice.stop(now)
            }
            for (const key of keysToHighlight) {
              deactivateKeyHighlight(key)
            }
          }, noteMs)
          progressionTimeouts.push(stopTimeout)
        }, startAtMs)
        progressionTimeouts.push(startTimeout)
      } else {
        for (const [noteIndex, midi] of step.midiNotes.entries()) {
          const noteStartMs = startAtMs + (slotMs / step.midiNotes.length) * noteIndex
          const startTimeout = window.setTimeout(() => {
            const maybeKey = midiToKeyboardKey.get(midi)
            if (maybeKey) {
              activateKeyHighlight(maybeKey)
            }

            const voice = startVoice(midiToFrequency(midi))
            progressionVoices.push(voice)

            const stopTimeout = window.setTimeout(() => {
              if (!audioContext) {
                return
              }
              voice.stop(audioContext.currentTime)
              if (maybeKey) {
                deactivateKeyHighlight(maybeKey)
              }
            }, Math.max(70, noteMs / step.midiNotes.length))
            progressionTimeouts.push(stopTimeout)
          }, noteStartMs)
          progressionTimeouts.push(startTimeout)
        }
      }

      elapsedMs += slotMs
    }
  }

  const finishTimeout = window.setTimeout(() => {
    stopProgression()
    playBtnEl.disabled = false
    playBtnEl.textContent = 'コード進行を再生'
  }, elapsedMs + 120)
  progressionTimeouts.push(finishTimeout)
}

const playProgressionEl = document.querySelector<HTMLButtonElement>('#play-progression')
if (playProgressionEl) {
  playProgressionEl.addEventListener('click', () => {
    if (isPlayingProgression) {
      return
    }
    playProgression()
  })
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
