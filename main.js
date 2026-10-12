
import './style.css'
import * as XLSX from 'xlsx'

const KEY = 'tabi-shiori-v1'
const FILE_KEY = 'tabi-shiori-excel-v1'

const initialTrip = {
  title: '北海道 3日間の旅',
  date: '2026-10-15',
  items: [
    {
      id: 1,
      time: '08:30',
      place: '東京駅 丸の内側',
      detail: '集合・受付',
      done: false,
      memo: ''
    },
    {
      id: 2,
      time: '09:00',
      place: '東京駅',
      detail: '出発・乗車案内',
      done: false,
      memo: ''
    },
    {
      id: 3,
      time: '12:00',
      place: '昼食会場',
      detail: '予約人数と到着時間を確認',
      done: false,
      memo: ''
    }
  ]
}

let trip
let excelItems = []
let viewMode = 'timeline'
let itineraryScope = 'daily'
let selectedItineraryDate = ''
let message = ''

try {
  trip = JSON.parse(localStorage.getItem(KEY)) || initialTrip
  excelItems = JSON.parse(localStorage.getItem(FILE_KEY)) || []
} catch {
  trip = initialTrip
  excelItems = []
}

function esc(value = '') {
  return String(value).replace(/[&<>"']/g, c => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  })[c])
}
function formatTripDate(value = '') {
  const text = String(value).trim()
  if (!text) return ''

  const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/)

  if (match) {
    const month = match[1].padStart(2, '0')
    const day = match[2].padStart(2, '0')
    let year = match[3]

    if (year.length === 2) {
      year = Number(year) >= 50 ? `19${year}` : `20${year}`
    }

    return `${year}/${month}/${day}`
  }

  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/)
  if (iso) {
    return `${iso[1]}/${iso[2].padStart(2, '0')}/${iso[3].padStart(2, '0')}`
  }

  return text
}

function saveTrip() {
  localStorage.setItem(KEY, JSON.stringify(trip))
}

function saveExcelItems() {
  localStorage.setItem(FILE_KEY, JSON.stringify(excelItems))
}

function getCell(row, name) {
  return row[name] == null ? '' : String(row[name]).trim()
}



function readExcel(file) {
  const reader = new FileReader()

  reader.onerror = () => {
    alert('Excelファイルを読み込めませんでした。')
  }

  reader.onload = event => {
    try {
      const data = new Uint8Array(event.target.result)

      const workbook = XLSX.read(data, {
        type: 'array',
        cellDates: true
      })

      console.log('Excelのシート一覧:', workbook.SheetNames)

      function readSheet(sheetName) {
        const actualName = workbook.SheetNames.find(
          name => name.trim() === sheetName
        )

        if (!actualName) {
          console.warn(`シート「${sheetName}」が見つかりません`)
          return []
        }

        const sheet = workbook.Sheets[actualName]

        return XLSX.utils.sheet_to_json(sheet, {
          defval: '',
          raw: false
        })
      }

      const rows = readSheet('行程')
      const facilities = readSheet('施設・予約')
      const contacts = readSheet('連絡先')
      const overviewRows = readSheet('全体行程')
      const tourInfo = readSheet('ツアー情報')

      console.log('行程の列名:', Object.keys(rows[0] || {}))
      console.log('行程データ:', rows)

      if (!rows.length) {
        alert(
          '「行程」シートのデータが見つかりません。\n' +
          'Excelのシート名と1行目の列名を確認してください。\n' +
          '検出したシート：' + workbook.SheetNames.join(', ')
        )
        return
      }

      const getValue = (row, ...names) => {
        for (const name of names) {
          const key = Object.keys(row).find(
            k => k.trim() === name
          )
          if (key && row[key] != null && String(row[key]).trim()) {
            return String(row[key]).trim()
          }
        }
        return ''
      }

      const tour = tourInfo[0] || {}

      trip.title = getValue(tour, 'ツアー名') || trip.title
      trip.startDate = getValue(tour, '開始日')
      trip.endDate = getValue(tour, '終了日')
      trip.code =
        getValue(
          tour,
          '出発日YYYYMMDD_ツアーコード',
          'ツアーコード',
          'ツアーID'
        )

      const imported = rows.map((row, index) => ({
        id: getValue(row, '行程ID') ||
          `excel-${index}-${Date.now()}`,
        date: getValue(row, '日付'),
        day: getValue(row, '日目'),
        time: getValue(row, '開始時刻') || '--:--',
        end: getValue(row, '終了時刻'),
        place:
          getValue(row, '予定名', '観光場所', '場所', '訪問先') ||
          getValue(row, '区分') ||
          '名称未設定',
        category: getValue(row, '区分'),
        detail: getValue(row, '詳細', '内容'),
        notes: getValue(row, '注意事項'),
        facilityId: getValue(row, '施設ID'),
        done: getValue(row, '完了状態', '対応済み') === '完了',
        memo: getValue(row, '添乗メモ', 'メモ'),
        order: Number(getValue(row, '表示順')) || index + 1
      }))

      excelItems = imported.sort((a, b) => {
        const dateCompare = a.date.localeCompare(b.date)
        return dateCompare || a.order - b.order
      })

      localStorage.setItem(
        'tabi-shiori-facilities-v1',
        JSON.stringify(facilities)
      )
      localStorage.setItem(
        'tabi-shiori-contacts-v1',
        JSON.stringify(contacts)
      )
      localStorage.setItem(
        'tabi-shiori-overview-v1',
        JSON.stringify(overviewRows)
      )

      saveExcelItems()
      saveTrip()

      message =
        `読み込み完了：行程 ${excelItems.length}件、` +
        `施設・予約 ${facilities.length}件、` +
        `連絡先 ${contacts.length}件`

      render()

      if (message) {
        alert(message)
      }      

    } catch (error) {
      console.error('Excel読み込みエラー:', error)
      alert(
        'Excel読み込みエラー：' +
        (error?.message || String(error))
      )
    }
  }

  reader.readAsArrayBuffer(file)
}

function render() {
  const items = excelItems.length ? excelItems : trip.items

  document.querySelector('#app').innerHTML = `
      
  <header class="hero">
    <h1 class="hero-title">${esc(trip.title)}</h1>

    <div class="hero-subtitle">
      <span>${esc(formatTripDate(trip.startDate || trip.date || ''))}${trip.endDate ? ' - ' + esc(formatTripDate(trip.endDate)) : ''}</span>
        <span class="hero-code">${esc(trip.code || '')}</span>
    </div>

    

    <div class="offline">● 端末内保存モード</div>
  </header>

    <main>
      
      
      <nav class="screen-tabs">
        <button data-screen="itinerary" class="${(window.currentScreen || 'itinerary') === 'itinerary' ? 'active' : ''}">行程</button>
        <button data-screen="history" class="${window.currentScreen === 'history' ? 'active' : ''}">記録</button>
        <button data-screen="contacts" class="${window.currentScreen === 'contacts' ? 'active' : ''}">連絡先</button>
        <button data-screen="facilities" class="${window.currentScreen === 'facilities' ? 'active' : ''}">施設・予約</button>
        <button data-screen="settings" class="${window.currentScreen === 'settings' ? 'active' : ''}">設定</button>
      </nav>

      
      
    ${(window.currentScreen || 'itinerary') === 'itinerary' ? `
      <nav class="scope-tabs">
        <button type="button" data-scope="daily"
          class="${itineraryScope === 'daily' ? 'active' : ''}">
          日付ごと
        </button>
        <button type="button" data-scope="all"
          class="${itineraryScope === 'all' ? 'active' : ''}">
          全体行程
        </button>
      </nav>
    ` : ''}

      <section id="screen-content">
        ${renderCurrentScreen(items)}
      </section>


      <p class="footnote">
        読み込んだExcelの旅程とメモはこのブラウザ内に保存されます。
        OneDriveへの同期はまだ設定されていません。
      </p>

      ${window.currentScreen === 'settings' ? `
        <section class="import-panel">
          <h2>Excelから旅程を読み込む</h2>
          <input id="excel-file" type="file" accept=".xlsx,.xls">
        </section>
      ` : ''}

    </main>
  `

const tripDateInput = document.querySelector('#trip-date')

if (tripDateInput) {
  tripDateInput.addEventListener('change', event => {
    trip.date = event.target.value
    saveTrip()
  })
}

  document.querySelector('#excel-file')?.addEventListener('change', event => {
    const file = event.target.files?.[0]
    if (file) readExcel(file)
  })



document.querySelectorAll('[data-itinerary-date]').forEach(button => {
  button.addEventListener('click', () => {
    selectedItineraryDate = button.dataset.itineraryDate
    render()
  })
})

  document.querySelector('#itinerary-date')?.addEventListener('change', event => {
    selectedItineraryDate = event.target.value
    render()
  })

  document.querySelectorAll('[data-screen]').forEach(button => {
    button.addEventListener('click', () => {
      window.currentScreen = button.dataset.screen
      render()
    })
  })

  document.querySelectorAll('[data-scope]').forEach(button => {
    button.addEventListener('click', () => {
      itineraryScope = button.dataset.scope
      render()
    })
  })


  
const saveSettingsButton = document.querySelector('#save-settings')

if (saveSettingsButton) {
  saveSettingsButton.addEventListener('click', () => {
    const selectedTheme = document.querySelector(
      'input[name="setting-theme"]:checked'
    )

    const settings = {
      theme: selectedTheme?.value || 'forest',
      background: document.querySelector('#setting-background')?.value || 'white',
      density: document.querySelector('#setting-density').value,
      fontSize: document.querySelector('#setting-font').value
    }

    localStorage.setItem(
      'tabi-shiori-settings-v1',
      JSON.stringify(settings)
    )

    applyDisplaySettings(settings)

    const message = document.querySelector('#settings-message')
    if (message) {
      message.textContent = '設定を保存しました。'
    }

    // 選択中のカード表示も更新
    render()
  })
}
 


function updateSettings() {
  const selectedTheme = document.querySelector(
    'input[name="setting-theme"]:checked'
  )

  const settings = {
    theme: selectedTheme?.value || 'forest',
    background: document.querySelector('#setting-background')?.value || 'white',
    density: document.querySelector('#setting-density')?.value || 'normal',
    fontSize: document.querySelector('#setting-font')?.value || 'normal'
  }

  localStorage.setItem(
    'tabi-shiori-settings-v1',
    JSON.stringify(settings)
  )

  applyDisplaySettings(settings)

  // テーマカードの選択表示を更新
  document.querySelectorAll('.theme-card').forEach(card => {
    const radio = card.querySelector(
      'input[name="setting-theme"]'
    )
    const isSelected = radio?.checked

    card.classList.toggle('selected', isSelected)

    const label = card.querySelector('.selected-label')
    if (label) {
      label.textContent = isSelected ? '選択中' : '選択する'
    }
  })
}

document.querySelectorAll(
  'input[name="setting-theme"]'
).forEach(input => {
  input.addEventListener('change', updateSettings)
})

;['#setting-background', '#setting-density', '#setting-font']
  .forEach(selector => {
    document.querySelector(selector)?.addEventListener(
      'change',
      updateSettings
    )
  })

  
const historyForm = document.querySelector('#history-form')

if (historyForm) {
  historyForm.addEventListener('submit', event => {
    event.preventDefault()

    const formData = new FormData(historyForm)
    const place = formData.get('item')
    const type = formData.get('type')
    const reason = formData.get('reason') || ''

    addChangeRecord(
      place,
      type,
      reason
    )

    const history = getChangeHistory()
    const latest = history[0]

    if (latest) {
      latest.item = place
      latest.before = type
      latest.after = type
      latest.reason = reason
      saveChangeHistory(history)
    }

    render()
  })
}

  document.querySelectorAll('[data-done]').forEach(el => {
    el.addEventListener('change', event => {
      const item = items[Number(event.target.dataset.done)]
      item.done = event.target.checked
      if (excelItems.length) saveExcelItems()
      else saveTrip()
      render()
    })
  })

  document.querySelectorAll('[data-memo]').forEach(el => {
    el.addEventListener('change', event => {
      const item = items[Number(event.target.dataset.memo)]
      item.memo = event.target.value
      if (excelItems.length) saveExcelItems()
      else saveTrip()
      message = 'メモを端末内に保存しました。'
      render()
    })
  })

  document.querySelector('#add-stop')?.addEventListener('click', () => {
    trip.items.push({
      id: Date.now(),
      time: '10:00',
      place: '新しい訪問先',
      detail: '予定内容を入力',
      done: false,
      memo: ''
    })
    saveTrip()
    if (!excelItems.length) render()
    else alert('Excel読み込み中はサンプル行程に追加しました。')
  })

  
  document.querySelector('#backup')?.addEventListener('click', () => {
    const history = getChangeHistory()

    const workbook = XLSX.utils.book_new()

    const currentItems = excelItems.length ? excelItems : trip.items

    const itineraryRows = currentItems.map(item => ({
      日付: item.date || trip.date || '',
      日目: item.day || '',
      開始時刻: item.time || '',
      終了時刻: item.end || '',
      予定名: item.place || '',
      区分: item.category || '',
      詳細: item.detail || '',
      注意事項: item.notes || '',
      添乗メモ: item.memo || '',
      対応済み: item.done ? '完了' : '未完了'
    }))

    const facilities = JSON.parse(
      localStorage.getItem('tabi-shiori-facilities-v1') || '[]'
    )

    const contacts = JSON.parse(
      localStorage.getItem('tabi-shiori-contacts-v1') || '[]'
    )

    const historyRows = history.map(record => ({
      記録日時: record.recordedAt || '',
      項目: record.item || '',
      変更前: record.before || '',
      変更後: record.after || '',
      変更理由: record.reason || ''
    }))

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(itineraryRows),
      '行程'
    )

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(facilities),
      '施設・予約'
    )

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(contacts),
      '連絡先'
    )

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(historyRows),
      '記録'
    )

    XLSX.writeFile(workbook, '旅のしおり_バックアップ.xlsx')
  })
}

function getChangeHistory() {
  try {
    return JSON.parse(
      localStorage.getItem('tabi-shiori-history-v1') || '[]'
    )
  } catch {
    return []
  }
}

function saveChangeHistory(history) {
  localStorage.setItem(
    'tabi-shiori-history-v1',
    JSON.stringify(history)
  )
}

function addChangeRecord(before, after, reason) {
  const history = getChangeHistory()

  history.unshift({
    id: Date.now(),
    recordedAt: new Date().toLocaleString(),
    before,
    after,
    reason
  })

  saveChangeHistory(history)
}

function applyDisplaySettings(settings) {
  const root = document.documentElement

  
const themes = {
  forest: {
    color: '#285943',
    sub: '#879B83',
    background: '#FDF9F6'
  },
  french: {
    color: '#824C55',
    sub: '#C58E91',
    background: '#FCFBF7'
  },
  modern: {
    color: '#414A46',
    sub: '#9BAEB4',
    background: '#F8F9F6'
  },
  italian: {
    color: '#687A42',
    sub: '#C95F4B',
    background: '#F5EBDD'
  },
  clean: {
    color: '#347C7A',
    sub: '#82B6D0',
    background: '#FAFBF9'
  }
}

const selectedTheme = themes[settings.theme] || themes.forest

root.style.setProperty('--theme-color', selectedTheme.color)
root.style.setProperty('--theme-sub-color', selectedTheme.sub)
root.style.setProperty('--page-background', selectedTheme.background)

  const fontSizes = {
    normal: '16px',
    large: '18px',
    xlarge: '21px'
  }

  const densities = {
    comfortable: '18px',
    normal: '12px',
    compact: '7px'
  }

  root.style.setProperty(
    '--app-font-size',
    fontSizes[settings.fontSize] || fontSizes.normal
  )

  root.style.setProperty(
    '--item-padding',
    densities[settings.density] || densities.normal
  )

  root.style.colorScheme = 'light'
}

function renderOverview() {
  let rows = []

  try {
    rows = JSON.parse(
      localStorage.getItem('tabi-shiori-overview-v1') || '[]'
    )
  } catch {}

  if (!rows.length) {
    return '<p class="hint">全体行程データがありません。Excelを読み込んでください。</p>'
  }

  return `
    <section class="overview-list">
      ${rows.map(row => `
        <article class="overview-item">
          <div class="overview-date">
            <span class="overview-day">${esc(getCell(row, '日目'))}</span>
            <span class="overview-date-text">${esc(getCell(row, '日付'))}</span>
          </div>

          <div class="overview-body">
            <h3>${esc(getCell(row, '予定名'))}</h3>
            <p>${esc(getCell(row, '概要'))}</p>
          </div>
        </article>
      `).join('')}
    </section>
  `
}

function renderCurrentScreen(items) {
  const screen = window.currentScreen || 'itinerary'

  if (screen === 'facilities') {
    let facilities = []
    try {
      facilities = JSON.parse(
        localStorage.getItem('tabi-shiori-facilities-v1') || '[]'
      )
    } catch {}

    if (!facilities.length) {
      return '<p class="hint">施設・予約データがありません。Excelを読み込んでください。</p>'
    }

    
return facilities.map(row => {
  const phone = getCell(row, '電話番号')
  const phoneText = phone || ''

  return `
    <article class="stop contact-card facility-card">
      <div class="stop-body">

        <span class="contact-category">
          ${esc(getCell(row, '区分') || '施設・予約')}
        </span>

        <h3 class="facility-name">
          ${esc(getCell(row, '施設名') || '名称未登録')}
        </h3>

        ${getCell(row, '住所') ? `
          <div class="contact-field">
            <span class="contact-label">住所</span>
            <p>${esc(getCell(row, '住所'))}</p>
          </div>
        ` : ''}

        ${phone ? `
          <div class="contact-field">
            <div class="contact-phone-row">
              <a class="contact-call"
                 href="tel:${esc(phone.replace(/[^\d+]/g, ''))}"
                 aria-label="電話をかける"
                 title="電話をかける">☎</a>
              <span class="contact-phone-number">
                ${esc(phoneText)}
              </span>
            </div>
          </div>
        ` : ''}


        ${getCell(row, '地図URL') ? `
          <a class="facility-map"
             href="${esc(getCell(row, '地図URL'))}"
             target="_blank"
             rel="noopener noreferrer">
            地図を開く ↗
          </a>
        ` : ''}

        ${getCell(row, '備考') ? `
          <div class="contact-field">
            <span class="contact-label">備考</span>
            <p>${esc(getCell(row, '備考'))}</p>
          </div>
        ` : ''}

      </div>
    </article>
  `
}).join('')
  }

  if (screen === 'contacts') {
    let contacts = []
    try {
      contacts = JSON.parse(
        localStorage.getItem('tabi-shiori-contacts-v1') || '[]'
      )
    } catch {}

    if (!contacts.length) {
      return '<p class="hint">連絡先データがありません。Excelを読み込んでください。</p>'
    }

    return contacts.map(row => {
      const phone = getCell(row, '電話番号')
      const country = getCell(row, '国番号')
      const phoneText = [country, phone].filter(Boolean).join(' ')

      
    
return `
  <article class="stop contact-card">
    <div class="stop-body">

      <span class="contact-category">
        ${esc(getCell(row, '区分') || '未分類')}
      </span>

      <div class="contact-field">
        <p class="contact-person">
          ${esc(getCell(row, '担当者') || '未登録')}
        </p>
      </div>

      <div class="contact-field">
        <div class="contact-phone-row">
          ${phone ? `
            <a class="contact-call"
               href="tel:${esc(phone.replace(/[^\d+]/g, ''))}"
               aria-label="電話をかける"
               title="電話をかける">☎</a>
          ` : ''}
          <span class="contact-phone-number">
            ${esc(phoneText || '未登録')}
          </span>
        </div>
      </div>

      ${getCell(row, '備考') ? `
        <div class="contact-field">
          <span class="contact-label">備考</span>
          <p>${esc(getCell(row, '備考'))}</p>
        </div>
      ` : ''}

    </div>
  </article>
`
    }).join('')
  }
  
  
if (screen === 'history') {
  const history = getChangeHistory()

  return `
    <section class="import-panel">
      <h2>記録</h2>

      <form id="history-form">
        
    <div class="history-input-row">
      <label>
        項目
        <select name="type" required>
          <option value="到着">到着</option>
          <option value="出発">出発</option>
          <option value="その他">その他</option>
        </select>
      </label>

      <label>
        観光場所・施設名
        <input
          name="item"
          required
          placeholder="例：エッフェル塔"
        >
      </label>
    </div>

        <label>
          備考
          <textarea
            name="reason"
            rows="3"
            placeholder="例：10時に到着、チケット受取済み"
          ></textarea>
        </label>

        <button type="submit">保存</button>
      </form>
    </section>

    <h2>記録一覧（${history.length}件）</h2>

    ${
      history.length
        ? history.map(record => `
            <article class="stop">
              <div class="stop-body">
                
              <div class="history-record-heading">
                <span class="contact-category">
                  ${esc(record.after || record.before || 'その他')}
                </span>
                <h3>${esc(record.item || '名称未登録')}</h3>
              </div>

                ${
                  record.reason
                    ? `<p>${esc(record.reason)}</p>`
                    : ''
                }

                <p class="date-text">
                  ${esc(record.recordedAt || '')}
                </p>
              </div>
            </article>
          `).join('')
        : '<p class="hint">まだ記録はありません。</p>'
    }
  `
}
  
  if (screen === 'settings') {
    const settings = JSON.parse(
      localStorage.getItem('tabi-shiori-settings-v1') || '{}'
    )

    
const theme = settings.theme || 'forest'
const background = settings.background || 'white'


const themeOptions = [
  {
    id: 'forest',
    name: '森と植物',
    colors: ['#285943', '#879B83', '#F5F1E6']
  },
  {
    id: 'french',
    name: 'フレンチアンティーク',
    colors: ['#824C55', '#C58E91', '#F7F1E8']
  },
  {
    id: 'modern',
    name: 'モダン・ミニマル',
    colors: ['#414A46', '#9BAEB4', '#F8F9F6']
  },
  {
    id: 'italian',
    name: 'イタリアの夏',
    colors: ['#687A42', '#C95F4B', '#F5EBDD']
  },
  {
    id: 'clean',
    name: '空と雲',
    colors: ['#347C7A', '#82B6D0', '#FAFBF9']
  }
]

    const density = settings.density || 'normal'
    const fontSize = settings.fontSize || 'normal'

    return `
      <section class="settings-panel">
        <h2>表示設定</h2>

        
<label>テーマカラー</label>
<div class="theme-options">
  ${themeOptions.map(option => `
    <label class="theme-card ${theme === option.id ? 'selected' : ''}">
      <input
        type="radio"
        name="setting-theme"
        value="${option.id}"
        ${theme === option.id ? 'checked' : ''}
      >
      <span class="theme-name">${option.name}</span>
      <span class="theme-swatches">
        ${option.colors.map(color => `
          <span class="theme-swatch" style="background-color: ${color}"></span>
        `).join('')}
      </span>
      <span class="selected-label">${theme === option.id ? '✓ 選択中' : '選択する'}</span>
    </label>
  `).join('')}
</div>

       
      </section>
    `
  }

if (itineraryScope === 'all') {
  return renderOverview()
}

function normalizeDate(value) {
  if (!value) return ''

  const text = String(value).trim()

  // YYYY-MM-DD または YYYY/MM/DD 形式
  const match = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/)
  if (match) {
    return `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`
  }

  // Excelの日付文字列など、その他の形式
  const date = new Date(value)
  if (Number.isNaN(date.getTime())) return text

  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')

  return `${year}-${month}-${day}`
}

const dates = [...new Set(
  items.map(item => normalizeDate(item.date)).filter(Boolean)
)].sort()

if (!selectedItineraryDate && dates.length) {
  selectedItineraryDate = dates[0]
}


const dailyItems = selectedItineraryDate
  ? items.filter(item => normalizeDate(item.date) === selectedItineraryDate)
  : items

return `

<section class="daily-selector">
  <div class="date-scroller">
    ${dates.map((date, index) => {
      const d = new Date(date)
      const label = Number.isNaN(d.getTime())
        ? String(date)
        : `${d.getMonth() + 1}/${d.getDate()} (${['日','月','火','水','木','金','土'][d.getDay()]})`

      return `
        <button
          type="button"
          class="date-chip ${date === selectedItineraryDate ? 'active' : ''}"
          data-itinerary-date="${esc(date)}"
          aria-pressed="${date === selectedItineraryDate}">
          <span class="date-chip-day">${index + 1}日目</span>
          <span class="date-chip-date">${esc(label)}</span>
        </button>
      `
    }).join('')}
  </div>
</section>  

  ${viewMode === 'timeline'
    ? renderTimeline(dailyItems)
    : renderTable(dailyItems)}
`
}

function renderTimeline(items) {
  return items.map((item, index) => `
    <article class="stop ${item.done ? 'completed' : ''}">
      <div class="time">${esc(item.time)}</div>
      <div class="stop-body">
        <h3>${esc(item.place)}</h3>
        <p>${esc(item.detail)}</p>
        
        <label class="memo-label">
          添乗メモ
          <textarea data-memo="${index}" rows="2" placeholder="変更・到着時刻・申し送り">${esc(item.memo)}</textarea>
        </label>
      </div>
    </article>
  `).join('')
}

function renderTable(items) {
  return `
    <div class="table-wrap">
      <table>
        <thead>
          <tr><th>日付</th><th>時刻</th><th>予定</th><th>内容・注意事項</th><th>済</th></tr>
        </thead>
        <tbody>
          ${items.map((item, index) => `
            <tr>
              <td>${esc(item.date || '')}</td>
              <td>${esc(item.time)}${item.end ? `〜${esc(item.end)}` : ''}</td>
              <td><strong>${esc(item.place)}</strong></td>
              <td>${esc(item.detail)}${item.notes ? `<br>${esc(item.notes)}` : ''}
                <label class="memo-label">
                  <textarea data-memo="${index}" rows="2" placeholder="添乗メモ">${esc(item.memo)}</textarea>
                </label>
              </td>
              <td><input type="checkbox" data-done="${index}" ${item.done ? 'checked' : ''} aria-label="対応済み"></td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    </div>
  `
}

const initialSettings = JSON.parse(
  localStorage.getItem('tabi-shiori-settings-v1') || '{}'
)

applyDisplaySettings(initialSettings)

render()