import { useState } from 'react';
import WordEditor from './word/WordEditor';
import ExcelGrid from './excel/ExcelGrid';

export default function App() {
  const [tab, setTab] = useState<'word' | 'excel'>('word');

  return (
    <>
      <header className="topbar">
        <div className="logo">
          Open<span>Desk</span>
        </div>
        <nav className="tabs" aria-label="Модули">
          <button
            className={tab === 'word' ? 'tab active' : 'tab'}
            onClick={() => setTab('word')}
          >
            📝 Word
          </button>
          <button
            className={tab === 'excel' ? 'tab active' : 'tab'}
            onClick={() => setTab('excel')}
          >
            📊 Excel
          </button>
        </nav>
        <div style={{ marginLeft: 'auto' }} className="hint hide-mobile">
          MIT • v0.4.0 • локально, без сервера
        </div>
      </header>

      <main className="main">
        {tab === 'word' ? <WordEditor /> : <ExcelGrid />}

        <div className="card">
          <b>OpenDesk — свободный офисный suite.</b>
          <p className="hint">
            Word: 8 вкладок (incl. Рассылки), DOCX со стилями/картинками,
            SmartArt, диаграммы, слияние, мобильная вёрстка.
            <br />
            Excel: сетка 500×52, ~45 функций (<code>=ЕСЛИ</code>, <code>=ВПР</code>), спарклайны, сводные, XLSX.
          </p>
        </div>
      </main>
    </>
  );
}
