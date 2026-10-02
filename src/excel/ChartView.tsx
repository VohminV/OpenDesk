import { useMemo } from 'react';
import { parseRange, rangeRefsFull } from './logic';
import type { ChartDef } from './format';

interface Props {
  chart: ChartDef;
  getNum: (ref: string) => number;
  getText: (ref: string) => string;
  onDelete: () => void;
}

export default function ChartView({ chart, getNum, getText, onDelete }: Props) {
  const { values, labels } = useMemo(() => {
    const box = parseRange(chart.range);
    const refs = box ? rangeRefsFull(box) : [];
    const values = refs.map(getNum).filter((n) => Number.isFinite(n)).slice(0, 24);
    let labels = values.map((_, i) => `П${i + 1}`);
    if (chart.labels) {
      const lb = parseRange(chart.labels);
      if (lb) {
        const t = rangeRefsFull(lb).map(getText).slice(0, 24);
        if (t.length === values.length) labels = t;
      }
    }
    return { values, labels };
  }, [chart, getNum, getText]);

  const W = 340;
  const H = 180;
  const max = Math.max(...values, 1);

  let body: React.ReactNode = null;
  if (chart.type === 'pie') {
    const total = values.reduce((a, b) => a + b, 0) || 1;
    const colors = ['#2563eb', '#16a34a', '#eab308', '#db2777', '#7c3aed', '#ea580c', '#06b6d4', '#84cc16'];
    let acc = 0;
    const cx = 90;
    const cy = 90;
    const r = 70;
    const slices = values.map((v, i) => {
      const a0 = (acc / total) * Math.PI * 2;
      acc += v;
      const a1 = (acc / total) * Math.PI * 2;
      const large = a1 - a0 > Math.PI ? 1 : 0;
      const x0 = cx + r * Math.cos(a0);
      const y0 = cy + r * Math.sin(a0);
      const x1 = cx + r * Math.cos(a1);
      const y1 = cy + r * Math.sin(a1);
      return { d: `M${cx},${cy} L${x0.toFixed(1)},${y0.toFixed(1)} A${r},${r} 0 ${large},1 ${x1.toFixed(1)},${y1.toFixed(1)} Z`, fill: colors[i % colors.length], label: labels[i], v };
    });
    body = (
      <svg width={W} height={H} viewBox={`0 0 ${W} ${H}`}>
        {slices.map((s, i) => (
          <path key={i} d={s.d} fill={s.fill} stroke="#fff" strokeWidth="1">
            <title>{`${s.label}: ${s.v}`}</title>
          </path>
        ))}
        {slices.map((s, i) => (
          <text key={`t${i}`} x={200} y={30 + i * 18} fontSize="11" fill="#333">
            <tspan fill={s.fill}>■ </tspan>
            {s.label} ({s.v})
          </text>
        ))}
      </svg>
    );
  } else if (chart.type === 'line') {
    const step = values.length > 1 ? (W - 40) / (values.length - 1) : 0;
    const d = values
      .map((v, i) => `${i === 0 ? 'M' : 'L'}${(20 + i * step).toFixed(1)},${(H - 20 - (v / max) * (H - 50)).toFixed(1)}`)
      .join(' ');
    body = (
      <svg width={W} height={H} viewBox={`0 0 ${W} ${H}`}>
        <path d={d} fill="none" stroke="#2563eb" strokeWidth="2" />
        {values.map((v, i) => (
          <circle key={i} cx={20 + i * step} cy={H - 20 - (v / max) * (H - 50)} r="3" fill="#2563eb">
            <title>{`${labels[i]}: ${v}`}</title>
          </circle>
        ))}
        {labels.map((l, i) => (
          <text key={`l${i}`} x={20 + i * step} y={H - 5} fontSize="9" textAnchor="middle" fill="#666">
            {l}
          </text>
        ))}
      </svg>
    );
  } else {
    const bw = Math.max(8, Math.floor((W - 40) / Math.max(values.length, 1)) - 6);
    body = (
      <svg width={W} height={H} viewBox={`0 0 ${W} ${H}`}>
        {values.map((v, i) => {
          const h = Math.round((v / max) * (H - 50));
          return (
            <g key={i}>
              <rect x={20 + i * (bw + 6)} y={H - 20 - h} width={bw} height={h} rx="3" fill="#2563eb">
                <title>{`${labels[i]}: ${v}`}</title>
              </rect>
              <text x={20 + i * (bw + 6) + bw / 2} y={H - 6} fontSize="9" textAnchor="middle" fill="#666">
                {labels[i]}
              </text>
            </g>
          );
        })}
      </svg>
    );
  }

  return (
    <div className="chart-box">
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <b>{chart.title}</b>
        <button className="tab" onClick={onDelete} title="Удалить диаграмму">
          ✕
        </button>
      </div>
      <div className="hint">
        {chart.type === 'bar' ? 'Гистограмма' : chart.type === 'line' ? 'График' : 'Круговая'} • {chart.range}
      </div>
      {values.length ? body : <p className="hint">Нет числовых данных в {chart.range}</p>}
    </div>
  );
}
