import { useMemo } from 'react';
import { parseRange, rangeRefsFull } from './logic';

interface Props {
  range: string;
  type: 'line' | 'column' | 'winloss';
  getNum: (ref: string) => number;
}

export default function Sparkline({ range, type, getNum }: Props) {
  const { points, min, max } = useMemo(() => {
    const box = parseRange(range);
    const vals = box ? rangeRefsFull(box).map(getNum).filter((n) => Number.isFinite(n)) : [];
    if (!vals.length) return { points: [] as string[], min: 0, max: 0 };
    return { points: vals.map(String), min: Math.min(...vals), max: Math.max(...vals) };
  }, [range, getNum]);

  const W = 100;
  const H = 28;
  if (!points.length) return <span className="hint">—</span>;
  const nums = points.map(Number);

  if (type === 'column' || type === 'winloss') {
    const bw = Math.max(2, Math.floor(W / nums.length) - 1);
    return (
      <svg width={W} height={H} viewBox={`0 0 ${W} ${H}`} style={{ display: 'block' }}>
        {nums.map((v, i) => {
          const h = type === 'winloss' ? (v > 0 ? H / 2 - 2 : v < 0 ? H / 2 - 2 : 2) : Math.round(((v - min) / Math.max(max - min, 1e-9)) * (H - 4)) + 2;
          const y = type === 'winloss' ? (v >= 0 ? H / 2 - h : H / 2) : H - h;
          const fill = type === 'winloss' ? (v > 0 ? '#16a34a' : v < 0 ? '#dc2626' : '#9ca3af') : '#2563eb';
          return <rect key={i} x={i * (bw + 1)} y={y} width={bw} height={Math.max(1, h)} fill={fill} />;
        })}
      </svg>
    );
  }
  const step = nums.length > 1 ? W / (nums.length - 1) : 0;
  const d = nums
    .map((v, i) => {
      const y = H - 2 - ((v - min) / Math.max(max - min, 1e-9)) * (H - 4);
      return `${i === 0 ? 'M' : 'L'}${(i * step).toFixed(1)},${y.toFixed(1)}`;
    })
    .join(' ');
  const hi = nums.indexOf(max);
  const lo = nums.indexOf(min);
  return (
    <svg width={W} height={H} viewBox={`0 0 ${W} ${H}`} style={{ display: 'block' }}>
      <path d={d} fill="none" stroke="#2563eb" strokeWidth="1.5" />
      <circle cx={hi * step} cy={H - 2 - ((max - min) / Math.max(max - min, 1e-9)) * (H - 4)} r="2" fill="#16a34a" />
      {lo !== hi && <circle cx={lo * step} cy={H - 2} r="2" fill="#dc2626" />}
    </svg>
  );
}
