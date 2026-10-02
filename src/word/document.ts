export function downloadFile(filename: string, content: string, mime = 'text/plain') {
  const blob = new Blob([content], { type: `${mime};charset=utf-8` });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

export function countWords(html: string): { words: number; chars: number } {
  const text = html
    .replace(/<[^>]*>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .trim();
  if (!text) return { words: 0, chars: 0 };
  return {
    words: text.split(/\s+/).filter(Boolean).length,
    chars: text.length,
  };
}
