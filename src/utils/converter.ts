import * as XLSX from 'xlsx';

export interface ConversionOptions {
  fps: number;
  startFrame: number;
}

/**
 * テキストに含まれる特定のキーワード（＠縦右, ＠縦左, ＠斜）を、
 * 出現順を維持したまま、重複のみを削除する。
 * 例: "テスト＠縦右＠斜＠縦右" -> "テスト＠縦右＠斜"
 */
export function cleanDuplicateKeywords(text: string): string {
  if (typeof text !== 'string') return text;
  
  const keywords = ["＠縦右", "＠縦左", "＠斜"];
  // Build a regex matching any of the keywords
  const pattern = new RegExp(`(${keywords.join('|')})`, 'g');
  
  const matches = text.match(pattern);
  if (!matches || matches.length === 0) return text;
  
  const uniqueKeywords: string[] = [];
  for (const kw of matches) {
    if (!uniqueKeywords.includes(kw)) {
      uniqueKeywords.push(kw);
    }
  }
  
  // Remove keyword matches from the base text
  let baseText = text;
  for (const kw of keywords) {
    // Replace all occurrences of each keyword with empty string
    baseText = baseText.split(kw).join('');
  }
  
  return baseText + uniqueKeywords.join('');
}

export function frameToAssTime(timeStr: string, fps = 25, offsetFrames = -2, startFrame = 0): string {
  try {
    const parts = String(timeStr).trim().split(':');
    if (parts.length < 4) return "0:00:00.00";
    
    const h = parseInt(parts[0], 10) || 0;
    const m = parseInt(parts[1], 10) || 0;
    const s = parseInt(parts[2], 10) || 0;
    const f = parseInt(parts[3], 10) || 0;
    
    const totalFrames = (h * 3600 + m * 60 + s) * fps + f;
    let appliedFrames = totalFrames - startFrame + offsetFrames;
    if (appliedFrames < 0) appliedFrames = 0;
    
    const totalMs = Math.floor(appliedFrames * (1000 / fps));
    
    const newH = Math.floor(totalMs / 3600000);
    const newM = Math.floor((totalMs % 3600000) / 60000);
    const newS = Math.floor((totalMs % 60000) / 1000);
    const newCs = Math.floor((Math.round(totalMs) % 1000) / 10);
    
    const pad2 = (num: number) => num.toString().padStart(2, '0');
    
    return `${newH}:${pad2(newM)}:${pad2(newS)}.${pad2(newCs)}`;
  } catch (e) {
    return "0:00:00.00";
  }
}

export function convertExcelToAss(fileBuffer: ArrayBuffer, options: ConversionOptions): string {
  // Read array buffer
  const workbook = XLSX.read(fileBuffer, { type: 'array' });
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  
  // Parse rows (start from row 1 usually if there's header)
  const rows: any[] = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
  
  const header = `[Script Info]
Title: Translation Work File (Aegisub)
Script Type: v4.00+
PlayResX: 1920
PlayResY: 1080

[V4+ Styles]
Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding
Style: Audio,Arial,60,&H00FFFFFF,&H000000FF,&H00000000,&H00000000,0,0,0,0,100,100,0,0,1,2,2,2,30,30,50,1
Style: Telop,Arial,60,&H00FFFFFF,&H000000FF,&H00000000,&H00000000,0,0,0,0,100,100,0,0,1,2,2,8,30,30,150,1

[Events]
Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
`;

  let assContent = header;

  for (const row of rows) {
    if (!row['表示開始時間'] || !row['表示終了時間']) continue;
    
    const start = frameToAssTime(String(row['表示開始時間']), options.fps, -2, options.startFrame);
    const end = frameToAssTime(String(row['表示終了時間']), options.fps, -2, options.startFrame);
    
    const sourceText = String(row['原文'] || "").replace(/\n/g, '\\N');
    const style = String(row['トラック']) === "A" ? "Audio" : "Telop";
    
    let targetText = "";
    if (row['字幕'] !== undefined && row['字幕'] !== null && String(row['字幕']).trim() !== "") {
      targetText = String(row['字幕']).replace(/\n/g, '\\N');
    } else {
      // Original logic: use source_text if target_text is missing. Oh wait, source_text replaced newlines!
      // In python, it's: target_text = str(row['字幕']).replace('\n', '\\N') 
      // else: target_text = source_text
      targetText = sourceText;
    }
    
    targetText = cleanDuplicateKeywords(targetText);
    
    assContent += `Comment: 0,${start},${end},${style},SOURCE,0,0,150,,${sourceText}\n`;
    assContent += `Dialogue: 0,${start},${end},${style},JAPANESE,0,0,150,,${targetText}\n`;
  }
  
  return assContent;
}

export function formatFloat(val: any): string {
  if (val === undefined || val === null || val === '') return '';
  const num = parseFloat(val);
  if (isNaN(num)) return String(val);
  // Round to 1 decimal place (e.g. 0.15 -> 0.2)
  return num.toFixed(1);
}

export function convertExcelToTextSummary(fileBuffer: ArrayBuffer): string {
  const workbook = XLSX.read(fileBuffer, { type: 'array' });
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  
  // Use header: 1 to get raw arrays to reliably access columns by index
  const rows: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
  if (!rows || rows.length === 0) return "";
  
  const headerRow = rows[0] || [];
  
  const sourceIdx = headerRow.indexOf('原文');
  const targetIdx = headerRow.indexOf('字幕');
  
  let gIdx = headerRow.indexOf('最大');
  if (gIdx === -1) gIdx = 6; // Fallback to G column
  
  let hIdx = headerRow.indexOf('原稿');
  if (hIdx === -1) hIdx = gIdx + 1; // Fallback to right neighbor of G
  
  let txtContent = "";
  
  // Skip header row
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row.length === 0) continue;
    
    // 1. Source
    const sourceText = sourceIdx !== -1 && row[sourceIdx] !== undefined ? String(row[sourceIdx]) : "";
    // 2. Target
    const targetText = targetIdx !== -1 && row[targetIdx] !== undefined && String(row[targetIdx]).trim() !== "" ? String(row[targetIdx]) : "";
    
    // 3. G and H columns
    const gRaw = row[gIdx] !== undefined ? row[gIdx] : "";
    const hRaw = row[hIdx] !== undefined ? row[hIdx] : "";
    
    const gVal = formatFloat(gRaw);
    let line3 = `最大 ${gVal}`;
    
    const hasHColumn = headerRow.indexOf('原稿') !== -1 || (hRaw !== undefined && hRaw !== "");
    if (hasHColumn) {
      const hVal = formatFloat(hRaw);
      line3 += `, 原稿 ${hVal}`;
    }
    
    txtContent += `${sourceText}\n`;
    txtContent += `${targetText}\n`;
    txtContent += `${line3}\n`;
    txtContent += `\n`;
  }
  
  return txtContent;
}

