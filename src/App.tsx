import React, { useState, useRef, useEffect } from 'react';
import { UploadCloud, FileType, CheckCircle2, AlertCircle, Settings2, FileText } from 'lucide-react';
import { saveAs } from 'file-saver';
import { convertExcelToAss, convertExcelToTextSummary, type ConversionOptions } from './utils/converter';
import './index.css';

function App() {
  const [dragActive, setDragActive] = useState(false);
  const [file, setFile] = useState<File | null>(null);
  const [options, setOptions] = useState<ConversionOptions>({ fps: 25, startFrame: 0 });
  const [outputMode, setOutputMode] = useState<'ass' | 'txt'>('ass');
  const [status, setStatus] = useState<{ type: 'idle' | 'success' | 'error', msg: string }>({ type: 'idle', msg: '' });
  const [assContent, setAssContent] = useState<string | null>(null);

  const inputRef = useRef<HTMLInputElement>(null);

  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      processFile(e.dataTransfer.files[0], outputMode);
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    e.preventDefault();
    if (e.target.files && e.target.files[0]) {
      processFile(e.target.files[0], outputMode);
    }
  };

  async function processFile(selectedFile: File, mode: 'ass' | 'txt') {
    setFile(selectedFile);
    setStatus({ type: 'idle', msg: '' });
    setAssContent(null);

    try {
      if (!selectedFile.name.match(/\.(xlsx|xls)$/i)) {
        throw new Error("アップロードできるのは Excelファイル (.xlsx, .xls) のみです");
      }

      const buffer = await selectedFile.arrayBuffer();
      let outputData = "";

      if (mode === 'ass') {
        outputData = convertExcelToAss(buffer, options);
      } else {
        outputData = convertExcelToTextSummary(buffer);
      }

      setAssContent(outputData);
      setStatus({ type: 'success', msg: '変換に成功しました。ダウンロードボタンから保存してください。' });
    } catch (err: any) {
      setStatus({ type: 'error', msg: err.message || 'ファイルの変換中にエラーが発生しました。' });
      console.error(err);
    }
  }

  const handleDownload = () => {
    if (!assContent || !file) return;

    // Create a Blob with UTF-8 BOM so Aegisub/Notepad recognizes the encoding properly
    const blob = new Blob(["\uFEFF" + assContent], { type: 'text/plain;charset=utf-8' });
    const originalName = file.name.replace(/\.[^/.]+$/, "");

    if (outputMode === 'ass') {
      saveAs(blob, `${originalName}.ass`);
    } else {
      saveAs(blob, `${originalName}_summary.txt`);
    }
  };

  // Re-process if options change and file already exists
  useEffect(() => {
    if (file && status.type === 'success') {
      processFile(file, outputMode);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [options.fps, options.startFrame, outputMode]);

  return (
    <div className="app-container">
      <header className="header">
        <h1 className="title">Excel Converter PWA</h1>
        <p className="subtitle">Excel (.xlsx) をドロップして Aegisub用 の字幕ファイル／対訳テキストを生成します</p>
      </header>

      <main>
        <div
          className={`dropzone-container ${dragActive ? 'active' : ''}`}
          onDragEnter={handleDrag}
          onDragLeave={handleDrag}
          onDragOver={handleDrag}
          onDrop={handleDrop}
          onClick={() => inputRef.current?.click()}
        >
          <input
            ref={inputRef}
            type="file"
            className="file-input"
            accept=".xlsx, .xls"
            onChange={handleChange}
          />
          <div className="icon-wrapper">
            {file && status.type === 'success' ? (
              outputMode === 'ass' ? <FileType size={64} strokeWidth={1.5} /> : <FileText size={64} strokeWidth={1.5} />
            ) : (
              <UploadCloud size={64} strokeWidth={1.5} />
            )}
          </div>

          {file ? (
            <h3 style={{ fontSize: '1.25rem', marginBottom: '0.5rem' }}>{file.name}</h3>
          ) : (
            <h3 style={{ fontSize: '1.25rem', marginBottom: '0.5rem' }}>ファイルを選択するか、ここにドラッグ＆ドロップ</h3>
          )}

          <p style={{ color: 'var(--text-muted)' }}>
            ブラウザ内で完結するため、データは外部サーバーに送信されません
          </p>
        </div>

        <div className="controls-card">
          <div className="control-group">
            <label className="control-label">出力形式</label>
            <div style={{ display: 'flex', gap: '1rem', alignItems: 'center', height: '100%', marginTop: '0.25rem' }}>
              <label style={{ cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                <input type="radio" name="outputMode" value="ass" checked={outputMode === 'ass'} onChange={() => setOutputMode('ass')} style={{ accentColor: 'var(--brand-main)' }} />
                ASS字幕 (.ass)
              </label>
              <label style={{ cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                <input type="radio" name="outputMode" value="txt" checked={outputMode === 'txt'} onChange={() => setOutputMode('txt')} style={{ accentColor: 'var(--brand-main)' }} />
                対訳テキスト (.txt)
              </label>
            </div>
          </div>

          <div className="control-group" style={{ opacity: outputMode === 'ass' ? 1 : 0.5, pointerEvents: outputMode === 'ass' ? 'auto' : 'none', transition: 'opacity var(--transition)' }}>
            <label className="control-label">
              <Settings2 size={16} style={{ display: 'inline', verticalAlign: 'text-bottom', marginRight: '4px' }} />
              FPS
            </label>
            <input
              className="control-input"
              type="number"
              value={options.fps}
              onChange={(e) => setOptions({ ...options, fps: parseFloat(e.target.value) || 25 })}
              step="0.01"
              disabled={outputMode !== 'ass'}
            />
          </div>
          <div className="control-group" style={{ opacity: outputMode === 'ass' ? 1 : 0.5, pointerEvents: outputMode === 'ass' ? 'auto' : 'none', transition: 'opacity var(--transition)' }}>
            <label className="control-label">START FRAME</label>
            <input
              className="control-input"
              type="number"
              value={options.startFrame}
              onChange={(e) => setOptions({ ...options, startFrame: parseInt(e.target.value) || 0 })}
              disabled={outputMode !== 'ass'}
            />
          </div>
        </div>

        {status.msg && (
          <div style={{
            marginTop: '2rem',
            padding: '1rem',
            borderRadius: 'var(--radius-md)',
            backgroundColor: status.type === 'error' ? 'hsla(0, 80%, 65%, 0.1)' : 'hsla(120, 80%, 65%, 0.1)',
            border: `1px solid ${status.type === 'error' ? 'hsl(0, 80%, 65%)' : 'hsl(120, 80%, 65%)'}`,
            display: 'flex',
            alignItems: 'center',
            gap: '0.75rem',
            color: status.type === 'error' ? 'hsl(0, 80%, 75%)' : 'hsl(120, 80%, 75%)',
            animation: 'fadeInUp 0.3s ease-out'
          }}>
            {status.type === 'error' ? <AlertCircle size={20} /> : <CheckCircle2 size={20} />}
            <span>{status.msg}</span>
          </div>
        )}

        {assContent && (
          <div className="result-actions">
            <button className="btn btn-primary" onClick={handleDownload} style={{ width: '100%', justifyContent: 'center' }}>
              {outputMode === 'ass' ? <FileType size={20} /> : <FileText size={20} />}
              DOWNLOAD {outputMode === 'ass' ? '.ASS' : '.TXT'} FILE
            </button>
          </div>
        )}
      </main>
    </div>
  );
}

export default App;
