import React, { useCallback, useState } from 'react';

function UploadZone({ label, hint, onFileSelected, fileName }) {
  const [dragover, setDragover] = useState(false);

  const handleClick = useCallback(() => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx';
    input.onchange = (e) => {
      const file = e.target.files?.[0];
      if (file) onFileSelected(file);
    };
    input.click();
  }, [onFileSelected]);

  const handleDragOver = useCallback((e) => {
    e.preventDefault();
    setDragover(true);
  }, []);

  const handleDragLeave = useCallback(() => {
    setDragover(false);
  }, []);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    setDragover(false);
    const file = e.dataTransfer.files?.[0];
    if (file && file.name.endsWith('.xlsx')) {
      onFileSelected(file);
    }
  }, [onFileSelected]);

  const isReady = !!fileName;

  return (
    <div
      className={`upload-zone ${dragover ? 'dragover' : ''} ${isReady ? 'uploaded' : ''}`}
      onClick={handleClick}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      <div className="upload-icon">📁</div>
      <div className="upload-label">{label}</div>
      <div className="upload-hint">{hint}</div>
      <div className="upload-action">点击选择文件 或 拖拽文件到此处</div>
      {fileName && <div className="file-name">✓ 已选: {fileName}</div>}
    </div>
  );
}

export default UploadZone;
