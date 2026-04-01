import React, { useState, useCallback } from 'react';
import UploadZone from '../components/UploadZone';
import ResultsTable from '../components/ResultsTable';
import SummaryCards from '../components/SummaryCards';
import { useAnalysis, useFileUpload } from '../hooks/useAnalysis';

function SuspiciousAnalysis() {
  const [rangeInput, setRangeInput] = useState('100%-105%');
  const { files, setFile, clearFiles, hasAllRequired } = useFileUpload();
  const { loading, error, results, analyze, reset } = useAnalysis();

  const handleAnalyze = useCallback(async () => {
    if (!hasAllRequired(['perfFile', 'rateFile'])) return;

    const formData = new FormData();
    formData.append('perfFile', files.perfFile);
    formData.append('rateFile', files.rateFile);
    formData.append('rangeInput', rangeInput);

    await analyze('/analyze2', formData);
  }, [files, rangeInput, analyze, hasAllRequired]);

  const handleReset = useCallback(() => {
    clearFiles();
    reset();
    setRangeInput('100%-105%');
  }, [clearFiles, reset]);

  return (
    <div className="module">
      <div className="upload-section">
        <div className="section-label">📤 请上传两个Excel文件</div>

        <UploadZone
          label="咨询师绩效分析"
          hint="包含 病历号、咨询师、现金类实收、收费机构 等字段"
          onFileSelected={(file) => setFile('perfFile', file)}
          fileName={files.perfFile?.name}
        />

        <UploadZone
          label="咨询师业绩达成率"
          hint="包含 院区、咨询师、业绩完成率 等字段"
          onFileSelected={(file) => setFile('rateFile', file)}
          fileName={files.rateFile?.name}
        />
      </div>

      <div className="filter-section">
        <div className="filter-item">
          <span className="filter-label">可疑业绩达成率区间:</span>
          <input
            type="text"
            className="range-input"
            value={rangeInput}
            onChange={(e) => setRangeInput(e.target.value)}
            placeholder="如: 100%-105%"
          />
        </div>

        <button
          className="analyze-btn"
          onClick={handleAnalyze}
          disabled={!hasAllRequired(['perfFile', 'rateFile']) || loading}
        >
          {loading ? '分析中...' : '开始分析'}
        </button>
        <button className="reset-btn" onClick={handleReset}>重置</button>
      </div>

      {loading && <div className="loading"><div className="spinner" /></div>}

      {error && (
        <div className="error-box">
          <div className="error-title">分析失败</div>
          <div className="error-message">{error}</div>
        </div>
      )}

      {results && (
        <div className="results">
          <SummaryCards
            items={[
              { label: '可疑院区数', value: results.summary.suspiciousOrgCount, type: 'info' },
              { label: '可疑咨询师数', value: results.summary.suspiciousConsultantCount, type: 'warning' },
              { label: '存疑情况数', value: results.summary.suspiciousCount, type: 'danger' },
              { label: '存疑现金总额', value: formatCash(results.summary.totalSuspiciousCash), type: 'danger' }
            ]}
          />
          <ResultsTable data={results.suspicious} />
        </div>
      )}
    </div>
  );
}

function formatCash(num) {
  return num.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

export default SuspiciousAnalysis;
