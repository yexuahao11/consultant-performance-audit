import React, { useState, useCallback } from 'react';
import UploadZone from '../components/UploadZone';
import ResultsTable from '../components/ResultsTable';
import SummaryCards from '../components/SummaryCards';
import { useAnalysis, useFileUpload } from '../hooks/useAnalysis';

function ScheduleAnalysis() {
  const { files, setFile, clearFiles, hasAllRequired } = useFileUpload();
  const { loading, error, results, analyze, reset } = useAnalysis();

  const handleAnalyze = useCallback(async () => {
    if (!hasAllRequired(['perfFile', 'scheduleFile'])) return;

    const formData = new FormData();
    formData.append('perfFile', files.perfFile);
    formData.append('scheduleFile', files.scheduleFile);

    await analyze('/analyzeSchedule', formData);
  }, [files, analyze, hasAllRequired]);

  const handleReset = useCallback(() => {
    clearFiles();
    reset();
  }, [clearFiles, reset]);

  return (
    <div className="module">
      <div className="upload-section">
        <div className="section-label">📤 请上传两个Excel文件</div>

        <UploadZone
          label="咨询师绩效分析"
          hint="包含 病历号、咨询师、现金类实收、收费时间、收费机构 等字段"
          onFileSelected={(file) => setFile('perfFile', file)}
          fileName={files.perfFile?.name}
        />

        <UploadZone
          label="咨询师排班表"
          hint={"包含 咨询师、日期、排班状态 等字段，休息日标记为'休息'"}
          onFileSelected={(file) => setFile('scheduleFile', file)}
          fileName={files.scheduleFile?.name}
        />
      </div>

      <div className="filter-section">
        <button
          className="analyze-btn"
          onClick={handleAnalyze}
          disabled={!hasAllRequired(['perfFile', 'scheduleFile']) || loading}
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
              { label: '休息日异常人数', value: results.suspiciousConsultants, type: 'danger' },
              { label: '异常记录数', value: results.suspiciousRecords, type: 'danger' },
              { label: '涉及总金额', value: formatCash(results.totalCash), type: 'danger' }
            ]}
          />
          <ResultsTable data={results.records} />
        </div>
      )}
    </div>
  );
}

function formatCash(num) {
  return num.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

export default ScheduleAnalysis;
