import React from 'react';
import * as XLSX from 'xlsx';

function ResultsTable({ data, exportName = '分析结果' }) {
  if (!data || data.length === 0) {
    return <div className="no-results">暂无数据</div>;
  }

  const firstItem = data[0];
  const columns = Object.keys(firstItem).filter(k => k !== '可疑咨询师' && k !== '高度异常');

  const handleExport = () => {
    const exportData = data.map(row => {
      const newRow = {};
      columns.forEach(col => {
        newRow[col] = row[col];
      });
      return newRow;
    });

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    const colWidths = columns.map(col => ({ wch: Math.max(col.length, 12) }));
    worksheet['!cols'] = colWidths;

    const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    XLSX.writeFile(workbook, `${exportName}_${timestamp}.xlsx`);
  };

  return (
    <div className="results-section">
      <div className="results-header">
        <span>明细数据 (共 {data.length} 条)</span>
        <button className="export-btn" onClick={handleExport}>
          导出 Excel
        </button>
      </div>
      <div className="table-wrapper">
        <table>
          <thead>
            <tr>
              {columns.map(col => (
                <th key={col}>{col}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data.slice(0, 2000).map((row, index) => (
              <tr key={index}>
                {columns.map(col => (
                  <td key={col} className={col.includes('现金') ? 'cash' : ''}>
                    {formatValue(row[col])}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {data.length > 2000 && (
        <div className="table-footer">
          仅显示前 2000 条记录（共 {data.length} 条）
        </div>
      )}
    </div>
  );
}

function formatValue(val) {
  if (typeof val === 'number') {
    return val.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  return val ?? '';
}

export default ResultsTable;
