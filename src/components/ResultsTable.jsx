import React from 'react';

function ResultsTable({ data }) {
  if (!data || data.length === 0) {
    return <div className="no-results">暂无数据</div>;
  }

  const firstItem = data[0];
  const columns = Object.keys(firstItem).filter(k => k !== '可疑咨询师' && k !== '高度异常');

  return (
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
          {data.slice(0, 200).map((row, index) => (
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
  );
}

function formatValue(val) {
  if (typeof val === 'number') {
    return val.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  return val ?? '';
}

export default ResultsTable;
