import React from 'react';

function SummaryCards({ items }) {
  const typeMap = {
    info: 'blue',
    warning: 'yellow',
    danger: 'red',
    success: 'green'
  };

  return (
    <div className="summary-grid">
      {items.map((item, index) => (
        <div key={index} className={`stat-card ${typeMap[item.type] || 'blue'}`}>
          <div className="stat-label">{item.label}</div>
          <div className={`stat-value ${typeMap[item.type] || 'blue'}`}>{item.value}</div>
        </div>
      ))}
    </div>
  );
}

export default SummaryCards;
