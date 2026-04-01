import React from 'react';
import { NavLink, Routes, Route, Navigate } from 'react-router-dom';
import SuspiciousAnalysis from './pages/SuspiciousAnalysis';
import ScheduleAnalysis from './pages/ScheduleAnalysis';

function App() {
  return (
    <>
      <header className="header">
        <h1>咨询师绩效审计系统</h1>
        <p>基于可疑业绩达成率区间 + 排班休息日联合分析</p>
      </header>

      <nav className="tabs">
        <NavLink
          to="/suspicious"
          className={({ isActive }) => `tab ${isActive ? 'active' : ''}`}
        >
          可疑业绩区间分析
        </NavLink>
        <NavLink
          to="/schedule"
          className={({ isActive }) => `tab ${isActive ? 'active' : ''}`}
        >
          排班休息日异常分析
        </NavLink>
      </nav>

      <Routes>
        <Route path="/" element={<Navigate to="/suspicious" replace />} />
        <Route path="/suspicious" element={<SuspiciousAnalysis />} />
        <Route path="/schedule" element={<ScheduleAnalysis />} />
      </Routes>
    </>
  );
}

export default App;
