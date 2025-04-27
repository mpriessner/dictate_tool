import React from 'react';
import { MODE_COLORS } from './MonthlyUtils';

const MonthlyDistribution = ({ stats }) => {
  // Calculate total time and percentages
  const totalTime = stats.totalWork + stats.totalLeisure + stats.totalBreak + stats.totalInactive;
  
  // Calculate percentages
  const workPercent = totalTime > 0 ? stats.totalWork / totalTime : 0;
  const leisurePercent = totalTime > 0 ? stats.totalLeisure / totalTime : 0;
  const breakPercent = totalTime > 0 ? stats.totalBreak / totalTime : 0;
  const inactivePercent = totalTime > 0 ? stats.totalInactive / totalTime : 0;
  
  // Calculate stroke dash values
  const circumference = 251.2; // 2 * PI * r where r = 40
  const workDash = workPercent * circumference;
  const leisureDash = leisurePercent * circumference;
  const breakDash = breakPercent * circumference;
  const inactiveDash = inactivePercent * circumference;
  
  // Calculate offsets
  const inactiveOffset = 0;
  const breakOffset = -1 * inactiveDash;
  const leisureOffset = -1 * (inactiveDash + breakDash);
  const workOffset = -1 * (inactiveDash + breakDash + leisureDash);

  // Format time for display
  const formatTime = (minutes) => {
    const h = Math.floor(minutes / 60);
    const m = minutes % 60;
    return `${h}h ${m}m`;
  };

  return (
    <div>
      <h2 className="text-xl font-semibold mb-4">Focus Distribution</h2>
      <div className="flex justify-center">
        <div className="relative w-64 h-64">
          {/* Render donut chart */}
          <svg viewBox="0 0 100 100" className="w-full h-full">
            {/* Background circle */}
            <circle cx="50" cy="50" r="40" fill="transparent" stroke="#333" strokeWidth="20" />
            
            {/* Inactive segment first */}
            {stats.totalInactive > 0 && (
              <circle 
                cx="50" 
                cy="50" 
                r="40" 
                fill="transparent" 
                stroke={MODE_COLORS.inactive}
                strokeWidth="20"
                strokeDasharray={`${inactiveDash} ${circumference}`}
                strokeDashoffset={inactiveOffset}
                transform="rotate(-90 50 50)"
              />
            )}
            
            {/* Break segment */}
            {stats.totalBreak > 0 && (
              <circle 
                cx="50" 
                cy="50" 
                r="40" 
                fill="transparent" 
                stroke={MODE_COLORS.break}
                strokeWidth="20"
                strokeDasharray={`${breakDash} ${circumference}`}
                strokeDashoffset={breakOffset}
                transform="rotate(-90 50 50)"
              />
            )}
            
            {/* Leisure focus segment */}
            {stats.totalLeisure > 0 && (
              <circle 
                cx="50" 
                cy="50" 
                r="40" 
                fill="transparent" 
                stroke={MODE_COLORS.leisure}
                strokeWidth="20"
                strokeDasharray={`${leisureDash} ${circumference}`}
                strokeDashoffset={leisureOffset}
                transform="rotate(-90 50 50)"
              />
            )}
            
            {/* Work focus segment */}
            {stats.totalWork > 0 && (
              <circle 
                cx="50" 
                cy="50" 
                r="40" 
                fill="transparent" 
                stroke={MODE_COLORS.work}
                strokeWidth="20"
                strokeDasharray={`${workDash} ${circumference}`}
                strokeDashoffset={workOffset}
                transform="rotate(-90 50 50)"
              />
            )}
            
            {/* Center text - show detailed breakdown */}
            <text x="50" y="42" textAnchor="middle" className="text-xxs" fill="white" style={{ fontSize: '0.6rem' }}>
              {Math.round((stats.totalWork + stats.totalLeisure) / totalTime * 100)}% Focus
            </text>
            <text x="50" y="52" textAnchor="middle" className="text-xxs" fill="#FFCA28" style={{ fontSize: '0.6rem' }}>
              {stats.totalLeisure > 0 ? `${Math.round((stats.totalLeisure / totalTime) * 100)}% Leisure` : ''}
            </text>
            <text x="50" y="62" textAnchor="middle" className="text-xxs" fill="#3B82F6" style={{ fontSize: '0.6rem' }}>
              {stats.totalWork > 0 ? `${Math.round((stats.totalWork / totalTime) * 100)}% Work` : ''}
            </text>
          </svg>
        </div>
      </div>
      
      {/* Legend */}
      <div className="flex justify-center mt-6 space-x-4 flex-wrap">
        {/* Work category */}
        <div className="flex items-center bg-gray-700 px-3 py-2 rounded-md">
          <div className="w-4 h-4 bg-blue-500 mr-2 rounded-sm"></div>
          <span className="font-medium">Work: {formatTime(stats.totalWork)}</span>
        </div>
        
        {/* Leisure category */}
        <div className="flex items-center bg-gray-700 px-3 py-2 rounded-md">
          <div className="w-4 h-4 bg-amber-400 mr-2 rounded-sm"></div>
          <span className="font-medium">Leisure: {formatTime(stats.totalLeisure)}</span>
        </div>
        
        {/* Break category */}
        <div className="flex items-center bg-gray-700 px-3 py-2 rounded-md">
          <div className="w-4 h-4 bg-cyan-400 mr-2 rounded-sm"></div>
          <span className="font-medium">Break: {formatTime(stats.totalBreak)}</span>
        </div>
        
        {/* Inactive category */}
        <div className="flex items-center bg-gray-700 px-3 py-2 rounded-md">
          <div className="w-4 h-4 bg-gray-500 mr-2 rounded-sm"></div>
          <span className="font-medium">Inactive: {formatTime(stats.totalInactive)}</span>
        </div>
      </div>
    </div>
  );
};

export default MonthlyDistribution;
