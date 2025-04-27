import React from 'react';
import { MODE_COLORS } from './DailyUtils';

const FocusDistribution = ({ stats }) => {
  // Calculate total time and percentages
  const totalTime = stats.work + stats.leisure + stats.break + stats.inactive;
  
  // Calculate percentages
  const workPercent = totalTime > 0 ? stats.work / totalTime : 0;
  const leisurePercent = totalTime > 0 ? stats.leisure / totalTime : 0;
  const breakPercent = totalTime > 0 ? stats.break / totalTime : 0;
  const inactivePercent = totalTime > 0 ? stats.inactive / totalTime : 0;
  
  // Calculate stroke dash values
  const circumference = 251.2; // 2 * PI * r where r = 40
  const workDash = workPercent * circumference;
  const leisureDash = leisurePercent * circumference;
  const inactiveDash = inactivePercent * circumference;
  const breakDash = breakPercent * circumference;
  
  // Calculate offsets
  const breakOffset = 0;
  const inactiveOffset = -1 * breakDash;
  const leisureOffset = -1 * (breakDash + inactiveDash);
  const workOffset = -1 * (breakDash + inactiveDash + leisureDash);

  // Format time for display
  const formatTime = (hours) => {
    const h = Math.floor(hours);
    const m = Math.round((hours - h) * 60);
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
            
            {/* Break segment first */}
            {stats.break > 0 && (
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
            
            {/* Inactive segment */}
            {stats.inactive > 0 && (
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
            
            {/* Leisure focus segment */}
            {stats.leisure > 0 && (
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
            {stats.work > 0 && (
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
              {Math.round((stats.work + stats.leisure) / totalTime * 100)}% Focus
            </text>
            <text x="50" y="57" textAnchor="middle" className="text-xxs" fill="#FFCA28" style={{ fontSize: '0.6rem' }}>
              {stats.leisure > 0 ? `${Math.round((stats.leisure / totalTime) * 100)}% Leisure` : ''}
            </text>
            <text x="50" y="72" textAnchor="middle" className="text-xxs" fill="#3B82F6" style={{ fontSize: '0.6rem' }}>
              {stats.work > 0 ? `${Math.round((stats.work / totalTime) * 100)}% Work` : ''}
            </text>
          </svg>
        </div>
      </div>
      
      {/* Legend */}
      <div className="flex justify-center mt-6 space-x-6 flex-wrap">
        {/* Work category */}
        <div className="flex items-center bg-gray-700 px-3 py-2 rounded-md">
          <div className="w-5 h-5 bg-blue-500 mr-2 rounded-sm"></div>
          <span className="font-medium">Work: {formatTime(stats.work)}</span>
        </div>
        
        {/* Leisure category */}
        <div className="flex items-center bg-gray-700 px-3 py-2 rounded-md">
          <div className="w-5 h-5 bg-amber-400 mr-2 rounded-sm"></div>
          <span className="font-medium">Leisure: {formatTime(stats.leisure)}</span>
        </div>
        
        {/* Inactive category */}
        <div className="flex items-center bg-gray-700 px-3 py-2 rounded-md">
          <div className="w-5 h-5 bg-gray-600 mr-2 rounded-sm"></div>
          <span className="font-medium">Inactive: {formatTime(stats.inactive)}</span>
        </div>
        
        {/* Break category */}
        <div className="flex items-center bg-gray-700 px-3 py-2 rounded-md">
          <div className="w-5 h-5 bg-cyan-400 mr-2 rounded-sm"></div>
          <span className="font-medium">Break: {formatTime(stats.break)}</span>
        </div>
      </div>
    </div>
  );
};

export default FocusDistribution;
