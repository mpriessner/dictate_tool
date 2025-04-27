import React from 'react';
import { MODE_COLORS, formatTime } from './DailyUtils';

const DailyTimeline = ({ intervals }) => {
  const formatHour = (hour) => {
    return hour.toString().padStart(2, '0');
  };
  
  // Create hour labels
  const hourLabels = [];
  for (let h = 0; h <= 24; h += 4) {
    hourLabels.push(h);
  }

  return (
    <div className="w-full h-16 py-2">
      <div className="relative w-full h-8 bg-gray-800 rounded">
        {intervals && intervals.map((interval, index) => {
          const startHour = interval.start.getHours() + interval.start.getMinutes() / 60;
          const endHour = interval.end.getHours() + interval.end.getMinutes() / 60;
          
          return (
            <div
              key={index}
              className="absolute h-full"
              style={{
                left: `${(startHour / 24) * 100}%`,
                width: `${((endHour - startHour) / 24) * 100}%`,
                backgroundColor: MODE_COLORS[interval.mode]
              }}
              title={`${formatTime(interval.start)} - ${formatTime(interval.end)} (${interval.mode})`}
            />
          );
        })}
      </div>
      <div className="relative w-full h-6 mt-1">
        {hourLabels.map((hour) => (
          <div 
            key={hour} 
            className="absolute text-xs text-gray-400"
            style={{ left: `${(hour / 24) * 100}%`, transform: 'translateX(-50%)' }}
          >
            {formatHour(hour)}
          </div>
        ))}
      </div>
    </div>
  );
};

export default DailyTimeline;
