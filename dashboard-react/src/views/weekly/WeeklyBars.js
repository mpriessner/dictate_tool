import React from 'react';
import { MODE_COLORS, formatDate } from './WeeklyUtils';

const WeeklyBars = ({ dailyStats }) => {
  // Calculate max hours for Y-axis scaling
  const rawMaxHours = Math.max(
    ...dailyStats.map(day => {
      const total = day.work + day.leisure + day.break + day.inactive;
      return total;
    })
  ) || 8; // Default to 8 if no data

  // Add 20% buffer to max hours
  const chartMaxY = Math.max(rawMaxHours * 1.2, 1);

  // Generate Y-axis labels
  const yAxisLabels = [];
  let yAxisStep;
  
  // Determine appropriate step size
  if (chartMaxY <= 1) {
    yAxisStep = 0.2;
  } else if (chartMaxY <= 2) {
    yAxisStep = 0.5;
  } else if (chartMaxY <= 5) {
    yAxisStep = 1;
  } else if (chartMaxY <= 10) {
    yAxisStep = 2;
  } else {
    yAxisStep = Math.ceil(chartMaxY / 5);
  }
  
  // Generate labels
  for (let i = yAxisStep; i < chartMaxY; i += yAxisStep) {
    if (i < chartMaxY - yAxisStep / 10) {
      yAxisLabels.push(i % 1 === 0 ? i.toString() : i.toFixed(1));
    }
  }

  // Calculate average focus hours
  const nonZeroDays = dailyStats.filter(day => (day.work + day.leisure) > 0);
  const average = nonZeroDays.length > 0 
    ? (nonZeroDays.reduce((sum, day) => sum + day.work + day.leisure, 0) / nonZeroDays.length).toFixed(1)
    : 0;

  return (
    <div className="relative bg-gray-800 p-4 rounded-lg shadow-lg flex flex-col" style={{ height: '300px' }}>
      {/* Chart Title and Average */}
      <div className="flex justify-between items-center mb-4">
        <h3 className="text-lg font-semibold text-white">Daily Focus</h3>
        <span className="text-sm text-gray-400">Average: {average}h focus/day</span>
      </div>

      {/* Main Chart Area */}
      <div className="flex-1 flex mb-6">
        {/* Y-axis labels */}
        <div className="flex flex-col justify-between w-10 text-right pr-2 text-xs text-gray-400">
          <div className="h-0 relative">
            <span className="absolute -top-1.5 right-2">{chartMaxY.toFixed(1)}h</span>
          </div>
          {yAxisLabels.slice().reverse().map((label, index) => (
            <div key={index} className="h-0 relative">
              <span className="absolute -top-1.5 right-2">{label}h</span>
            </div>
          ))}
          <div className="h-0 relative">
            <span className="absolute -top-1.5 right-2">0h</span>
          </div>
        </div>

        {/* Chart Grid and Bars */}
        <div className="flex-1 relative">
          {/* Grid lines */}
          <div className="absolute inset-0 flex flex-col justify-between pointer-events-none">
            <hr className="border-t border-gray-700" />
            {yAxisLabels.map((_, index) => (
              <hr key={index} className="border-t border-gray-700" />
            ))}
            <hr className="border-t border-gray-700" />
          </div>

          {/* Bars */}
          <div className="absolute inset-0 flex justify-around">
            {dailyStats.map((day, index) => {
              // Calculate heights as percentages
              const workHeightPercent = (day.work / chartMaxY) * 100;
              const leisureHeightPercent = (day.leisure / chartMaxY) * 100;
              const breakHeightPercent = (day.break / chartMaxY) * 100;
              const inactiveHeightPercent = (day.inactive / chartMaxY) * 100;

              // Calculate absolute heights
              const totalHeight = 220; // Pixel height of chart area
              const minBarHeight = 2; // Minimum height for visible bars
              
              const workHeight = Math.max(minBarHeight, (workHeightPercent / 100) * totalHeight);
              const leisureHeight = Math.max(minBarHeight, (leisureHeightPercent / 100) * totalHeight);
              const breakHeight = Math.max(minBarHeight, (breakHeightPercent / 100) * totalHeight);
              const inactiveHeight = Math.max(minBarHeight, (inactiveHeightPercent / 100) * totalHeight);
              
              return (
                <div key={index} className="flex flex-col justify-end h-full" style={{ width: '14%' }}>
                  <div className="w-8 mx-auto flex flex-col-reverse">
                    {/* Stack bars with improved visibility and spacing */}
                    <div className="relative" style={{ height: `${workHeight + leisureHeight + breakHeight + inactiveHeight}px` }}>
                      {/* Work focus time (bottom) */}
                      {day.work > 0 && (
                        <div 
                          className="absolute bottom-0 w-full bg-blue-500 rounded-sm hover:opacity-90 transition-opacity"
                          style={{ height: `${workHeight}px` }}
                          title={`Work: ${day.work.toFixed(1)}h`}
                        />
                      )}
                      
                      {/* Leisure focus time (middle) */}
                      {day.leisure > 0 && (
                        <div 
                          className="absolute w-full bg-amber-400 rounded-sm hover:opacity-90 transition-opacity"
                          style={{ 
                            height: `${leisureHeight}px`,
                            bottom: `${workHeight}px`
                          }}
                          title={`Leisure: ${day.leisure.toFixed(1)}h`}
                        />
                      )}
                      
                      {/* Break time */}
                      {day.break > 0 && (
                        <div 
                          className="absolute w-full bg-cyan-400 rounded-sm hover:opacity-90 transition-opacity"
                          style={{ 
                            height: `${breakHeight}px`,
                            bottom: `${workHeight + leisureHeight}px`
                          }}
                          title={`Break: ${day.break.toFixed(1)}h`}
                        />
                      )}
                      
                      {/* Inactive time (top) */}
                      {day.inactive > 0 && (
                        <div 
                          className="absolute w-full bg-gray-400 rounded-sm hover:opacity-90 transition-opacity"
                          style={{ 
                            height: `${inactiveHeight}px`,
                            bottom: `${workHeight + leisureHeight + breakHeight}px`
                          }}
                          title={`Inactive: ${day.inactive.toFixed(1)}h`}
                        />
                      )}
                    </div>
                  </div>
                  {/* Day label */}
                  <div className="text-center text-xs text-gray-400 mt-2">
                    <div>{new Date(day.date).toLocaleDateString('en-US', { weekday: 'short' })}</div>
                    <div>{formatDate(day.date)}</div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* Legend */}
      <div className="flex justify-center space-x-4 mt-3 text-xs text-gray-400">
        <div className="flex items-center space-x-2">
          <span className="w-3 h-3 bg-blue-500 rounded-sm"></span>
          <span>Work</span>
          <span className="w-3 h-3 bg-amber-400 rounded-sm ml-2"></span>
          <span>Leisure</span>
          <span className="w-3 h-3 bg-cyan-400 rounded-sm ml-2"></span>
          <span>Break</span>
          <span className="w-3 h-3 bg-gray-400 rounded-sm ml-2"></span>
          <span>Inactive</span>
        </div>
      </div>
    </div>
  );
};

export default WeeklyBars;
