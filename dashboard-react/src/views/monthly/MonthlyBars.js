import React from 'react';
import { MODE_COLORS } from './MonthlyUtils';

const MonthlyBars = ({ weeklyStats }) => {
  console.log("MonthlyBars received weeklyStats:", weeklyStats);

  // Make sure weeklyStats is an array
  if (!weeklyStats) {
    console.error("weeklyStats is undefined or null");
    return (
      <div className="relative bg-gray-800 p-4 rounded-lg shadow-lg flex flex-col" style={{ height: '300px' }}>
        <div className="flex justify-center items-center h-full">
          <span className="text-gray-400">No weekly data available (null or undefined)</span>
        </div>
      </div>
    );
  }
  
  if (!Array.isArray(weeklyStats)) {
    console.error("weeklyStats is not an array:", typeof weeklyStats, weeklyStats);
    return (
      <div className="relative bg-gray-800 p-4 rounded-lg shadow-lg flex flex-col" style={{ height: '300px' }}>
        <div className="flex justify-center items-center h-full">
          <span className="text-gray-400">No weekly data available (not an array)</span>
        </div>
      </div>
    );
  }
  
  if (weeklyStats.length === 0) {
    console.error("weeklyStats is an empty array");
    return (
      <div className="relative bg-gray-800 p-4 rounded-lg shadow-lg flex flex-col" style={{ height: '300px' }}>
        <div className="flex justify-center items-center h-full">
          <span className="text-gray-400">No weekly data available (empty array)</span>
        </div>
      </div>
    );
  }

  // Check if all values are zero
  let allZero = true;
  for (const week of weeklyStats) {
    if ((week.work || 0) > 0 || (week.leisure || 0) > 0 || (week.break || 0) > 0 || (week.inactive || 0) > 0) {
      allZero = false;
      break;
    }
  }
  
  // Calculate max hours for Y-axis scaling
  let rawMaxHours = 20; // Default to 20 if all data is zero
  
  if (!allZero) {
    rawMaxHours = Math.max(
      ...weeklyStats.map(week => {
        const total = (week.work || 0) + (week.leisure || 0) + (week.break || 0) + (week.inactive || 0);
        return total;
      })
    );
  }
  
  // Ensure we have a sensible minimum value
  rawMaxHours = Math.max(rawMaxHours, 1);

  // Add 20% buffer to max hours
  const chartMaxY = Math.max(rawMaxHours * 1.2, 1);

  // Generate Y-axis labels
  const yAxisLabels = [];
  let yAxisStep;
  
  // Determine appropriate step size
  if (chartMaxY <= 5) {
    yAxisStep = 1;
  } else if (chartMaxY <= 10) {
    yAxisStep = 2;
  } else if (chartMaxY <= 20) {
    yAxisStep = 4;
  } else if (chartMaxY <= 50) {
    yAxisStep = 10;
  } else {
    yAxisStep = Math.ceil(chartMaxY / 5);
  }
  
  // Generate labels
  for (let i = yAxisStep; i < chartMaxY; i += yAxisStep) {
    if (i < chartMaxY - yAxisStep / 10) {
      yAxisLabels.push(i % 1 === 0 ? i.toString() : i.toFixed(1));
    }
  }

  // Calculate average focus hours per week
  const nonZeroWeeks = weeklyStats.filter(week => (week.work + week.leisure) > 0);
  const average = nonZeroWeeks.length > 0 
    ? (nonZeroWeeks.reduce((sum, week) => sum + week.work + week.leisure, 0) / nonZeroWeeks.length).toFixed(1)
    : 0;

  return (
    <div className="relative bg-gray-800 p-4 rounded-lg shadow-lg flex flex-col" style={{ height: '300px' }}>
      {/* Chart Title and Average */}
      <div className="flex justify-between items-center mb-4">
        <h3 className="text-lg font-semibold text-white">Weekly Focus</h3>
        <span className="text-sm text-gray-400">Average: {average}h focus/week</span>
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
            {weeklyStats.map((week, index) => {
              // Calculate heights as percentages
              const workHeightPercent = (week.work / chartMaxY) * 100;
              const leisureHeightPercent = (week.leisure / chartMaxY) * 100;
              const breakHeightPercent = (week.break / chartMaxY) * 100;
              const inactiveHeightPercent = (week.inactive / chartMaxY) * 100;

              // Calculate absolute heights
              const totalHeight = 220; // Pixel height of chart area
              const minBarHeight = 2; // Minimum height for visible bars
              
              const workHeight = Math.max(minBarHeight, (workHeightPercent / 100) * totalHeight);
              const leisureHeight = Math.max(minBarHeight, (leisureHeightPercent / 100) * totalHeight);
              const breakHeight = Math.max(minBarHeight, (breakHeightPercent / 100) * totalHeight);
              const inactiveHeight = Math.max(minBarHeight, (inactiveHeightPercent / 100) * totalHeight);
              
              // Format date range for the week
              const weekStart = new Date(week.start);
              // Create a new date object for the end date to avoid mutation issues
              const weekEnd = new Date(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate() + 6);
              
              const formatWeekRange = () => {
                try {
                  if (weekStart.getMonth() === weekEnd.getMonth()) {
                    // Same month
                    return `${weekStart.getDate()}-${weekEnd.getDate()} ${weekStart.toLocaleDateString('en-US', { month: 'short' })}`;
                  } else {
                    // Different months
                    return `${weekStart.getDate()} ${weekStart.toLocaleDateString('en-US', { month: 'short' })} - ${weekEnd.getDate()} ${weekEnd.toLocaleDateString('en-US', { month: 'short' })}`;
                  }
                } catch (error) {
                  console.error("Error formatting date range:", error, weekStart, weekEnd);
                  return `Week ${index + 1}`;
                }
              };
              
              return (
                <div key={index} className="flex flex-col justify-end h-full" style={{ width: `${100 / weeklyStats.length}%` }}>
                  <div className="w-12 mx-auto flex flex-col-reverse">
                    {/* Stack bars with improved visibility and spacing */}
                    <div className="relative" style={{ height: `${workHeight + leisureHeight + breakHeight + inactiveHeight}px` }}>
                      {/* Work focus time (bottom) */}
                      {week.work > 0 && (
                        <div 
                          className="absolute bottom-0 w-full bg-blue-500 rounded-sm hover:opacity-90 transition-opacity"
                          style={{ height: `${workHeight}px` }}
                          title={`Work: ${week.work.toFixed(1)}h`}
                        />
                      )}
                      
                      {/* Leisure focus time (middle) */}
                      {week.leisure > 0 && (
                        <div 
                          className="absolute w-full bg-amber-400 rounded-sm hover:opacity-90 transition-opacity"
                          style={{ 
                            height: `${leisureHeight}px`,
                            bottom: `${workHeight}px`
                          }}
                          title={`Leisure: ${week.leisure.toFixed(1)}h`}
                        />
                      )}
                      
                      {/* Break time */}
                      {week.break > 0 && (
                        <div 
                          className="absolute w-full bg-cyan-400 rounded-sm hover:opacity-90 transition-opacity"
                          style={{ 
                            height: `${breakHeight}px`,
                            bottom: `${workHeight + leisureHeight}px`
                          }}
                          title={`Break: ${week.break.toFixed(1)}h`}
                        />
                      )}
                      
                      {/* Inactive time (top) */}
                      {week.inactive > 0 && (
                        <div 
                          className="absolute w-full bg-gray-400 rounded-sm hover:opacity-90 transition-opacity"
                          style={{ 
                            height: `${inactiveHeight}px`,
                            bottom: `${workHeight + leisureHeight + breakHeight}px`
                          }}
                          title={`Inactive: ${week.inactive.toFixed(1)}h`}
                        />
                      )}
                    </div>
                  </div>
                  {/* Week label */}
                  <div className="text-center text-xs text-gray-400 mt-2">
                    <div>Week {weekStart.getDate() === 1 || index === 0 ? weekStart.toLocaleDateString('en-US', { month: 'short' }) : index + 1}</div>
                    <div>{formatWeekRange()}</div>
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

export default MonthlyBars;