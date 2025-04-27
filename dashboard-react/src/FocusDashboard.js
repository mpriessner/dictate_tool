import React, { useState, useEffect } from 'react';

// Get Monday of the current week
const getMondayOfWeek = (d) => {
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1); // adjust when day is Sunday
  return new Date(d.setDate(diff));
};

// Timeline Component
const Timeline = ({ intervals, modeColors }) => {
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
                backgroundColor: modeColors[interval.mode]
              }}
              title={`${interval.start.toLocaleTimeString()} - ${interval.end.toLocaleTimeString()} (${interval.mode})`}
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

// Weekly View Component with Fixed Bar Display
const WeeklyView = ({ dailyHours }) => {
  // Make sure dailyHours is an array
  const days = Array.isArray(dailyHours) ? dailyHours : [];
  console.log('WeeklyView received dailyHours:', dailyHours);
  
  if (days.length === 0) {
    console.warn('No days data to render in WeeklyView');
    return <div className="text-center p-8">No data available for this week.</div>;
  }
  
  // Calculate max hours for Y-axis scaling
  const rawMaxHours = days.length > 0
    ? Math.max(...days.map(day => {
        const workHours = parseFloat(day.focusWorkHours) || 0;
        const leisureHours = parseFloat(day.focusLeisureHours) || 0;
        const breakHours = parseFloat(day.breakHours) || 0;
        console.log(`Day data: work=${workHours}h, leisure=${leisureHours}h, break=${breakHours}h`);
        const total = workHours + leisureHours + breakHours;
        console.log(`Total hours: ${total}h`);
        return total;
      }))
    : 0;
  console.log('Raw max hours:', rawMaxHours);
    
  // Calculate average focus hours from days with data
  const nonZeroDays = days.filter(day => day.focusHours > 0);
  const average = nonZeroDays.length > 0 
    ? (nonZeroDays.reduce((sum, day) => sum + day.focusHours, 0) / nonZeroDays.length).toFixed(1)
    : 0;

  // Set a better maximum for the Y-axis that adapts to the data
  // Add just a small buffer (20%) above the maximum value
  let chartMaxY = rawMaxHours <= 0 ? 1 : rawMaxHours * 1.2;
  
  // Set minimum scale to at least 1 hour if data is very small
  chartMaxY = Math.max(chartMaxY, 1);
  
  console.log('Using chartMaxY for scaling:', chartMaxY);

  // Generate Y-axis labels with appropriate step size
  const yAxisLabels = [];
  let yAxisStep;
  
  // Determine appropriate step size based on the max value
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
  
  // Generate labels at regular intervals
  for (let i = yAxisStep; i < chartMaxY; i += yAxisStep) {
    // Avoid floating point issues
    if (i < chartMaxY - yAxisStep / 10) {
      // Format as integer if it's a whole number
      yAxisLabels.push(i % 1 === 0 ? i.toString() : i.toFixed(1));
    }
  }
  
  console.log('Y-axis labels:', yAxisLabels, 'with step size:', yAxisStep);

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

          {/* Bars - FIXED: Changed display approach */}
          <div className="absolute inset-0 flex justify-around">
            {days.map((day, index) => {
              // Convert minutes to hours
              const workHours = (parseFloat(day.focusWorkMinutes) || 0) / 60;
              const leisureHours = (parseFloat(day.focusLeisureMinutes) || 0) / 60;
              const breakHours = (parseFloat(day.breakMinutes) || 0) / 60;
              const inactiveHours = (parseFloat(day.inactive) || 0) / 60;
              
              // Calculate heights as percentages of chartMaxY
              const workHeightPercent = chartMaxY > 0 ? (workHours / chartMaxY) * 100 : 0;
              const leisureHeightPercent = chartMaxY > 0 ? (leisureHours / chartMaxY) * 100 : 0;
              const breakHeightPercent = chartMaxY > 0 ? (breakHours / chartMaxY) * 100 : 0;
              const inactiveHeightPercent = chartMaxY > 0 ? (inactiveHours / chartMaxY) * 100 : 0;
              
              console.log(`Bar heights for ${day.day}: work=${workHeightPercent.toFixed(1)}%, leisure=${leisureHeightPercent.toFixed(1)}%, break=${breakHeightPercent.toFixed(1)}%, inactive=${inactiveHeightPercent.toFixed(1)}%`);

              // Calculate absolute heights with better minimum values
              const totalHeight = 220; // Pixel height of chart area
              const minBarHeight = 2; // Minimum height for visible bars
              
              // Calculate proportional heights
              const workHeight = Math.max(minBarHeight, (workHeightPercent / 100) * totalHeight);
              const leisureHeight = Math.max(minBarHeight, (leisureHeightPercent / 100) * totalHeight);
              const breakHeight = Math.max(minBarHeight, (breakHeightPercent / 100) * totalHeight);
              const inactiveHeight = Math.max(minBarHeight, (inactiveHeightPercent / 100) * totalHeight);
              
              // Calculate stack height ensuring proper spacing
              const stackHeight = workHeight + leisureHeight + breakHeight;
              
              console.log(`Bar for ${day.day}: work=${workHours}h (${workHeight.toFixed(1)}px), ` +
                          `leisure=${leisureHours}h (${leisureHeight.toFixed(1)}px), ` +
                          `break=${breakHours}h (${breakHeight.toFixed(1)}px), ` +
                          `total=${stackHeight.toFixed(1)}px`);

              return (
                <div key={index} className="flex flex-col justify-end h-full" style={{ width: '14%' }}>
                  {/* FIXED: Stack the different bar types from bottom up */}
                  <div className="w-8 mx-auto flex flex-col-reverse">
                    {/* Stack bars with improved visibility and spacing */}
                    <div className="relative" style={{ height: `${stackHeight}px` }}>
                      {/* Work focus time (bottom) */}
                      {workHours > 0 && (
                        <div 
                          className="absolute bottom-0 w-full bg-blue-500 rounded-sm hover:opacity-90 transition-opacity"
                          style={{ height: `${workHeight}px` }}
                          title={`Work: ${workHours.toFixed(1)}h`}
                        />
                      )}
                      
                      {/* Leisure focus time (middle) */}
                      {leisureHours > 0 && (
                        <div 
                          className="absolute w-full bg-amber-400 rounded-sm hover:opacity-90 transition-opacity"
                          style={{ 
                            height: `${leisureHeight}px`,
                            bottom: `${workHeight}px`
                          }}
                          title={`Leisure: ${leisureHours.toFixed(1)}h`}
                        />
                      )}
                      
                      {/* Break time */}
                      {breakHours > 0 && (
                        <div 
                          className="absolute w-full bg-cyan-400 rounded-sm hover:opacity-90 transition-opacity"
                          style={{ 
                            height: `${breakHeight}px`,
                            bottom: `${workHeight + leisureHeight}px`
                          }}
                          title={`Break: ${breakHours.toFixed(1)}h`}
                        />
                      )}
                      
                      {/* Inactive time (top) */}
                      {inactiveHours > 0 && (
                        <div 
                          className="absolute w-full bg-gray-400 rounded-sm hover:opacity-90 transition-opacity"
                          style={{ 
                            height: `${inactiveHeight}px`,
                            bottom: `${workHeight + leisureHeight + breakHeight}px`
                          }}
                          title={`Inactive: ${inactiveHours.toFixed(1)}h`}
                        />
                      )}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* X-axis labels */}
      <div className="flex justify-around pl-10 pr-0 -mt-4 text-xs text-gray-400">
        {days.map((day, index) => (
          <div key={index} className="flex flex-col items-center text-center" style={{ width: '14%' }}>
            <div>{day.day}</div>
            <div>{day.date}</div>
          </div>
        ))}
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

// Monthly View Component
const MonthlyView = ({ weeklyData }) => {
  console.log('MonthlyView weeklyData:', weeklyData);
  return (
    <div className="w-full">
      {weeklyData.map((week, index) => {
        // Ensure we have valid numbers
        const workMinutes = typeof week.focusWorkMinutes === 'number' ? week.focusWorkMinutes : 0;
        const leisureMinutes = typeof week.focusLeisureMinutes === 'number' ? week.focusLeisureMinutes : 0;
        const breakMinutes = typeof week.breakMinutes === 'number' ? week.breakMinutes : 0;
        const totalMinutes = week.totalMinutes || (workMinutes + leisureMinutes + breakMinutes);
        
        // Calculate percentages for bar widths
        const workPercentage = totalMinutes > 0 ? (workMinutes / totalMinutes) * 100 : 0;
        const leisurePercentage = totalMinutes > 0 ? (leisureMinutes / totalMinutes) * 100 : 0;
        const breakPercentage = totalMinutes > 0 ? (breakMinutes / totalMinutes) * 100 : 0;
        
        // Calculate positions for stacked bars
        const leisureStart = workPercentage;
        const breakStart = workPercentage + leisurePercentage;
        
        return (
          <div key={index} className="flex items-center bg-gray-700 rounded p-3 mb-2">
            <div className="flex-shrink-0 w-16 text-center">
              <div className="text-sm font-medium">{week.week}</div>
              <div className="text-xs text-gray-300">{week.dateRange}</div>
            </div>
            <div className="flex-grow mx-4">
              <div className="w-full h-6 bg-gray-800 rounded-full overflow-hidden relative">
                {/* Work focus time bar */}
                {workPercentage > 0 && (
                  <div 
                    className="h-6 bg-blue-500 absolute left-0" 
                    style={{ width: `${workPercentage}%` }}
                    title={`Work: ${Math.floor(workMinutes / 60)}h ${workMinutes % 60}m`}
                  />
                )}
                
                {/* Leisure focus time bar */}
                {leisurePercentage > 0 && (
                  <div 
                    className="h-6 bg-amber-400 absolute" 
                    style={{ 
                      width: `${leisurePercentage}%`,
                      left: `${leisureStart}%`
                    }}
                    title={`Leisure: ${Math.floor(leisureMinutes / 60)}h ${leisureMinutes % 60}m`}
                  />
                )}
                
                {/* Break time bar */}
                {breakPercentage > 0 && (
                  <div 
                    className="h-6 bg-cyan-400 absolute" 
                    style={{ 
                      width: `${breakPercentage}%`,
                      left: `${breakStart}%`
                    }}
                    title={`Break: ${Math.floor(breakMinutes / 60)}h ${breakMinutes % 60}m`}
                  />
                )}
              </div>
            </div>
            <div className="flex-shrink-0 text-right">
              <div className="text-sm">{Math.floor(totalMinutes / 60)}h {totalMinutes % 60}m</div>
              <div className="text-xs text-gray-300">
                <span className="text-blue-400">{Math.round(workPercentage)}%</span> work
                {leisurePercentage > 0 && (
                  <span>, <span className="text-amber-400">{Math.round(leisurePercentage)}%</span> leisure</span>
                )}
              </div>
            </div>
          </div>
        );
      })}
      
      {/* Legend */}
      <div className="flex items-center space-x-4 mt-4 justify-center">
        <div className="flex items-center">
          <div className="w-4 h-4 bg-blue-500 mr-1"></div>
          <span className="text-xs">Work</span>
        </div>
        <div className="flex items-center">
          <div className="w-4 h-4 bg-amber-400 mr-1"></div>
          <span className="text-xs">Leisure</span>
        </div>
        <div className="flex items-center">
          <div className="w-4 h-4 bg-cyan-400 mr-1"></div>
          <span className="text-xs">Break</span>
        </div>
      </div>
    </div>
  );
};

// App Usage Component
const AppUsage = ({ appUsage }) => {
  const [showAll, setShowAll] = useState(false);
  
  // Sort apps by usage percentage (descending)
  const sortedApps = Object.entries(appUsage || {}).sort((a, b) => b[1] - a[1]);
  
  // Determine which apps to display
  const displayApps = showAll ? sortedApps : sortedApps.slice(0, 5);
  
  // Check if there are more than 5 apps
  const hasMore = sortedApps.length > 5;
  
  // Function to get a color based on the app name
  const getAppColor = (appName) => {
    const colors = {
      'Claude': '#EB5757',
      'Electron': '#F2994A',
      'python': '#6FCF97',
      'Google Chrome': '#2F80ED',
      'Arc': '#9B51E0',
      'Finder': '#56CCF2',
      'Rize': '#BB6BD9',
      'Web App': '#27AE60',
      'Notes': '#F2C94C',
      'Preview': '#828282'
    };
    
    return colors[appName] || '#BDBDBD';
  };
  
  return (
    <div className="w-full">
      {displayApps.map(([app, percentage], index) => (
        <div key={index} className="mb-2">
          <div className="flex justify-between items-center mb-1">
            <span className="text-sm">{app}</span>
            <span className="text-sm">{percentage}%</span>
          </div>
          <div className="w-full h-2 bg-gray-700 rounded-full">
            <div 
              className="h-2 rounded-full" 
              style={{ 
                width: `${percentage}%`,
                backgroundColor: getAppColor(app)
              }}
            />
          </div>
        </div>
      ))}
      
      {hasMore && (
        <button 
          className="mt-2 px-3 py-1 bg-gray-700 hover:bg-gray-600 rounded text-cyan-400 text-sm flex items-center"
          onClick={() => setShowAll(!showAll)}
        >
          {showAll ? (
            <>
              <span>Show Less</span>
              <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4 ml-1" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 15l7-7 7 7" />
              </svg>
            </>
          ) : (
            <>
              <span>Show More</span>
              <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4 ml-1" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
              </svg>
            </>
          )}
        </button>
      )}
    </div>
  );
};

const FocusDashboard = () => {
  const [mode, setMode] = useState('day');
  const [data, setData] = useState(null);
  const [date, setDate] = useState(new Date());
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [pyHandlerReady, setPyHandlerReady] = useState(false);
  const [showDropdown, setShowDropdown] = useState(false);

  // Activity mode colors
  const MODE_COLORS = {
    work: '#2196F3',    // Blue
    leisure: '#FFCA28', // Yellow/Amber
    break: '#00BCD4',   // Cyan
    inactive: '#9E9E9E' // Gray
  };

  // Listen for the pyHandlerReady event
  useEffect(() => {
    console.log("Setting up pyHandlerReady listener");
    
    // Check if pyHandler is already available
    if (window.pyHandler) {
      console.log("pyHandler is already available");
      setPyHandlerReady(true);
    }
    
    // Function to handle the pyHandlerReady event
    const handlePyHandlerReady = () => {
      console.log("pyHandlerReady event received");
      setPyHandlerReady(true);
    };
    
    // Listen for the qtReady event (dispatched when QWebChannel is initialized)
    const handleQtReady = () => {
      console.log("qtReady event received");
      if (window.pyHandler) {
        console.log("pyHandler is available after qtReady");
        setPyHandlerReady(true);
      }
    };
    
    // Add event listeners
    document.addEventListener('pyHandlerReady', handlePyHandlerReady);
    document.addEventListener('qtReady', handleQtReady);
    
    // Check again after a short delay
    const timeoutId = setTimeout(() => {
      console.log("Delayed check for pyHandler");
      if (window.pyHandler) {
        console.log("pyHandler is available after timeout");
        setPyHandlerReady(true);
      }
    }, 1000);
    
    // Cleanup
    return () => {
      document.removeEventListener('pyHandlerReady', handlePyHandlerReady);
      document.removeEventListener('qtReady', handleQtReady);
      clearTimeout(timeoutId);
    };
  }, []);

  // Fetch data when mode or date changes, but only if pyHandler is ready
  useEffect(() => {
    if (!pyHandlerReady) {
      console.log("pyHandler not ready yet, skipping data fetch");
      return;
    }

    const fetchData = async () => {
      setLoading(true);
      setError(null);
      
      try {
        if (!window.pyHandler) {
          throw new Error("Python handler not available");
        }

        const result = await window.pyHandler.get_data(
          mode,
          date.toISOString().split('T')[0]
        );

        if (!result || Object.keys(result).length === 0) {
          throw new Error("No data available for the selected period");
        }

        console.log('Data received from Python:', result);
        console.log('App usage data:', result.app_usage);
        console.log('Raw app times:', result.raw_app_times);
        console.log('Focus Work:', result.focusWork, 'Focus Leisure:', result.focusLeisure, 'Break:', result.break);
        
        // Convert interval ISO strings to Date objects
        if (result.intervals && Array.isArray(result.intervals)) {
          // Process intervals with proper mode handling and inactive detection
          const processedIntervals = [];
          let lastEndTime = new Date(date);
          lastEndTime.setHours(0, 0, 0, 0);

          // Sort intervals by start time
          const sortedIntervals = result.intervals.sort((a, b) => 
            new Date(a.start) - new Date(b.start)
          );

          // Process each interval
          sortedIntervals.forEach(interval => {
            const startTime = new Date(interval.start);
            const endTime = new Date(interval.end);

            // Add inactive period if gap > 1 hour
            if (startTime - lastEndTime > 3600000) { // 1 hour in milliseconds
              processedIntervals.push({
                start: new Date(lastEndTime),
                end: new Date(startTime),
                mode: 'inactive'
              });
            }

            // Add the actual activity interval
            processedIntervals.push({
              start: startTime,
              end: endTime,
              mode: interval.active === 1 ? 'work' : 
                    interval.active === 2 ? 'leisure' : 
                    interval.active === 0 ? 'break' : 'inactive'
            });

            lastEndTime = endTime;
          });

          // Add final inactive period if needed
          const dayEnd = new Date(date);
          dayEnd.setHours(23, 59, 59, 999);
          if (dayEnd - lastEndTime > 3600000) {
            processedIntervals.push({
              start: new Date(lastEndTime),
              end: dayEnd,
              mode: 'inactive'
            });
          }

          result.intervals = processedIntervals;
          console.log('Processed intervals:', processedIntervals);
        }
        
        setData(result);
      } catch (err) {
        console.error("Error fetching data:", err);
        setError(err.message);
        setData(null);
      } finally {
        setLoading(false);
      }
    };

    fetchData();
  }, [mode, date, pyHandlerReady]);

  // Handle date changes
  const changeDate = (offset) => {
    const newDate = new Date(date);
    if (mode === 'day') {
      newDate.setDate(newDate.getDate() + offset);
    } else if (mode === 'week') {
      newDate.setDate(newDate.getDate() + (7 * offset));
    } else {
      newDate.setMonth(newDate.getMonth() + offset);
    }

    if (newDate <= new Date()) {
      setDate(newDate);
    }
  };

  if (loading) {
    return <div className="flex items-center justify-center h-screen">Loading...</div>;
  }

  if (error) {
    return (
      <div className="flex items-center justify-center h-screen text-red-500">
        Error: {error}
      </div>
    );
  }

  if (!data) {
    return (
      <div className="flex items-center justify-center h-screen">
        No data available for the selected period
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-900 text-white p-6">
      {/* Header with mode selector and date navigation */}
      <div className="flex justify-between items-center mb-8">
        <div className="flex space-x-4 relative">
          <button 
            className={`px-4 py-2 rounded-md ${mode === 'day' ? 'bg-blue-500' : 'bg-gray-700'}`}
            onClick={() => {
              setMode('day');
              setShowDropdown(false);
            }}
          >
            Day
          </button>
          <button 
            className={`px-4 py-2 rounded-md ${mode === 'week' ? 'bg-blue-500' : 'bg-gray-700'}`}
            onClick={() => {
              setMode('week');
              setShowDropdown(false);
            }}
          >
            Week
          </button>
          <button 
            className={`px-4 py-2 rounded-md ${mode === 'month' ? 'bg-blue-500' : 'bg-gray-700'}`}
            onClick={() => {
              setMode('month');
              setShowDropdown(false);
            }}
          >
            Month
          </button>
          <div className="relative">
            <button 
              className="px-4 py-2 rounded-md bg-gray-700 hover:bg-gray-600"
              onClick={() => setShowDropdown(!showDropdown)}
            >
              Menu
            </button>
            {showDropdown && (
              <div className="absolute left-0 mt-2 w-48 rounded-md shadow-lg bg-gray-800 ring-1 ring-black ring-opacity-5 z-50">
                <div className="py-1" role="menu" aria-orientation="vertical">
                  <button
                    className="block w-full text-left px-4 py-2 text-sm text-gray-300 hover:bg-gray-700"
                    role="menuitem"
                    onClick={() => {
                      console.log('Add Manual Log clicked - functionality to be implemented');
                      setShowDropdown(false);
                    }}
                  >
                    Add Manual Log
                  </button>
                  {/* Add more menu items here as needed */}
                </div>
              </div>
            )}
          </div>
        </div>
        
        <div className="flex items-center space-x-4">
          <button 
            className="p-2 bg-gray-700 rounded-md"
            onClick={() => changeDate(-1)}
          >
            &lt;
          </button>
          <div className="text-lg">
            {mode === 'day' && date.toLocaleDateString()}
            {mode === 'week' && `Week of ${date.toLocaleDateString()}`}
            {mode === 'month' && date.toLocaleDateString(undefined, { month: 'long', year: 'numeric' })}
          </div>
          <button 
            className="p-2 bg-gray-700 rounded-md"
            onClick={() => changeDate(1)}
          >
            &gt;
          </button>
        </div>
      </div>

      {/* Main content */}
      {/* Calendar view for day mode */}
      {mode === 'day' && data.intervals && (
        <div className="p-4 mb-4 bg-gray-800 rounded-lg">
          <h2 className="text-lg font-bold mb-2">Calendar (sessions)</h2>
          <Timeline intervals={data.intervals} modeColors={MODE_COLORS} />
        </div>
      )}
      
      <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
        {/* Focus vs Break donut chart */}
        <div className="bg-gray-800 p-6 rounded-lg">
          <h2 className="text-xl font-semibold mb-4">Focus Distribution</h2>
          <div className="flex justify-center">
            <div className="relative w-64 h-64">
              {/* Render donut chart */}
              <svg viewBox="0 0 100 100" className="w-full h-full">
                {/* Background circle */}
                <circle cx="50" cy="50" r="40" fill="transparent" stroke="#333" strokeWidth="20" />
                
                {/* Calculate total time and percentages */}
                {(() => {
                  // Calculate total time based on view mode
                  const showInactive = mode === 'day';
                  const totalTime = showInactive
                    ? (data.focusWork || 0) + (data.focusLeisure || 0) + (data.inactive || 0) + data.break
                    : (data.focusWork || 0) + (data.focusLeisure || 0) + data.break;
                  
                  // Calculate percentages
                  const workPercent = totalTime > 0 ? (data.focusWork || 0) / totalTime : 0;
                  const leisurePercent = totalTime > 0 ? (data.focusLeisure || 0) / totalTime : 0;
                  const inactivePercent = showInactive && totalTime > 0 ? (data.inactive || 0) / totalTime : 0;
                  const breakPercent = totalTime > 0 ? data.break / totalTime : 0;
                  
                  // Calculate stroke dash values
                  const circumference = 251.2; // 2 * PI * r where r = 40
                  const workDash = workPercent * circumference;
                  const leisureDash = leisurePercent * circumference;
                  const inactiveDash = inactivePercent * circumference;
                  const breakDash = breakPercent * circumference;
                  
                  // Calculate offsets based on view mode
                  const breakOffset = 0;
                  const inactiveOffset = showInactive ? -1 * breakDash : 0;
                  const leisureOffset = showInactive 
                    ? -1 * (breakDash + inactiveDash)
                    : -1 * breakDash;
                  const workOffset = showInactive 
                    ? -1 * (breakDash + inactiveDash + leisureDash)
                    : -1 * (breakDash + leisureDash);
                  
                  return (
                    <>
                      {/* Break segment first */}
                      {data.break > 0 && (
                        <circle 
                          cx="50" 
                          cy="50" 
                          r="40" 
                          fill="transparent" 
                          stroke="#74C7EC" 
                          strokeWidth="20"
                          strokeDasharray={`${breakDash} ${circumference}`}
                          strokeDashoffset={breakOffset}
                          transform="rotate(-90 50 50)"
                        />
                      )}
                      
                      {/* Inactive segment - only in day view */}
                      {mode === 'day' && data.inactive > 0 && (
                        <circle 
                          cx="50" 
                          cy="50" 
                          r="40" 
                          fill="transparent" 
                          stroke="#4B5563" 
                          strokeWidth="20"
                          strokeDasharray={`${inactiveDash} ${circumference}`}
                          strokeDashoffset={inactiveOffset}
                          transform="rotate(-90 50 50)"
                        />
                      )}
                      
                      {/* Leisure focus segment */}
                      {data.focusLeisure > 0 && (
                        <circle 
                          cx="50" 
                          cy="50" 
                          r="40" 
                          fill="transparent" 
                          stroke="#FFCA28" 
                          strokeWidth="20"
                          strokeDasharray={`${leisureDash} ${circumference}`}
                          strokeDashoffset={leisureOffset}
                          transform="rotate(-90 50 50)"
                        />
                      )}
                      
                      {/* Work focus segment */}
                      {data.focusWork > 0 && (
                        <circle 
                          cx="50" 
                          cy="50" 
                          r="40" 
                          fill="transparent" 
                          stroke="#3B82F6" 
                          strokeWidth="20"
                          strokeDasharray={`${workDash} ${circumference}`}
                          strokeDashoffset={workOffset}
                          transform="rotate(-90 50 50)"
                        />
                      )}
                      
                      {/* Center text - show detailed breakdown */}
                      <text x="50" y="42" textAnchor="middle" className="text-xxs" fill="white" style={{ fontSize: '0.6rem' }}>
                        {Math.round(((data.focusWork || 0) + (data.focusLeisure || 0)) / totalTime * 100)}% Focus
                      </text>
                      <text x="50" y="57" textAnchor="middle" className="text-xxs" fill="#FFCA28" style={{ fontSize: '0.6rem' }}>
                        {data.focusLeisure > 0 ? `${Math.round((data.focusLeisure / totalTime) * 100)}% Leisure` : ''}
                      </text>
                      <text x="50" y="72" textAnchor="middle" className="text-xxs" fill="#3B82F6" style={{ fontSize: '0.6rem' }}>
                        {data.focusWork > 0 ? `${Math.round((data.focusWork / totalTime) * 100)}% Work` : ''}
                      </text>
                    </>
                  );
                })()}
              </svg>
            </div>
          </div>
          
          {/* Legend */}
          <div className="flex justify-center mt-6 space-x-6 flex-wrap">
            {/* Work category */}
            <div className="flex items-center bg-gray-700 px-3 py-2 rounded-md">
              <div className="w-5 h-5 bg-blue-500 mr-2 rounded-sm"></div>
              <span className="font-medium">Work: {Math.floor((data.focusWork || 0) / 60)}h {(data.focusWork || 0) % 60}m</span>
            </div>
            
            {/* Leisure category */}
            <div className="flex items-center bg-gray-700 px-3 py-2 rounded-md">
              <div className="w-5 h-5 bg-amber-400 mr-2 rounded-sm"></div>
              <span className="font-medium">Leisure: {Math.floor((data.focusLeisure || 0) / 60)}h {(data.focusLeisure || 0) % 60}m</span>
            </div>
            
            {/* Inactive category - only show in day view */}
            {mode === 'day' && (
              <div className="flex items-center bg-gray-700 px-3 py-2 rounded-md">
                <div className="w-5 h-5 bg-gray-600 mr-2 rounded-sm"></div>
                <span className="font-medium">Inactive: {Math.floor((data.inactive || 0) / 60)}h {(data.inactive || 0) % 60}m</span>
              </div>
            )}
            
            {/* Break category */}
            <div className="flex items-center bg-gray-700 px-3 py-2 rounded-md">
              <div className="w-5 h-5 bg-cyan-400 mr-2 rounded-sm"></div>
              <span className="font-medium">Break: {Math.floor(data.break / 60)}h {data.break % 60}m</span>
            </div>
          </div>
        </div>
        
        {/* App usage breakdown */}
        <div className="bg-gray-800 p-6 rounded-lg">
          <h2 className="text-xl font-semibold mb-4">App Usage</h2>
          <AppUsage appUsage={data.app_usage || {}} />
        </div>
      </div>
      
      {/* Weekly view */}
      {mode === 'week' && data.daily_hours && (
        <div className="p-4 mb-4 bg-gray-800 rounded-lg">
          <WeeklyView dailyHours={data.daily_hours} />
        </div>
      )}
      
      {/* Monthly view */}
      {mode === 'month' && data.weekly_hours && (
        <div className="mt-8 bg-gray-800 p-6 rounded-lg">
          <h2 className="text-xl font-semibold mb-4">
            Weekly Breakdown
          </h2>
          <MonthlyView weeklyData={data.weekly_hours.map(week => ({
            week: week.week,
            dateRange: week.dateRange || '',
            totalMinutes: week.totalMinutes || week.hours * 60,
            focusMinutes: week.focusMinutes || (week.hours * 60 * 0.8), // Assume 80% focus if not provided
          }))}/>
        </div>
      )}
      
      {/* Debug info */}
      <div className="mt-8 text-xs text-gray-500">
        <p>Python handler ready: {pyHandlerReady ? 'Yes' : 'No'}</p>
        <p>Mode: {mode}</p>
        <p>Date: {date.toISOString()}</p>
      </div>
    </div>
  );
};

export default FocusDashboard;
