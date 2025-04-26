import React, { useState, useEffect } from 'react';

// Get Monday of the current week
const getMondayOfWeek = (d) => {
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1); // adjust when day is Sunday
  return new Date(d.setDate(diff));
};

// Timeline Component
const Timeline = ({ intervals }) => {
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
          
          const startPercent = (startHour / 24) * 100;
          const width = ((endHour - startHour) / 24) * 100;
          
          // Determine color based on mode
          let bgColorClass = 'bg-gray-400'; // Default for breaks
          if (interval.focus) {
            if (interval.mode === 'work') {
              bgColorClass = 'bg-blue-500'; // Blue for work focus
            } else if (interval.mode === 'leisure') {
              bgColorClass = 'bg-amber-400'; // Amber/yellow for leisure focus
            } else {
              bgColorClass = 'bg-blue-400'; // Fallback for old data without mode
            }
          }
          
          return (
            <div
              key={index}
              className={`absolute h-8 ${bgColorClass}`}
              style={{
                left: `${startPercent}%`,
                width: `${Math.max(0.5, width)}%`
              }}
              title={`${interval.start.toLocaleTimeString()} - ${interval.end.toLocaleTimeString()} (${interval.mode || (interval.focus ? 'focus' : 'break')})`}
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

// Weekly View Component
const WeeklyView = ({ dailyHours }) => {
  // Make sure dailyHours is an array
  const days = Array.isArray(dailyHours) ? dailyHours : [];
  console.log('WeeklyView received dailyHours:', dailyHours);
  console.log('Processed days array:', days);
  if (days.length === 0) {
    console.warn('No days data to render in WeeklyView');
  } else {
    console.log('Days with non-zero work hours:', days.filter(d => d.focusWorkHours > 0));
    console.log('Days with non-zero leisure hours:', days.filter(d => d.focusLeisureHours > 0));
    console.log('Days with non-zero break hours:', days.filter(d => d.breakHours > 0));
  }
  
  // Calculate average focus hours from days with data
  const nonZeroDays = days.filter(day => day.focusHours > 0);
  const average = nonZeroDays.length > 0 
    ? (nonZeroDays.reduce((sum, day) => sum + day.focusHours, 0) / nonZeroDays.length).toFixed(1)
    : 0;
  
  // Find max for scaling
  // Include all time types for max calculation
  const rawMaxHours = days.length > 0 
    ? Math.max(0.5, ...days.map(day => {
        const workHours = parseFloat(day.focusWorkHours) || 0;
        const leisureHours = parseFloat(day.focusLeisureHours) || 0;
        const breakHours = parseFloat(day.breakHours) || 0;
        console.log(`Day ${day.day}: work=${workHours}, leisure=${leisureHours}, break=${breakHours}`);
        return workHours + leisureHours + breakHours;
      }))
    : 0.5;
  
  // Calculate separate max for focus time only (to make focus bars more visible)
  const focusMaxHours = days.length > 0
    ? Math.max(0.5, ...days.map(day => {
        const workHours = parseFloat(day.focusWorkHours) || 0;
        const leisureHours = parseFloat(day.focusLeisureHours) || 0;
        return workHours + leisureHours;
      }))
    : 0.5;
    
  console.log('Max hours for scaling (total):', rawMaxHours);
  console.log('Max hours for focus only:', focusMaxHours);
  
  // Set the max hours to be 1 hour above the highest bar instead of using a fixed scale
  // This makes the chart adapt to the actual data range
  const maxHours = Math.ceil(rawMaxHours) + 1;
  console.log('Dynamic max hours set to:', maxHours, '(ceiling of', rawMaxHours, '+ 1)');
  
  // Generate dynamic Y-axis labels based on the max hours
  // Use a step size appropriate for the data range
  const yAxisLabels = [];
  // Determine appropriate step size based on the max hours
  let yAxisStep;
  if (maxHours <= 2) {
    yAxisStep = 0.2;
  } else if (maxHours <= 4) {
    yAxisStep = 0.5;
  } else if (maxHours <= 10) {
    yAxisStep = 1;
  } else {
    yAxisStep = 2;
  }
  
  // Generate labels from 0 to maxHours
  for (let i = 1; i <= Math.floor(maxHours / yAxisStep); i++) {
    yAxisLabels.push((i * yAxisStep).toFixed(1).replace('.0', '')); // Format label
  }
  console.log('Generated y-axis labels:', yAxisLabels, 'with step size:', yAxisStep);

  return (
    <div className="relative bg-gray-800 p-4 rounded-lg shadow-lg flex flex-col" style={{ height: '300px' }}> {/* Increased height slightly */}
      {/* Chart Title and Average */}
      <div className="flex justify-between items-center mb-4">
        <h3 className="text-lg font-semibold text-white">Daily Focus</h3>
        <span className="text-sm text-gray-400">Average: {average}h focus/day</span>
      </div>

      {/* Main Chart Area (Y-axis + Bars + Grid) */}
      <div className="flex-1 flex mb-6"> {/* Added margin-bottom for X-axis labels */} 
        {/* Y-axis labels - Align items to space them correctly */}
        <div className="flex flex-col justify-between w-10 text-right pr-2 text-xs text-gray-400">
          {/* Add the maximum value at the top */}
          <div className="h-0 relative">
            <span className="absolute -top-1.5 right-2">{maxHours}h</span>
          </div>
          {/* Add intermediate labels */}
          {yAxisLabels.slice().reverse().map((label, index) => (
            <div key={index} className="h-0 relative">
              {/* Position label slightly above the line it corresponds to */} 
              <span className="absolute -top-1.5 right-2">{label}h</span> 
            </div>
          ))}
          {/* Explicit 0h label at the bottom, aligned with the bottom grid line */}
          <div className="h-0 relative">
            <span className="absolute -top-1.5 right-2">0h</span>
          </div>
        </div>
        
        {/* Chart Bars and Grid Lines Area */}
        <div className="flex-1 relative">
          {/* Horizontal grid lines - Dynamically generated based on labels */} 
          <div className="absolute inset-0 flex flex-col justify-between pointer-events-none">
            {/* Top grid line for max value */}
            <hr className="border-t border-gray-700" />
            {/* Intermediate grid lines */}
            {yAxisLabels.map((_, index) => (
              <hr key={index} className="border-t border-gray-700" />
            ))}
            {/* Bottom border line (for 0h) */}
            <hr className="border-t border-gray-700" /> 
          </div>
          
          {/* Bars - Use padding to align with grid */} 
          <div className="absolute inset-0 flex items-end justify-around">
            {days.map((day, index) => {
              // Ensure we have valid numbers
              const workHours = parseFloat(day.focusWorkHours) || 0;
              const leisureHours = parseFloat(day.focusLeisureHours) || 0;
              const breakHours = parseFloat(day.breakHours) || 0;
              
              // Calculate heights as percentages
              const workHeight = maxHours > 0 ? (workHours / maxHours) * 100 : 0;
              const leisureHeight = maxHours > 0 ? (leisureHours / maxHours) * 100 : 0;
              const breakHeight = maxHours > 0 ? (breakHours / maxHours) * 100 : 0;
              
              console.log(`Rendering bar for ${day.day}: workHeight=${workHeight}%, leisureHeight=${leisureHeight}%, breakHeight=${breakHeight}%`);
              
              return (
                <div key={index} className="flex flex-col items-center" style={{ width: '14%' }}>
                  <div className="relative w-8 h-full flex flex-col-reverse">
                    {/* Work focus time bar (bottom) */}
                    {workHours > 0 && (
                      <div 
                        className="w-full bg-blue-500"
                        style={{ 
                          height: `${workHeight}%`, 
                          minHeight: workHours > 0 ? '4px' : '0' 
                        }}
                        title={`Work: ${Math.floor(workHours)}h ${Math.round((workHours % 1) * 60)}m`}
                      />
                    )}
                    
                    {/* Leisure focus time bar (middle) */}
                    {leisureHours > 0 && (
                      <div 
                        className="w-full bg-amber-400"
                        style={{ 
                          height: `${leisureHeight}%`, 
                          minHeight: leisureHours > 0 ? '4px' : '0' 
                        }}
                        title={`Leisure: ${Math.floor(leisureHours)}h ${Math.round((leisureHours % 1) * 60)}m`}
                      />
                    )}
                    
                    {/* Break time bar (top) */}
                    {breakHours > 0 && (
                      <div 
                        className="w-full bg-cyan-400 rounded-t"
                        style={{ 
                          height: `${breakHeight}%`, 
                          minHeight: breakHours > 0 ? '4px' : '0' 
                        }}
                        title={`Break: ${Math.floor(breakHours)}h ${Math.round((breakHours % 1) * 60)}m`}
                      />
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* X-Axis Labels - Positioned below the chart area */} 
      <div className="flex justify-around pl-10 pr-0 -mt-4 text-xs text-gray-400"> {/* Adjusted margin-top and padding */} 
        {days.map((day, index) => (
          <div key={index} className="flex flex-col items-center text-center" style={{ width: '14%' }}>
            <div>{day.day || ''}</div>
            <div>{day.date || ''}</div>
          </div>
        ))}
      </div>
      
      {/* Legend - Moved below X-Axis Labels */} 
      <div className="flex justify-center space-x-4 mt-3 text-xs text-gray-400">
        <div className="flex items-center">
          <span className="w-3 h-3 bg-blue-500 rounded-sm mr-1"></span>Work
          <span className="w-3 h-3 bg-amber-400 rounded-sm mr-1"></span>Leisure
          <span className="w-3 h-3 bg-cyan-400 rounded-sm mr-1"></span>Break
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

  // Set up colors for category bars
  const categoryColors = [
    "#F2CDCD", "#DDB6F2", "#F5C2E7", "#E8A2AF", "#F28FAD",
    "#ABE9B3", "#FAE3B0", "#F8BD96", "#EE99A0", "#89DCEB"
  ];

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
          const processedIntervals = result.intervals.map(interval => ({
            start: new Date(interval.start),
            end: new Date(interval.end),
            focus: interval.focus,
            mode: interval.mode || (interval.focus ? 'work' : 'break') // Default to 'work' for backward compatibility
          }));
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
        <div className="flex space-x-4">
          <button 
            className={`px-4 py-2 rounded-md ${mode === 'day' ? 'bg-blue-500' : 'bg-gray-700'}`}
            onClick={() => setMode('day')}
          >
            Day
          </button>
          <button 
            className={`px-4 py-2 rounded-md ${mode === 'week' ? 'bg-blue-500' : 'bg-gray-700'}`}
            onClick={() => setMode('week')}
          >
            Week
          </button>
          <button 
            className={`px-4 py-2 rounded-md ${mode === 'month' ? 'bg-blue-500' : 'bg-gray-700'}`}
            onClick={() => setMode('month')}
          >
            Month
          </button>
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
          <Timeline intervals={data.intervals} />
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
                  const totalTime = (data.focusWork || 0) + (data.focusLeisure || 0) + data.break;
                  const workPercent = totalTime > 0 ? (data.focusWork || 0) / totalTime : 0;
                  const leisurePercent = totalTime > 0 ? (data.focusLeisure || 0) / totalTime : 0;
                  const breakPercent = totalTime > 0 ? data.break / totalTime : 0;
                  
                  // Calculate stroke dash values
                  const circumference = 251.2; // 2 * PI * r where r = 40
                  const workDash = workPercent * circumference;
                  const leisureDash = leisurePercent * circumference;
                  const breakDash = breakPercent * circumference;
                  
                  // Calculate offsets
                  const breakOffset = 0;
                  const leisureOffset = -1 * breakDash;
                  const workOffset = -1 * (breakDash + leisureDash);
                  
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
