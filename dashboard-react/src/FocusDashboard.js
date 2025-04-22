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
          
          return (
            <div
              key={index}
              className={`absolute h-8 ${interval.focus ? 'bg-blue-400' : 'bg-cyan-400'}`}
              style={{
                left: `${startPercent}%`,
                width: `${Math.max(0.5, width)}%`
              }}
              title={`${interval.start.toLocaleTimeString()} - ${interval.end.toLocaleTimeString()}`}
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
  const daysOfWeek = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  
  // Calculate average from days with data
  const nonZeroDays = dailyHours.filter(day => day.hours > 0);
  const average = nonZeroDays.length > 0 
    ? (nonZeroDays.reduce((sum, day) => sum + day.hours, 0) / nonZeroDays.length).toFixed(1)
    : 0;
  
  // Find max for scaling (at least 8 hours to match the image)
  const maxHours = Math.max(8, ...dailyHours.map(day => day.hours));
  
  return (
    <div className="w-full p-4">
      <h2 className="text-xl font-semibold mb-4 ml-4">Daily Breakdown</h2>
      
      <div className="relative h-80">
        <h3 className="text-lg font-semibold text-center mb-6">Work Hours by Day</h3>
        
        {/* Y-axis labels and grid lines */}
        {[0, 2, 4, 6, 8].map(value => (
          <div key={value} className="absolute w-full" style={{ bottom: `${(value / maxHours) * 85}%` }}>
            <div className="absolute -left-6 text-gray-400">{value}</div>
            <div className="w-full h-px bg-gray-700" />
          </div>
        ))}
        
        {/* Average line with value */}
        {average > 0 && (
          <div 
            className="absolute w-full border-t border-dashed border-white opacity-60 z-10"
            style={{ bottom: `${(average / maxHours) * 85}%` }}
          >
            <div className="absolute right-0 -top-5 text-white opacity-80">Average</div>
            <div className="absolute right-0 top-1 text-white opacity-80">{average}</div>
          </div>
        )}
        
        {/* Bars for each day */}
        <div className="absolute bottom-8 left-0 right-0 h-4/5 flex justify-between">
          {dailyHours.map((dayData, index) => (
            <div key={index} className="relative flex flex-col items-center justify-end" style={{ width: `${100/7}%` }}>
              {dayData.hours > 0 && (
                <>
                  <div className="absolute -top-6 text-white text-sm">{dayData.hours}</div>
                  <div 
                    className="w-4/5 bg-blue-400 rounded-t"
                    style={{ height: `${(dayData.hours / maxHours) * 85}%` }}
                  />
                </>
              )}
            </div>
          ))}
        </div>
        
        {/* X-axis labels */}
        <div className="absolute bottom-0 left-0 right-0 flex justify-between">
          {daysOfWeek.map((day, index) => (
            <div key={index} className="text-gray-300 pb-2" style={{ width: `${100/7}%`, textAlign: 'center' }}>
              {day}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

// Monthly View Component
const MonthlyView = ({ weeklyData }) => {
  return (
    <div className="w-full p-2">
      {weeklyData.map((week, index) => (
        <div key={index} className="flex items-center bg-gray-700 rounded p-3 mb-2">
          <div className="flex-shrink-0 w-16 text-center">
            <div className="text-sm font-medium">{week.week}</div>
            <div className="text-xs text-gray-300">{week.dateRange}</div>
          </div>
          <div className="flex-grow mx-4">
            <div className="w-full h-4 bg-gray-800 rounded-full">
              <div 
                className="h-4 bg-blue-400 rounded-full" 
                style={{ width: `${(week.focusMinutes / week.totalMinutes) * 100}%` }}
              />
            </div>
          </div>
          <div className="flex-shrink-0 text-right">
            <div className="text-sm">{Math.floor(week.totalMinutes / 60)}h {week.totalMinutes % 60}m</div>
            <div className="text-xs text-gray-300">{Math.round((week.focusMinutes / week.totalMinutes) * 100)}% focus</div>
          </div>
        </div>
      ))}
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
        
        // Convert interval ISO strings to Date objects
        if (result.intervals && Array.isArray(result.intervals)) {
          const processedIntervals = result.intervals.map(interval => ({
            start: new Date(interval.start),
            end: new Date(interval.end),
            focus: interval.focus
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
                
                {/* Focus segment */}
                {data.focus > 0 && (
                  <circle 
                    cx="50" 
                    cy="50" 
                    r="40" 
                    fill="transparent" 
                    stroke="#89B4FA" 
                    strokeWidth="20"
                    strokeDasharray={`${(data.focus / (data.focus + data.break)) * 251.2} 251.2`}
                    strokeDashoffset="0"
                    transform="rotate(-90 50 50)"
                  />
                )}
                
                {/* Break segment */}
                {data.break > 0 && (
                  <circle 
                    cx="50" 
                    cy="50" 
                    r="40" 
                    fill="transparent" 
                    stroke="#74C7EC" 
                    strokeWidth="20"
                    strokeDasharray={`${(data.break / (data.focus + data.break)) * 251.2} 251.2`}
                    strokeDashoffset={`${-1 * (data.focus / (data.focus + data.break)) * 251.2}`}
                    transform="rotate(-90 50 50)"
                  />
                )}
                
                {/* Center text */}
                <text x="50" y="45" textAnchor="middle" className="text-lg font-semibold" fill="white">
                  {Math.round((data.focus / (data.focus + data.break)) * 100)}%
                </text>
                <text x="50" y="60" textAnchor="middle" className="text-sm" fill="white">
                  Focus
                </text>
              </svg>
            </div>
          </div>
          
          {/* Legend */}
          <div className="flex justify-center mt-4 space-x-8">
            <div className="flex items-center">
              <div className="w-4 h-4 bg-blue-400 mr-2"></div>
              <span>Focus: {Math.floor(data.focus / 60)}h {data.focus % 60}m</span>
            </div>
            <div className="flex items-center">
              <div className="w-4 h-4 bg-cyan-400 mr-2"></div>
              <span>Break: {Math.floor(data.break / 60)}h {data.break % 60}m</span>
            </div>
          </div>
        </div>
        
        {/* App usage breakdown */}
        <div className="bg-gray-800 p-6 rounded-lg">
          <h2 className="text-xl font-semibold mb-4">App Usage</h2>
          <div className="space-y-4">
            {Object.entries(data.app_usage || {})
              .sort(([, a], [, b]) => b - a)
              .slice(0, 10)
              .map(([app, percentage], index) => (
                <div key={app} className="w-full">
                  <div className="flex justify-between mb-1">
                    <span>{app}</span>
                    <span>{percentage.toFixed(1)}%</span>
                  </div>
                  <div className="w-full bg-gray-700 rounded-full h-2.5">
                    <div 
                      className="h-2.5 rounded-full" 
                      style={{ 
                        width: `${percentage}%`,
                        backgroundColor: categoryColors[index % categoryColors.length]
                      }}
                    ></div>
                  </div>
                </div>
              ))}
          </div>
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