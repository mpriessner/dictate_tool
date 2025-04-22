import React, { useState, useEffect } from 'react';

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
      
      {/* Weekly or Monthly view */}
      {mode !== 'day' && (
        <div className="mt-8 bg-gray-800 p-6 rounded-lg">
          <h2 className="text-xl font-semibold mb-4">
            {mode === 'week' ? 'Daily Breakdown' : 'Weekly Breakdown'}
          </h2>
          
          <div className="h-64 relative">
            {/* Y-axis */}
            <div className="absolute left-0 top-0 bottom-0 w-10 flex flex-col justify-between">
              <span>8h</span>
              <span>6h</span>
              <span>4h</span>
              <span>2h</span>
              <span>0h</span>
            </div>
            
            {/* Bars */}
            <div className="ml-10 h-full flex items-end justify-between">
              {mode === 'week' && data.daily_hours && data.daily_hours.map((day, index) => (
                <div key={index} className="flex flex-col items-center">
                  <div 
                    className="w-12 bg-blue-400 rounded-t-md"
                    style={{ height: `${(day.hours / 8) * 100}%` }}
                  ></div>
                  <div className="mt-2">{day.day}</div>
                </div>
              ))}
              
              {mode === 'month' && data.weekly_hours && data.weekly_hours.map((week, index) => (
                <div key={index} className="flex flex-col items-center">
                  <div 
                    className="w-12 bg-blue-400 rounded-t-md"
                    style={{ height: `${(week.hours / 40) * 100}%` }}
                  ></div>
                  <div className="mt-2">{week.week}</div>
                </div>
              ))}
            </div>
          </div>
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