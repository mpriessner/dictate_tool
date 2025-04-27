import React, { useState } from 'react';
import DailyDashboard from './views/daily/DailyDashboard';
import WeeklyDashboard from './views/weekly/WeeklyDashboard';
import MonthlyDashboard from './views/monthly/MonthlyDashboard';

const MainDashboard = () => {
  const [mode, setMode] = useState('day');
  const [date, setDate] = useState(new Date());
  const [showDropdown, setShowDropdown] = useState(false);

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

      {/* Dashboard Content */}
      <div className="container mx-auto">
        {mode === 'day' && <DailyDashboard date={date} />}
        {mode === 'week' && <WeeklyDashboard date={date} />}
        {mode === 'month' && <MonthlyDashboard date={date} />}
      </div>
    </div>
  );
};

export default MainDashboard;
