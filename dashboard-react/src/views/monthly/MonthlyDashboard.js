import React, { useState, useEffect } from 'react';
import MonthlyCalendar from './MonthlyCalendar';
import MonthlyStats from './MonthlyStats';
import MonthlyDistribution from './MonthlyDistribution';
import { getFirstDayOfMonth, getLastDayOfMonth, calculateMonthlyStats } from './MonthlyUtils';

const MonthlyDashboard = ({ date = new Date() }) => {
  const [monthData, setMonthData] = useState([]);
  const [stats, setStats] = useState(null);

  useEffect(() => {
    const fetchMonthlyData = async () => {
      try {
        const firstDay = getFirstDayOfMonth(date);
        const lastDay = getLastDayOfMonth(date);
        
        // Try to get data from Python backend first
        if (window.pyHandler) {
          try {
            const result = await window.pyHandler.get_data(
              'month',
              firstDay.toISOString().split('T')[0]
            );
            
            if (result && Array.isArray(result)) {
              setMonthData(result);
              const monthlyStats = calculateMonthlyStats(result);
              setStats(monthlyStats);
              return;
            }
          } catch (error) {
            console.warn('Failed to get data from Python backend:', error);
          }
        }

        // Fallback: Fetch each day's data from CSV
        const monthlyData = [];
        for (let d = new Date(firstDay); d <= lastDay; d.setDate(d.getDate() + 1)) {
          const formattedDate = d.toISOString().split('T')[0];

          try {
            const response = await fetch(`/focus_logs/${formattedDate}/activity_log.csv`);
            const csvText = await response.text();
            
            // Process CSV data
            const lines = csvText.split('\n').filter(line => line.trim());
            const data = {
              date: formattedDate,
              focusWork: 0,
              focusLeisure: 0,
              break: 0,
              inactive: 0
            };

            // Simple accumulation of active time
            lines.slice(1).forEach(line => {
              const [, active, , , , elapsed] = line.split(',');
              const activeValue = parseInt(active);
              const elapsedHours = parseFloat(elapsed);
              
              if (activeValue === 1) data.focusWork += elapsedHours;
              else if (activeValue === 2) data.focusLeisure += elapsedHours;
              else if (activeValue === 0) data.break += elapsedHours;
            });

            monthlyData.push(data);
          } catch (error) {
            // If data doesn't exist for this day, add empty data
            monthlyData.push({
              date: formattedDate,
              focusWork: 0,
              focusLeisure: 0,
              break: 0,
              inactive: 0
            });
          }
        }

        setMonthData(monthlyData);
        const monthlyStats = calculateMonthlyStats(monthlyData);
        setStats(monthlyStats);
      } catch (error) {
        console.error('Error loading monthly data:', error);
      }
    };

    fetchMonthlyData();
  }, [date]);

  return (
    <div className="p-4">
      <div className="max-w-6xl mx-auto">
        <h2 className="text-2xl font-bold mb-4">
          Monthly Focus - {date.toLocaleDateString(undefined, { month: 'long', year: 'numeric' })}
        </h2>
        
        {/* Monthly Calendar */}
        <div className="mb-6">
          <MonthlyCalendar monthData={monthData} />
        </div>

        {/* Focus Distribution and Monthly Statistics */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-8 mb-4">
          {/* Focus Distribution Pie Chart */}
          <div className="bg-gray-800 p-6 rounded-lg">
            {stats && <MonthlyDistribution stats={stats} />}
          </div>
          
          {/* Monthly Statistics */}
          <div className="bg-gray-800 p-6 rounded-lg">
            <h2 className="text-xl font-semibold mb-4">Monthly Summary</h2>
            {stats && <MonthlyStats stats={stats} />}
          </div>
        </div>
      </div>
    </div>
  );
};

export default MonthlyDashboard;
