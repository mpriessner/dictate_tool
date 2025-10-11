import React, { useState, useEffect } from 'react';
import MonthlyCalendar from './MonthlyCalendar';
import MonthlyStats from './MonthlyStats';
import MonthlyDistribution from './MonthlyDistribution';
import MonthlyBars from './MonthlyBars';
import { getFirstDayOfMonth, getLastDayOfMonth, calculateMonthlyStats } from './MonthlyUtils';

const MonthlyDashboard = ({ date = new Date() }) => {
  const [monthData, setMonthData] = useState([]);
  const [stats, setStats] = useState(null);

  useEffect(() => {
    const fetchMonthlyData = async () => {
      try {
        console.log("Fetching monthly data for date:", date);
        const firstDay = getFirstDayOfMonth(date);
        const lastDay = getLastDayOfMonth(date);
        
        console.log("Date range:", firstDay.toISOString(), "to", lastDay.toISOString());
        
        // Try to get data from Python backend first
        if (window.pyHandler) {
          console.log("pyHandler exists, attempting to get data");
          try {
            const result = await window.pyHandler.get_data(
              'month',
              firstDay.toISOString().split('T')[0]
            );
            
            console.log("Python backend returned:", result);
            
            if (result && Array.isArray(result)) {
              setMonthData(result);
              const monthlyStats = calculateMonthlyStats(result);
              console.log("Monthly stats calculated from Python data:", monthlyStats);
              setStats(monthlyStats);
              return;
            } else {
              console.warn("Python backend returned invalid data format");
            }
          } catch (error) {
            console.warn('Failed to get data from Python backend:', error);
          }
        } else {
          console.log("pyHandler not available, falling back to CSV");
        }

        // Fallback: Fetch each day's data from CSV
        console.log("Starting CSV fallback...");
        const monthlyData = [];
        for (let d = new Date(firstDay); d <= lastDay; d.setDate(d.getDate() + 1)) {
          const formattedDate = d.toISOString().split('T')[0];
          console.log("Processing date:", formattedDate);
          
          try {
            const csvUrl = `/focus_logs/${formattedDate}/activity_log.csv`;
            console.log("Fetching CSV from:", csvUrl);
            const response = await fetch(csvUrl);
            
            if (!response.ok) {
              throw new Error(`HTTP error ${response.status}`);
            }
            
            const csvText = await response.text();
            console.log(`Received ${csvText.length} bytes for ${formattedDate}`);
            
            // Process CSV data
            const lines = csvText.split('\n').filter(line => line.trim());
            const data = {
              date: formattedDate,
              focusWork: 0,
              focusLeisure: 0,
              break: 0,
              inactive: 0
            };

            console.log(`Processing ${lines.length - 1} log entries for ${formattedDate}`);
            
            // Simple accumulation of active time
            lines.slice(1).forEach(line => {
              const parts = line.split(',');
              if (parts.length < 6) {
                console.warn("Invalid CSV line format:", line);
                return;
              }
              
              const [, active, , , , elapsed] = parts;
              const activeValue = parseInt(active);
              const elapsedHours = parseFloat(elapsed);
              
              if (isNaN(activeValue) || isNaN(elapsedHours)) {
                console.warn("Invalid values in CSV line:", line);
                return;
              }
              
              if (activeValue === 1) data.focusWork += elapsedHours;
              else if (activeValue === 2) data.focusLeisure += elapsedHours;
              else if (activeValue === 0) data.break += elapsedHours;
            });

            console.log(`Data for ${formattedDate}:`, data);
            monthlyData.push(data);
          } catch (error) {
            console.log(`No data for ${formattedDate}:`, error.message);
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

        console.log(`Collected data for ${monthlyData.length} days`);
        setMonthData(monthlyData);
        const monthlyStats = calculateMonthlyStats(monthlyData);
        console.log("Monthly stats calculated from CSV data:", monthlyStats);
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
        
        {/* Weekly Bar Chart */}
        <div className="mb-6">
          {stats ? (
            stats.weeklyStats ? (
              <MonthlyBars weeklyStats={stats.weeklyStats} />
            ) : (
              <div className="bg-gray-800 p-4 rounded-lg text-center">
                <span className="text-gray-400">Weekly stats data is missing in stats object: {JSON.stringify(stats)}</span>
              </div>
            )
          ) : (
            <div className="bg-gray-800 p-4 rounded-lg text-center">
              <span className="text-gray-400">Loading stats data...</span>
            </div>
          )}
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
