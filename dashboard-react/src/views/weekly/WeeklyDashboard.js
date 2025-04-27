import React, { useState, useEffect } from 'react';
import WeeklyBars from './WeeklyBars';
import WeeklyStats from './WeeklyStats';
import WeeklyDistribution from './WeeklyDistribution';
import { getMondayOfWeek, calculateWeeklyStats } from './WeeklyUtils';

const WeeklyDashboard = ({ date = new Date() }) => {
  const [weekData, setWeekData] = useState([]);
  const [stats, setStats] = useState({
    totalWork: 0,
    totalLeisure: 0,
    totalBreak: 0,
    dailyStats: []
  });

  useEffect(() => {
    const fetchWeeklyData = async () => {
      try {
        const monday = getMondayOfWeek(new Date(date));
        
        // Try to get data from Python backend first
        if (window.pyHandler) {
          try {
            const result = await window.pyHandler.get_data('week', monday.toISOString().split('T')[0]);
            if (result && Array.isArray(result)) {
              setWeekData(result);
              const weeklyStats = calculateWeeklyStats(result);
              setStats(weeklyStats);
              return;
            }
          } catch (error) {
            console.warn('Failed to get data from Python backend:', error);
          }
        }

        // Fallback: Fetch each day's data from CSV
        const weeklyData = [];
        for (let i = 0; i < 7; i++) {
          const currentDate = new Date(monday);
          currentDate.setDate(monday.getDate() + i);
          const formattedDate = currentDate.toISOString().split('T')[0];

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

            weeklyData.push(data);
          } catch (error) {
            // If data doesn't exist for this day, add empty data
            weeklyData.push({
              date: formattedDate,
              focusWork: 0,
              focusLeisure: 0,
              break: 0
            });
          }
        }

        setWeekData(weeklyData);
        const weeklyStats = calculateWeeklyStats(weeklyData);
        setStats(weeklyStats);
      } catch (error) {
        console.error('Error loading weekly data:', error);
      }
    };

    fetchWeeklyData();
  }, [date]);

  return (
    <div className="p-4">
      <div className="max-w-6xl mx-auto">
        <h2 className="text-2xl font-bold mb-4">
          Weekly Focus - {getMondayOfWeek(new Date(date)).toLocaleDateString()} to {new Date(getMondayOfWeek(new Date(date)).getTime() + 6 * 24 * 60 * 60 * 1000).toLocaleDateString()}
        </h2>
        
        {/* Weekly Bar Chart */}
        <div className="mb-6">
          <WeeklyBars dailyStats={stats.dailyStats} />
        </div>

        {/* Focus Distribution and Weekly Statistics */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-8 mb-4">
          {/* Focus Distribution Pie Chart */}
          <div className="bg-gray-800 p-6 rounded-lg">
            <WeeklyDistribution stats={stats} />
          </div>
          
          {/* Weekly Statistics */}
          <div className="bg-gray-800 p-6 rounded-lg">
            <h2 className="text-xl font-semibold mb-4">Weekly Summary</h2>
            <WeeklyStats stats={stats} />
          </div>
        </div>
      </div>
    </div>
  );
};

export default WeeklyDashboard;
