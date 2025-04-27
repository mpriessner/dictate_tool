import React, { useState, useEffect } from 'react';
import DailyTimeline from './DailyTimeline';
import DailyStats from './DailyStats';
import FocusDistribution from './FocusDistribution';
import AppUsage from './AppUsage';
import { processIntervals, calculateDailyStats } from './DailyUtils';

const DailyDashboard = ({ date = new Date() }) => {
  const [intervals, setIntervals] = useState([]);
  const [stats, setStats] = useState({
    work: 0,
    leisure: 0,
    break: 0,
    inactive: 0,
    total: 0,
    totalWithBreaks: 0
  });
  const [appUsage, setAppUsage] = useState({});

  useEffect(() => {
    const fetchDailyData = async () => {
      try {
        // Format the date as YYYY-MM-DD
        const formattedDate = date.toISOString().split('T')[0];
        
        // Try to get data from Python backend first
        if (window.pyHandler) {
          try {
            const result = await window.pyHandler.get_data('day', formattedDate);
            if (result && Object.keys(result).length > 0) {
              // Process intervals
              const processedIntervals = processIntervals(result);
              setIntervals(processedIntervals);

              // Set statistics directly from Python result
              setStats({
                work: result.focusWork || 0,
                leisure: result.focusLeisure || 0,
                break: result.break || 0,
                inactive: result.inactive || 0,
                total: (result.focusWork || 0) + (result.focusLeisure || 0),
                totalWithBreaks: (result.focusWork || 0) + (result.focusLeisure || 0) + (result.break || 0)
              });
              
              // Set app usage data
              setAppUsage(result.app_usage || {});
              return;
            }
          } catch (error) {
            console.warn('Failed to get data from Python backend:', error);
          }
        }

        // Fallback to CSV if Python backend fails or isn't available
        const response = await fetch(`/focus_logs/${formattedDate}/activity_log.csv`);
        const csvText = await response.text();
        
        // Parse CSV
        const lines = csvText.split('\n');
        const headers = lines[0].split(',');
        const data = lines
          .slice(1)
          .filter(line => line.trim())
          .map(line => {
            const values = line.split(',');
            const row = {};
            headers.forEach((header, index) => {
              if (header === 'timestamp') {
                row[header] = values[index];
              } else if (header === 'active' || header.includes('elapsed')) {
                row[header] = parseFloat(values[index]);
              } else {
                row[header] = values[index];
              }
            });
            return row;
          });

        // Process data into intervals
        const processedIntervals = processIntervals(data);
        setIntervals(processedIntervals);

        // Calculate statistics
        const dailyStats = calculateDailyStats(processedIntervals);
        setStats(dailyStats);
      } catch (error) {
        console.error('Error loading daily activity data:', error);
      }
    };

    fetchDailyData();
  }, [date]);

  return (
    <div className="p-4">
      <div className="max-w-6xl mx-auto">
        <h2 className="text-2xl font-bold mb-4">
          Daily Focus - {date.toLocaleDateString()}
        </h2>
        
        {/* Timeline */}
        <div className="bg-gray-800 rounded-lg p-4 mb-4">
          <h3 className="text-lg font-semibold mb-2">Timeline</h3>
          <DailyTimeline intervals={intervals} />
        </div>

        {/* Focus Distribution and App Usage */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-8 mb-4">
          {/* Focus Distribution Pie Chart */}
          <div className="bg-gray-800 p-6 rounded-lg">
            <FocusDistribution stats={stats} />
          </div>
          
          {/* App Usage */}
          <div className="bg-gray-800 p-6 rounded-lg">
            <h2 className="text-xl font-semibold mb-4">App Usage</h2>
            <AppUsage appUsage={appUsage} />
          </div>
        </div>

        {/* Statistics */}
        <DailyStats stats={stats} />
      </div>
    </div>
  );
};

export default DailyDashboard;
