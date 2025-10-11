// Utility functions for monthly view

// Get the first day of the month
export const getFirstDayOfMonth = (date) => {
  return new Date(date.getFullYear(), date.getMonth(), 1);
};

// Get the last day of the month
export const getLastDayOfMonth = (date) => {
  return new Date(date.getFullYear(), date.getMonth() + 1, 0);
};

// Get all days in the month
export const getDaysInMonth = (date) => {
  const firstDay = getFirstDayOfMonth(date);
  const lastDay = getLastDayOfMonth(date);
  const days = [];
  
  for (let d = new Date(firstDay); d <= lastDay; d.setDate(d.getDate() + 1)) {
    days.push(new Date(d));
  }
  
  return days;
};

// Calculate monthly statistics
export const calculateMonthlyStats = (monthData) => {
  if (!monthData || !Array.isArray(monthData)) {
    console.error("Invalid monthData:", monthData);
    // Return a valid object with empty weeklyStats array to prevent crashes
    return {
      totalWork: 0,
      totalLeisure: 0,
      totalBreak: 0,
      totalInactive: 0,
      daysWithActivity: 0,
      weeklyStats: [{
        start: new Date().toISOString().split('T')[0],
        work: 0,
        leisure: 0,
        break: 0,
        inactive: 0,
        daysWithActivity: 0
      }]
    };
  }

  console.log("calculateMonthlyStats received data:", monthData);

  const stats = {
    totalWork: 0,
    totalLeisure: 0,
    totalBreak: 0,
    totalInactive: 0,
    daysWithActivity: 0,
    weeklyStats: []
  };

  // If there's no data, create dummy data for the current month
  if (monthData.length === 0) {
    console.log("Empty monthData, creating dummy weeks for current month");
    const today = new Date();
    const firstDay = getFirstDayOfMonth(today);
    const lastDay = getLastDayOfMonth(today);
    
    // Create weekly buckets for the current month
    const currentDate = new Date(firstDay);
    while (currentDate <= lastDay) {
      const weekStart = getWeekStart(currentDate);
      const weekKey = weekStart.toISOString().split('T')[0];
      
      // Only add the week if it's not already in the stats
      if (!stats.weeklyStats.some(w => w.start === weekKey)) {
        stats.weeklyStats.push({
          start: weekKey,
          work: 0,
          leisure: 0,
          break: 0,
          inactive: 0,
          daysWithActivity: 0
        });
      }
      
      // Move to next week
      currentDate.setDate(currentDate.getDate() + 7);
    }
    
    console.log("Created dummy weeks:", stats.weeklyStats);
    return stats;
  }

  // Group days by week
  const weeks = {};
  monthData.forEach(day => {
    // Skip invalid days
    if (!day || !day.date) {
      console.warn("Skipping invalid day:", day);
      return;
    }

    try {
      const date = new Date(day.date);
      if (isNaN(date.getTime())) {
        console.warn(`Invalid date: ${day.date}`);
        return;
      }
      
      const weekStart = getWeekStart(date).toISOString().split('T')[0];
      
      if (!weeks[weekStart]) {
        weeks[weekStart] = {
          start: weekStart,
          work: 0,
          leisure: 0,
          break: 0,
          inactive: 0,
          daysWithActivity: 0
        };
      }

      const week = weeks[weekStart];
      // Make sure we're using 0 for undefined values
      const workHours = parseFloat(day.focusWork || 0);
      const leisureHours = parseFloat(day.focusLeisure || 0);
      const breakHours = parseFloat(day.break || 0);
      const inactiveHours = parseFloat(day.inactive || 0);
      
      week.work += isNaN(workHours) ? 0 : workHours;
      week.leisure += isNaN(leisureHours) ? 0 : leisureHours;
      week.break += isNaN(breakHours) ? 0 : breakHours;
      week.inactive += isNaN(inactiveHours) ? 0 : inactiveHours;
      
      if (workHours > 0 || leisureHours > 0) {
        week.daysWithActivity++;
        stats.daysWithActivity++;
      }

      stats.totalWork += workHours;
      stats.totalLeisure += leisureHours;
      stats.totalBreak += breakHours;
      stats.totalInactive += inactiveHours;
    } catch (error) {
      console.error("Error processing day data:", error, day);
    }
  });

  // Convert weeks object to array and sort by date
  stats.weeklyStats = Object.values(weeks).sort((a, b) => a.start.localeCompare(b.start));
  
  // Add debugging
  console.log("Calculated weekly stats:", stats.weeklyStats);
  console.log("Total weeks:", stats.weeklyStats.length);

  // Ensure we have at least empty weekly stats to prevent empty rendering
  if (stats.weeklyStats.length === 0) {
    console.log("No weekly stats found, adding dummy week");
    // Add a dummy week with today's date if no data
    const today = new Date();
    const mondayOfWeek = getWeekStart(today);
    stats.weeklyStats.push({
      start: mondayOfWeek.toISOString().split('T')[0],
      work: 0,
      leisure: 0,
      break: 0,
      inactive: 0,
      daysWithActivity: 0
    });
  }

  return stats;
};

// Get the start of the week (Monday) for a given date
export const getWeekStart = (date) => {
  const d = new Date(date);
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1);
  return new Date(d.setDate(diff));
};

// Format hours for display
export const formatHours = (hours) => {
  const h = Math.floor(hours);
  const m = Math.round((hours - h) * 60);
  return `${h}h ${m}m`;
};

// Colors for different activity types
export const MODE_COLORS = {
  work: '#2196F3',    // Blue
  leisure: '#FFCA28', // Yellow/Amber
  break: '#00BCD4',   // Cyan
  inactive: '#9E9E9E' // Gray
};
