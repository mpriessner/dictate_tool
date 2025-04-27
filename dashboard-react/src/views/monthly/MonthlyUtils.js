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
  if (!monthData || !Array.isArray(monthData)) return null;

  const stats = {
    totalWork: 0,
    totalLeisure: 0,
    totalBreak: 0,
    totalInactive: 0,
    daysWithActivity: 0,
    weeklyStats: []
  };

  // Group days by week
  const weeks = {};
  monthData.forEach(day => {
    const date = new Date(day.date);
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
    week.work += day.focusWork || 0;
    week.leisure += day.focusLeisure || 0;
    week.break += day.break || 0;
    week.inactive += day.inactive || 0;
    
    if ((day.focusWork || 0) + (day.focusLeisure || 0) > 0) {
      week.daysWithActivity++;
      stats.daysWithActivity++;
    }

    stats.totalWork += day.focusWork || 0;
    stats.totalLeisure += day.focusLeisure || 0;
    stats.totalBreak += day.break || 0;
    stats.totalInactive += day.inactive || 0;
  });

  // Convert weeks object to array and sort by date
  stats.weeklyStats = Object.values(weeks).sort((a, b) => a.start.localeCompare(b.start));

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
