// Utility functions for weekly view

// Get Monday of the current week
export const getMondayOfWeek = (d) => {
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1); // adjust when day is Sunday
  return new Date(d.setDate(diff));
};

// Format date for display
export const formatDate = (date) => {
  return date.toLocaleDateString(undefined, { 
    weekday: 'short', 
    month: 'short', 
    day: 'numeric' 
  });
};

// Calculate weekly statistics
export const calculateWeeklyStats = (weekData) => {
  if (!weekData) return {};
  
  return {
    totalWork: weekData.reduce((sum, day) => sum + (day.focusWork || 0), 0),
    totalLeisure: weekData.reduce((sum, day) => sum + (day.focusLeisure || 0), 0),
    totalBreak: weekData.reduce((sum, day) => sum + (day.break || 0), 0),
    totalInactive: weekData.reduce((sum, day) => sum + (day.inactive || 0), 0),
    dailyStats: weekData.map(day => ({
      date: new Date(day.date),
      work: day.focusWork || 0,
      leisure: day.focusLeisure || 0,
      break: day.break || 0,
      inactive: day.inactive || 0,
      total: (day.focusWork || 0) + (day.focusLeisure || 0)
    }))
  };
};

// Colors for different activity types
export const MODE_COLORS = {
  work: '#2196F3',    // Blue
  leisure: '#FFCA28', // Yellow/Amber
  break: '#00BCD4',   // Cyan
  inactive: '#9E9E9E' // Gray
};
