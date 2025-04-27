import React from 'react';
import { MODE_COLORS, formatHours } from './MonthlyUtils';

const MonthlyCalendar = ({ monthData }) => {
  const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  const weeks = [];
  let currentWeek = [];

  // Get the first day of the month
  const firstDay = new Date(monthData[0]?.date);
  const firstDayOfMonth = new Date(firstDay.getFullYear(), firstDay.getMonth(), 1);
  
  // Get the day of week for the first day (0 = Sunday, 1 = Monday, ..., 6 = Saturday)
  let firstDayOfWeek = firstDayOfMonth.getDay();
  firstDayOfWeek = firstDayOfWeek === 0 ? 6 : firstDayOfWeek - 1; // Convert to Monday-based

  // Add empty cells for days before the first day of the month
  for (let i = 0; i < firstDayOfWeek; i++) {
    currentWeek.push(null);
  }

  // Add all days of the month
  monthData.forEach((dayData) => {
    const date = new Date(dayData.date);
    
    // Calculate total focus hours
    const totalFocus = (dayData.focusWork || 0) + (dayData.focusLeisure || 0);
    const totalHours = totalFocus + (dayData.break || 0);
    
    // Calculate color intensity based on focus hours
    const maxHoursForIntensity = 8; // Adjust this value based on your needs
    const intensity = Math.min(totalFocus / maxHoursForIntensity, 1);
    
    currentWeek.push({
      date,
      ...dayData,
      totalFocus,
      totalHours,
      intensity
    });

    if (currentWeek.length === 7) {
      weeks.push(currentWeek);
      currentWeek = [];
    }
  });

  // Add empty cells for remaining days in the last week
  if (currentWeek.length > 0) {
    while (currentWeek.length < 7) {
      currentWeek.push(null);
    }
    weeks.push(currentWeek);
  }

  return (
    <div className="bg-gray-800 rounded-lg p-4">
      {/* Calendar header */}
      <div className="grid grid-cols-7 gap-1 mb-1">
        {days.map(day => (
          <div key={day} className="text-center text-sm text-gray-400">
            {day}
          </div>
        ))}
      </div>

      {/* Calendar grid */}
      <div className="grid gap-1">
        {weeks.map((week, weekIndex) => (
          <div key={weekIndex} className="grid grid-cols-7 gap-1">
            {week.map((day, dayIndex) => (
              <div
                key={dayIndex}
                className={`aspect-square p-2 rounded-lg ${
                  day ? 'bg-gray-700' : 'bg-gray-800'
                }`}
              >
                {day && (
                  <div className="h-full flex flex-col">
                    {/* Date number */}
                    <div className="text-sm text-gray-400">
                      {day.date.getDate()}
                    </div>

                    {/* Activity indicator */}
                    {day.totalFocus > 0 && (
                      <div className="flex-1 flex items-center justify-center">
                        <div
                          className="w-full h-full rounded-lg"
                          style={{
                            background: `linear-gradient(to bottom, 
                              ${MODE_COLORS.work}${Math.round(day.focusWork / day.totalFocus * 255).toString(16).padStart(2, '0')},
                              ${MODE_COLORS.leisure}${Math.round(day.focusLeisure / day.totalFocus * 255).toString(16).padStart(2, '0')}
                            )`,
                            opacity: day.intensity
                          }}
                          title={`Work: ${formatHours(day.focusWork)}
Leisure: ${formatHours(day.focusLeisure)}
Break: ${formatHours(day.break)}`}
                        />
                      </div>
                    )}

                    {/* Total hours */}
                    {day.totalHours > 0 && (
                      <div className="text-xs text-center text-gray-400">
                        {day.totalHours.toFixed(1)}h
                      </div>
                    )}
                  </div>
                )}
              </div>
            ))}
          </div>
        ))}
      </div>
    </div>
  );
};

export default MonthlyCalendar;
