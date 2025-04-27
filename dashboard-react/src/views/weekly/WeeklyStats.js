import React from 'react';
import { MODE_COLORS } from './WeeklyUtils';

const WeeklyStats = ({ stats }) => {
  // Format hours for display
  const formatHours = (minutes) => {
    const h = Math.floor(minutes / 60);
    const m = minutes % 60;
    return `${h}h ${m}m`;
  };

  // Calculate daily averages
  const activeDays = stats.dailyStats.filter(day => day.total > 0).length;
  const avgWork = activeDays > 0 ? stats.totalWork / activeDays : 0;
  const avgLeisure = activeDays > 0 ? stats.totalLeisure / activeDays : 0;
  const avgTotal = activeDays > 0 ? (stats.totalWork + stats.totalLeisure) / activeDays : 0;

  return (
    <div className="space-y-4">
      {/* Weekly Totals */}
      <div className="space-y-2">
        <div className="flex justify-between items-center">
          <div className="flex items-center">
            <div className="w-3 h-3 rounded-full mr-2" style={{ backgroundColor: MODE_COLORS.work }}></div>
            <span>Work Focus</span>
          </div>
          <div>
            <div>{formatHours(stats.totalWork)}</div>
            <div className="text-xs text-gray-400">
              avg: {formatHours(avgWork)}/active day
            </div>
          </div>
        </div>
        <div className="flex justify-between items-center">
          <div className="flex items-center">
            <div className="w-3 h-3 rounded-full mr-2" style={{ backgroundColor: MODE_COLORS.leisure }}></div>
            <span>Leisure Focus</span>
          </div>
          <div>
            <div>{formatHours(stats.totalLeisure)}</div>
            <div className="text-xs text-gray-400">
              avg: {formatHours(avgLeisure)}/active day
            </div>
          </div>
        </div>
        <div className="border-t border-gray-700 my-2"></div>
        <div className="flex justify-between items-center font-semibold">
          <span>Total Focus</span>
          <div>
            <div>{formatHours(stats.totalWork + stats.totalLeisure)}</div>
            <div className="text-xs text-gray-400">
              avg: {formatHours(avgTotal)}/active day
            </div>
          </div>
        </div>
      </div>

      {/* Break Time */}
      <div className="space-y-2 mt-4">
        <div className="flex justify-between items-center">
          <div className="flex items-center">
            <div className="w-3 h-3 rounded-full mr-2" style={{ backgroundColor: MODE_COLORS.break }}></div>
            <span>Break Time</span>
          </div>
          <div>
            <div>{formatHours(stats.totalBreak)}</div>
            <div className="text-xs text-gray-400">
              avg: {formatHours(activeDays > 0 ? stats.totalBreak / activeDays : 0)}/active day
            </div>
          </div>
        </div>
        <div className="border-t border-gray-700 my-2"></div>
        <div className="flex justify-between items-center">
          <span>Active Days</span>
          <div className="text-right">
            <div>{activeDays} days</div>
            <div className="text-xs text-gray-400">with focus activity</div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default WeeklyStats;
