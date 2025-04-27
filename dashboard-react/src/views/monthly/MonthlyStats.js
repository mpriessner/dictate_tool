import React from 'react';
import { MODE_COLORS, formatHours } from './MonthlyUtils';

const MonthlyStats = ({ stats }) => {
  if (!stats) return null;

  const calculateDailyAverage = (total) => {
    return stats.daysWithActivity > 0 ? total / stats.daysWithActivity : 0;
  };

  return (
    <div className="grid grid-cols-2 gap-4 mt-4">
      {/* Monthly Focus Stats */}
      <div className="bg-gray-800 rounded-lg p-4">
        <h3 className="text-lg font-semibold mb-2">Monthly Focus</h3>
        <div className="space-y-2">
          <div className="flex justify-between items-center">
            <div className="flex items-center">
              <div className="w-3 h-3 rounded-full mr-2" style={{ backgroundColor: MODE_COLORS.work }}></div>
              <span>Work</span>
            </div>
            <div>
              <div>{formatHours(stats.totalWork)}</div>
              <div className="text-xs text-gray-400">
                avg: {formatHours(calculateDailyAverage(stats.totalWork))}/active day
              </div>
            </div>
          </div>
          <div className="flex justify-between items-center">
            <div className="flex items-center">
              <div className="w-3 h-3 rounded-full mr-2" style={{ backgroundColor: MODE_COLORS.leisure }}></div>
              <span>Leisure</span>
            </div>
            <div>
              <div>{formatHours(stats.totalLeisure)}</div>
              <div className="text-xs text-gray-400">
                avg: {formatHours(calculateDailyAverage(stats.totalLeisure))}/active day
              </div>
            </div>
          </div>
          <div className="border-t border-gray-700 my-2"></div>
          <div className="flex justify-between items-center font-semibold">
            <span>Total Focus</span>
            <div>
              <div>{formatHours(stats.totalWork + stats.totalLeisure)}</div>
              <div className="text-xs text-gray-400">
                avg: {formatHours(calculateDailyAverage(stats.totalWork + stats.totalLeisure))}/active day
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Monthly Activity Stats */}
      <div className="bg-gray-800 rounded-lg p-4">
        <h3 className="text-lg font-semibold mb-2">Monthly Activity</h3>
        <div className="space-y-2">
          <div className="flex justify-between items-center">
            <div className="flex items-center">
              <div className="w-3 h-3 rounded-full mr-2" style={{ backgroundColor: MODE_COLORS.break }}></div>
              <span>Break Time</span>
            </div>
            <div>
              <div>{formatHours(stats.totalBreak)}</div>
              <div className="text-xs text-gray-400">
                avg: {formatHours(calculateDailyAverage(stats.totalBreak))}/active day
              </div>
            </div>
          </div>
          <div className="flex justify-between items-center">
            <div className="flex items-center">
              <div className="w-3 h-3 rounded-full mr-2" style={{ backgroundColor: MODE_COLORS.inactive }}></div>
              <span>Inactive Time</span>
            </div>
            <div>
              <div>{formatHours(stats.totalInactive)}</div>
              <div className="text-xs text-gray-400">
                avg: {formatHours(calculateDailyAverage(stats.totalInactive))}/active day
              </div>
            </div>
          </div>
          <div className="border-t border-gray-700 my-2"></div>
          <div className="flex justify-between items-center">
            <span>Active Days</span>
            <div className="text-right">
              <div>{stats.daysWithActivity} days</div>
              <div className="text-xs text-gray-400">with focus activity</div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default MonthlyStats;
