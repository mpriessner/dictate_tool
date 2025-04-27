import React from 'react';
import { MODE_COLORS } from './DailyUtils';

const DailyStats = ({ stats }) => {
  const formatHours = (hours) => {
    const h = Math.floor(hours);
    const m = Math.round((hours - h) * 60);
    return `${h}h ${m}m`;
  };

  return (
    <div className="mt-4 grid grid-cols-2 gap-4">
      {/* Work and Leisure Stats */}
      <div className="bg-gray-800 rounded-lg p-4">
        <h3 className="text-lg font-semibold mb-2">Focus Time</h3>
        <div className="space-y-2">
          <div className="flex justify-between items-center">
            <div className="flex items-center">
              <div className="w-3 h-3 rounded-full mr-2" style={{ backgroundColor: MODE_COLORS.work }}></div>
              <span>Work</span>
            </div>
            <span>{formatHours(stats.work)}</span>
          </div>
          <div className="flex justify-between items-center">
            <div className="flex items-center">
              <div className="w-3 h-3 rounded-full mr-2" style={{ backgroundColor: MODE_COLORS.leisure }}></div>
              <span>Leisure</span>
            </div>
            <span>{formatHours(stats.leisure)}</span>
          </div>
          <div className="border-t border-gray-700 my-2"></div>
          <div className="flex justify-between items-center font-semibold">
            <span>Total Focus</span>
            <span>{formatHours(stats.total)}</span>
          </div>
        </div>
      </div>

      {/* Break and Inactive Stats */}
      <div className="bg-gray-800 rounded-lg p-4">
        <h3 className="text-lg font-semibold mb-2">Other Time</h3>
        <div className="space-y-2">
          <div className="flex justify-between items-center">
            <div className="flex items-center">
              <div className="w-3 h-3 rounded-full mr-2" style={{ backgroundColor: MODE_COLORS.break }}></div>
              <span>Breaks</span>
            </div>
            <span>{formatHours(stats.break)}</span>
          </div>
          <div className="flex justify-between items-center">
            <div className="flex items-center">
              <div className="w-3 h-3 rounded-full mr-2" style={{ backgroundColor: MODE_COLORS.inactive }}></div>
              <span>Inactive</span>
            </div>
            <span>{formatHours(stats.inactive)}</span>
          </div>
          <div className="border-t border-gray-700 my-2"></div>
          <div className="flex justify-between items-center font-semibold">
            <span>Total Time</span>
            <span>{formatHours(stats.totalWithBreaks)}</span>
          </div>
        </div>
      </div>
    </div>
  );
};

export default DailyStats;
