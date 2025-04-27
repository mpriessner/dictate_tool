import React, { useState } from 'react';

const AppUsage = ({ appUsage }) => {
  const [showAll, setShowAll] = useState(false);
  
  // Sort apps by usage percentage (descending)
  const sortedApps = Object.entries(appUsage || {})
    .sort((a, b) => b[1] - a[1])
    .filter(([app, percentage]) => percentage > 0);
  
  // Limit the number of apps shown initially
  const displayApps = showAll ? sortedApps : sortedApps.slice(0, 5);
  
  // Check if we have more apps to show
  const hasMoreApps = sortedApps.length > 5;

  return (
    <div className="space-y-3">
      {displayApps.length === 0 ? (
        <div className="text-center text-gray-400 py-4">No app usage data available</div>
      ) : (
        <>
          {displayApps.map(([app, percentage]) => (
            <div key={app} className="flex items-center space-x-2">
              <div className="w-1/3 text-sm truncate" title={app}>
                {app}
              </div>
              <div className="flex-1 h-4 bg-gray-700 rounded-full overflow-hidden">
                <div 
                  className="h-full bg-blue-500 rounded-full"
                  style={{ width: `${percentage}%` }}
                ></div>
              </div>
              <div className="w-12 text-right text-sm">{percentage.toFixed(1)}%</div>
            </div>
          ))}
          
          {hasMoreApps && (
            <button 
              className="text-sm text-blue-400 hover:text-blue-300 mt-2"
              onClick={() => setShowAll(!showAll)}
            >
              {showAll ? 'Show Less' : `Show ${sortedApps.length - 5} More`}
            </button>
          )}
        </>
      )}
    </div>
  );
};

export default AppUsage;
