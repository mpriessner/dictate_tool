// Utility functions for daily view

// Process intervals from raw activity data or pre-processed intervals
export const processIntervals = (data) => {
  // Check if data is already in interval format
  if (data.intervals && Array.isArray(data.intervals)) {
    return processPreProcessedIntervals(data.intervals, new Date(data.intervals[0].start));
  }
  return processRawActivityData(data);
};

// Process pre-processed intervals from Python backend
const processPreProcessedIntervals = (intervals, date) => {
  const processedIntervals = [];
  let lastEndTime = new Date(date);
  lastEndTime.setHours(0, 0, 0, 0);

  // Sort intervals by start time
  const sortedIntervals = intervals.sort((a, b) => 
    new Date(a.start) - new Date(b.start)
  );

  // Process each interval
  sortedIntervals.forEach(interval => {
    const startTime = new Date(interval.start);
    const endTime = new Date(interval.end);

    // Add inactive period if gap > 1 hour
    if (startTime - lastEndTime > 3600000) { // 1 hour in milliseconds
      processedIntervals.push({
        start: new Date(lastEndTime),
        end: new Date(startTime),
        mode: 'inactive'
      });
    }

    // Add the actual activity interval
    processedIntervals.push({
      start: startTime,
      end: endTime,
      mode: interval.active === 1 ? 'work' : 
            interval.active === 2 ? 'leisure' : 
            interval.active === 0 ? 'break' : 'inactive'
    });

    lastEndTime = endTime;
  });

  // Add final inactive period if needed
  const dayEnd = new Date(date);
  dayEnd.setHours(23, 59, 59, 999);
  if (dayEnd - lastEndTime > 3600000) {
    processedIntervals.push({
      start: new Date(lastEndTime),
      end: dayEnd,
      mode: 'inactive'
    });
  }

  return processedIntervals;
};

// Process raw activity data
const processRawActivityData = (data) => {
  const SLOT_MINUTES = 10;
  const MIN_BREAK_MINUTES = 5;

  // Get day's start time (midnight)
  const firstTimestamp = new Date(data[0].timestamp);
  const dayStart = new Date(
    firstTimestamp.getFullYear(),
    firstTimestamp.getMonth(),
    firstTimestamp.getDate(),
    0, 0, 0, 0
  );

  // First, organize data by minute
  const timelineMap = new Map();
  data.forEach(row => {
    const timestamp = new Date(row.timestamp);
    const minuteKey = new Date(
      timestamp.getFullYear(),
      timestamp.getMonth(),
      timestamp.getDate(),
      timestamp.getHours(),
      timestamp.getMinutes(),
      0, 0
    );
    timelineMap.set(minuteKey.getTime(), row.active);
  });

  // Sort timeline by time
  const sortedMinutes = Array.from(timelineMap.keys()).sort();
  
  const slots = [];
  let currentSlot = {
    start: null,
    activities: [],
    zeroStreak: 0
  };

  // Process minute by minute
  sortedMinutes.forEach((minuteTime, index) => {
    const minute = new Date(minuteTime);
    const activity = timelineMap.get(minuteTime);
    
    if (!currentSlot.start) {
      currentSlot.start = minute;
    }
    
    currentSlot.activities.push(activity);
    
    // Update zero streak
    if (activity === 0) {
      currentSlot.zeroStreak++;
    } else {
      currentSlot.zeroStreak = 0;
    }

    // Check if we should close the current slot
    const minutesInSlot = currentSlot.activities.length;
    let shouldCloseSlot = false;

    // Close slot if:
    // 1. We've reached SLOT_MINUTES
    // 2. We have MIN_BREAK_MINUTES consecutive zeros
    if (minutesInSlot >= SLOT_MINUTES || currentSlot.zeroStreak >= MIN_BREAK_MINUTES) {
      shouldCloseSlot = true;
    }

    if (shouldCloseSlot || index === sortedMinutes.length - 1) {
      // Determine the dominant activity in this slot
      const activities = currentSlot.activities;
      let mode;

      if (currentSlot.zeroStreak >= MIN_BREAK_MINUTES) {
        mode = 'break';
      } else {
        // Count occurrences of each activity type
        const counts = activities.reduce((acc, act) => {
          acc[act] = (acc[act] || 0) + 1;
          return acc;
        }, {});
        
        // Get the most common activity
        const dominantActivity = Object.entries(counts)
          .sort((a, b) => b[1] - a[1])[0][0];
        
        mode = {
          '1': 'work',
          '2': 'leisure',
          '0': 'break'
        }[dominantActivity] || 'inactive';
      }

      slots.push({
        start: currentSlot.start,
        end: new Date(minute.getTime() + 60000), // Add 1 minute
        mode: mode
      });

      // Start new slot
      currentSlot = {
        start: new Date(minute.getTime() + 60000),
        activities: [],
        zeroStreak: activity === 0 ? currentSlot.zeroStreak : 0
      };
    }
  });

  // Merge adjacent slots of the same type
  const mergedSlots = [];
  let currentMerged = slots[0];

  for (let i = 1; i < slots.length; i++) {
    const nextSlot = slots[i];
    if (currentMerged.mode === nextSlot.mode &&
        currentMerged.end.getTime() === nextSlot.start.getTime()) {
      // Merge slots
      currentMerged.end = nextSlot.end;
    } else {
      // Add current slot and start new one
      mergedSlots.push(currentMerged);
      currentMerged = nextSlot;
    }
  }

  // Add the last merged slot
  if (currentMerged) {
    mergedSlots.push(currentMerged);
  }

  // Add inactive period from midnight to first activity if needed
  const firstSlot = mergedSlots[0];
  if (firstSlot && firstSlot.start > dayStart) {
    mergedSlots.unshift({
      start: dayStart,
      end: firstSlot.start,
      mode: 'inactive'
    });
  }
  
  return mergedSlots;
};

// Calculate daily statistics
export const calculateDailyStats = (intervals) => {
  const stats = {
    work: 0,
    leisure: 0,
    break: 0,
    inactive: 0
  };

  intervals.forEach(interval => {
    // Calculate duration in milliseconds
    const durationMs = interval.end.getTime() - interval.start.getTime();
    
    // Convert milliseconds to hours (3600000 ms = 1 hour)
    const durationHours = durationMs / 3600000;
    
    // Add to the appropriate category
    stats[interval.mode] += durationHours;
  });

  return {
    ...stats,
    total: stats.work + stats.leisure,
    totalWithBreaks: stats.work + stats.leisure + stats.break
  };
};

// Format time for display
export const formatTime = (date) => {
  return date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
};

// Mode colors for visualization
export const MODE_COLORS = {
  work: '#2196F3',    // Blue
  leisure: '#FFCA28', // Yellow/Amber
  break: '#00BCD4',   // Cyan
  inactive: '#9E9E9E' // Gray
};
