import React, { useState, useEffect } from 'react';
import { MODE_COLORS, formatTime } from './DailyUtils';

const SLOT_MINUTES = 10;
const MIN_BREAK_MINUTES = 5;

const processIntervals = (data) => {
  // Get day's start and end time (midnight to midnight)
  const firstTimestamp = new Date(data[0].timestamp);
  const dayStart = new Date(
    firstTimestamp.getFullYear(),
    firstTimestamp.getMonth(),
    firstTimestamp.getDate(),
    0, 0, 0, 0
  );
  const dayEnd = new Date(
    firstTimestamp.getFullYear(),
    firstTimestamp.getMonth(),
    firstTimestamp.getDate(),
    23, 59, 59, 999
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
      0,
      0
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

  // Add inactive period from last activity to midnight if needed
  const lastSlot = mergedSlots[mergedSlots.length - 1];
  if (lastSlot && lastSlot.end < dayEnd) {
    mergedSlots.push({
      start: lastSlot.end,
      end: new Date(dayEnd.getTime() + 1000), // Add 1 second to include 23:59:59.999
      mode: 'inactive'
    });
  }

  return mergedSlots;
};

const DailyTimeline = ({ data }) => {
  const [intervals, setIntervals] = useState([]);

  useEffect(() => {
    if (data && data.length > 0) {
      const processedIntervals = processIntervals(data);
      setIntervals(processedIntervals);
    }
  }, [data]);

  const formatHour = (hour) => {
    return hour.toString().padStart(2, '0');
  };

  // Create hour labels
  const hourLabels = [];
  for (let h = 0; h <= 24; h += 4) {
    hourLabels.push(h);
  }

  return (
    <div className="w-full h-16 py-2">
      <div className="relative w-full h-8 bg-gray-800 rounded">
        {intervals.map((interval, index) => {
          const startHour = interval.start.getHours() + interval.start.getMinutes() / 60 + interval.start.getSeconds() / 3600;
          const endHour = interval.end.getHours() + interval.end.getMinutes() / 60 + interval.end.getSeconds() / 3600;
          
          return (
            <div
              key={index}
              className="absolute h-full"
              style={{
                left: `${(startHour / 24) * 100}%`,
                width: `${((endHour - startHour) / 24) * 100}%`,
                backgroundColor: MODE_COLORS[interval.mode],
                opacity: interval.mode === 'inactive' ? 0.5 : 1
              }}
              title={`${formatTime(interval.start)} - ${formatTime(interval.end)} (${interval.mode})`}
            />
          );
        })}
      </div>
      <div className="relative w-full h-6 mt-1">
        {hourLabels.map((hour) => (
          <div 
            key={hour} 
            className="absolute text-xs text-gray-400"
            style={{ left: `${(hour / 24) * 100}%`, transform: 'translateX(-50%)' }}
          >
            {formatHour(hour)}
          </div>
        ))}
      </div>
    </div>
  );
};

export default DailyTimeline;
