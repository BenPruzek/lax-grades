"use client";

import { useEffect, useState, useRef } from "react";

export function AnimatedScore({ value }: { value: string }) {
  // Check if value is a number
  const isNumber = !isNaN(parseFloat(value)) && value !== "N/A";
  const numericValue = isNumber ? parseFloat(value) : 0;
  
  const [displayValue, setDisplayValue] = useState(numericValue);
  const previousValue = useRef(numericValue);

  useEffect(() => {
    // If not a number or hasn't changed, don't animate
    if (!isNumber || previousValue.current === numericValue) {
        if (!isNumber) setDisplayValue(0);
        return;
    }

    const start = previousValue.current;
    const end = numericValue;
    const duration = 1000; // 1 second animation speed
    const startTime = performance.now();

    const animate = (currentTime: number) => {
      const elapsed = currentTime - startTime;
      const progress = Math.min(elapsed / duration, 1);
      
      // "Ease Out" math makes the number slow down as it reaches the end
      const easeOut = 1 - Math.pow(1 - progress, 3);

      const current = start + (end - start) * easeOut;
      setDisplayValue(current);

      if (progress < 1) {
        requestAnimationFrame(animate);
      } else {
        // Animation done
        setDisplayValue(end);
        previousValue.current = end;
      }
    };

    requestAnimationFrame(animate);
  }, [numericValue, isNumber]);

  // Dynamic Color Logic: Changes color as the number counts up/down
  const getColor = (val: number) => {
    if (!isNumber) return "text-gray-900 dark:text-white";
    if (val >= 4.0) return "text-emerald-600 dark:text-emerald-400"; // Green for good
    if (val >= 3.0) return "text-gray-900 dark:text-white"; // Black/White for average
    return "text-red-600 dark:text-red-400"; // Red for intense/bad
  };

  return (
    <span className={`text-3xl font-bold transition-colors duration-300 ${getColor(displayValue)}`}>
      {value === "N/A" ? "N/A" : displayValue.toFixed(1)}
    </span>
  );
}
