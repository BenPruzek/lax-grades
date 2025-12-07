"use client";

import { useEffect, useState, useRef } from "react";

export function AnimatedScore({ value }: { value: string }) {
  // Check if value is a number
  const isNumber = !isNaN(parseFloat(value)) && value !== "N/A";
  const numericValue = isNumber ? parseFloat(value) : 0;
  
  const [displayValue, setDisplayValue] = useState(numericValue);
  const previousValue = useRef(numericValue);

  useEffect(() => {
    // FIX: If it's not a number (N/A), reset everything to 0 and update the ref
    if (!isNumber) {
        setDisplayValue(0);
        previousValue.current = 0; // <--- THIS WAS MISSING
        return;
    }

    // If the value hasn't effectively changed, do nothing
    if (previousValue.current === numericValue) {
        return;
    }

    const start = previousValue.current;
    const end = numericValue;
    const duration = 1000;
    const startTime = performance.now();

    const animate = (currentTime: number) => {
      const elapsed = currentTime - startTime;
      const progress = Math.min(elapsed / duration, 1);
      
      const easeOut = 1 - Math.pow(1 - progress, 3);

      const current = start + (end - start) * easeOut;
      setDisplayValue(current);

      if (progress < 1) {
        requestAnimationFrame(animate);
      } else {
        setDisplayValue(end);
        previousValue.current = end;
      }
    };

    requestAnimationFrame(animate);
  }, [numericValue, isNumber]);

  const getColor = (val: number) => {
    if (!isNumber) return "text-gray-900 dark:text-white";
    if (val >= 4.0) return "text-emerald-600 dark:text-emerald-400";
    if (val >= 3.0) return "text-gray-900 dark:text-white";
    return "text-red-600 dark:text-red-400";
  };

  return (
    <span className={`text-3xl font-bold transition-colors duration-300 ${getColor(displayValue)}`}>
      {value === "N/A" ? "N/A" : displayValue.toFixed(1)}
    </span>
  );
}