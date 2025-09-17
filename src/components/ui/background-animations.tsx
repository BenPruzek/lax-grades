"use client";
import { cn } from "@/lib/utils";
import { useEffect, useRef, useState } from "react";

export const BackgroundGradientAnimation = ({
    gradientBackgroundStart = "rgb(255, 255, 255)", // light default
    gradientBackgroundEnd = "rgb(255, 255, 255)",
    firstColor = "128, 0, 32", // Light Pinkish Red
    secondColor = "128, 0, 32", // Classic Vibrant Red
    thirdColor = "139, 0, 0", // Deep Dark Red
    fourthColor = "128, 0, 32", // Bright Fiery Red
    fifthColor = "128, 0, 32", // Rich Burgundy Red
    // pointerColor = "205, 92, 92", // Desaturated Red
    size = "10%",
    blendingValue = "hard-light",
    children,
    className,
    interactive = true,
    containerClassName,
}: {
    gradientBackgroundStart?: string;
    gradientBackgroundEnd?: string;
    firstColor?: string;
    secondColor?: string;
    thirdColor?: string;
    fourthColor?: string;
    fifthColor?: string;
    pointerColor?: string;
    size?: string;
    blendingValue?: string;
    children?: React.ReactNode;
    className?: string;
    interactive?: boolean;
    containerClassName?: string;
}) => {
    const interactiveRef = useRef<HTMLDivElement>(null);

    const [curX, setCurX] = useState(0);
    const [curY, setCurY] = useState(0);
    const [tgX, setTgX] = useState(0);
    const [tgY, setTgY] = useState(0);
    useEffect(() => {
        const mql = window.matchMedia("(prefers-color-scheme: dark)");
        const apply = () => {
            const isDark = mql.matches;
            // In dark mode, make white backgrounds beige
            const start = isDark ? "rgb(246, 246, 239)" : gradientBackgroundStart;
            const end = isDark ? "rgb(246, 246, 239)" : gradientBackgroundEnd;
            document.body.style.setProperty("--gradient-background-start", start);
            document.body.style.setProperty("--gradient-background-end", end);
            document.body.style.setProperty("--first-color", firstColor);
            document.body.style.setProperty("--second-color", secondColor);
            document.body.style.setProperty("--third-color", thirdColor);
            document.body.style.setProperty("--fourth-color", fourthColor);
            document.body.style.setProperty("--fifth-color", fifthColor);
            // document.body.style.setProperty("--pointer-color", pointerColor);
            document.body.style.setProperty("--size", size);
            document.body.style.setProperty("--blending-value", blendingValue);
        };
        apply();
        mql.addEventListener("change", apply);
        return () => mql.removeEventListener("change", apply);
    }, [gradientBackgroundStart, gradientBackgroundEnd, firstColor, secondColor, thirdColor, fourthColor, fifthColor, size, blendingValue]);

    useEffect(() => {
        function move() {
            if (!interactiveRef.current) {
                return;
            }
            setCurX(curX + (tgX - curX) / 20);
            setCurY(curY + (tgY - curY) / 20);
            interactiveRef.current.style.transform = `translate(${Math.round(
                curX
            )}px, ${Math.round(curY)}px)`;
        }

        move();
    }, [tgX, tgY]);

    const handleMouseMove = (event: React.MouseEvent<HTMLDivElement>) => {
        if (interactiveRef.current) {
            const rect = interactiveRef.current.getBoundingClientRect();
            setTgX(event.clientX - rect.left);
            setTgY(event.clientY - rect.top);
        }
    };

    return (
        <div
            className={cn(
                "h-screen w-full relative overflow-hidden top-0 left-0 bg-[linear-gradient(40deg,var(--gradient-background-start),var(--gradient-background-end))]",
                containerClassName
            )}
        >
            <svg className="hidden">
                <defs>
                    <filter id="blurMe">
                        <feGaussianBlur
                            in="SourceGraphic"
                            stdDeviation="10"
                            result="blur"
                        />
                        <feColorMatrix
                            in="blur"
                            mode="matrix"
                            values="1 0 0 0 0  0 1 0 0 0  0 0 1 0 0  0 0 0 18 -8"
                            result="goo"
                        />
                        <feBlend in="SourceGraphic" in2="goo" />
                    </filter>
                </defs>
            </svg>
            <div className={cn("", className)}>{children}</div>
            <div className="gradients-container [filter:url(#blurMe)_blur(40px)] h-full w-full">
                <div
                    className={cn(
                        `absolute [background:radial-gradient(circle_at_center,_var(--first-color)_0,_var(--first-color)_50%)_no-repeat]`,
                        `[mix-blend-mode:var(--blending-value)] w-[var(--size)] h-[var(--size)] top-[calc(50%-var(--size)/2)] left-[calc(50%-var(--size)/2)]`,
                        `[transform-origin:center_center]`,
                        `animate-first`,
                        `opacity-100`
                    )}
                ></div>
                <div
                    className={cn(
                        `absolute [background:radial-gradient(circle_at_center,_rgba(var(--second-color),_0.8)_0,_rgba(var(--second-color),_0)_50%)_no-repeat]`,
                        `[mix-blend-mode:var(--blending-value)] w-[var(--size)] h-[var(--size)] top-[calc(50%-var(--size)/2)] left-[calc(50%-var(--size)/2)]`,
                        `[transform-origin:calc(50%-400px)]`,
                        `animate-second`,
                        `opacity-100`
                    )}
                ></div>
                <div
                    className={cn(
                        `absolute [background:radial-gradient(circle_at_center,_rgba(var(--third-color),_0.8)_0,_rgba(var(--third-color),_0)_50%)_no-repeat]`,
                        `[mix-blend-mode:var(--blending-value)] w-[var(--size)] h-[var(--size)] top-[calc(50%-var(--size)/2)] left-[calc(50%-var(--size)/2)]`,
                        `[transform-origin:calc(50%+400px)]`,
                        `animate-third`,
                        `opacity-100`
                    )}
                ></div>
                <div
                    className={cn(
                        `absolute [background:radial-gradient(circle_at_center,_rgba(var(--fourth-color),_0.8)_0,_rgba(var(--fourth-color),_0)_50%)_no-repeat]`,
                        `[mix-blend-mode:var(--blending-value)] w-[var(--size)] h-[var(--size)] top-[calc(50%-var(--size)/2)] left-[calc(50%-var(--size)/2)]`,
                        `[transform-origin:calc(50%-200px)]`,
                        `animate-fourth`,
                        `opacity-70`
                    )}
                ></div>
                <div
                    className={cn(
                        `absolute [background:radial-gradient(circle_at_center,_rgba(var(--fifth-color),_0.8)_0,_rgba(var(--fifth-color),_0)_50%)_no-repeat]`,
                        `[mix-blend-mode:var(--blending-value)] w-[var(--size)] h-[var(--size)] top-[calc(50%-var(--size)/2)] left-[calc(50%-var(--size)/2)]`,
                        `[transform-origin:calc(50%-800px)_calc(50%+800px)]`,
                        `animate-fifth`,
                        `opacity-100`
                    )}
                ></div>

                {interactive && (
                    <div
                        ref={interactiveRef}
                        onMouseMove={handleMouseMove}
                        className={cn(
                            `absolute [background:radial-gradient(circle_at_center,_rgba(var(--pointer-color),_0.8)_0,_rgba(var(--pointer-color),_0)_50%)_no-repeat]`,
                            `[mix-blend-mode:var(--blending-value)] w-full h-full -top-1/2 -left-1/2`,
                            `opacity-70`
                        )}
                    ></div>
                )}
            </div>
        </div>
    );
};
