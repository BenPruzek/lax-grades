import type { Config } from "tailwindcss";

const config: Config = {
  content: [
    "./src/pages/**/*.{js,ts,jsx,tsx,mdx}",
    "./src/components/**/*.{js,ts,jsx,tsx,mdx}",
    "./src/app/**/*.{js,ts,jsx,tsx,mdx}",
    "./src/**/*.{ts,tsx}",
  ],
  darkMode: "media",
  theme: {
    extend: {
      colors: {
        maroon: "#840024",
        beige: "#f6f6ef",
      },
      backgroundImage: {
        "gradient-radial": "radial-gradient(var(--tw-gradient-stops))",
        "gradient-conic":
          "conic-gradient(from 180deg at 50% 50%, var(--tw-gradient-stops))",
      },
      animation: {
        first: "moveVertical 30s ease infinite",
        second: "moveInCircle 20s reverse infinite",
        third: "moveInCircle 40s linear infinite",
        fourth: "moveHorizontal 40s ease infinite",
        fifth: "moveInCircle 20s ease infinite",
      },
      keyframes: {
        moveHorizontal: {
          "0%": { transform: "translateX(-50%) translateY(-10%)" },
          "50%": { transform: "translateX(50%) translateY(10%)" },
          "100%": { transform: "translateX(-50%) translateY(-10%)" },
        },
        moveInCircle: {
          "0%": { transform: "rotate(0deg)" },
          "50%": { transform: "rotate(180deg)" },
          "100%": { transform: "rotate(360deg)" },
        },
        moveVertical: {
          "0%": { transform: "translateY(-50%)" },
          "50%": { transform: "translateY(50%)" },
          "100%": { transform: "translateY(-50%)" },
        },
      },
    },
  },
  plugins: [
    require("tailwindcss-animate"),
    function addVariablesForColors({ addBase, theme }: any) {
      const flattenColorPalette =
        require("tailwindcss/lib/util/flattenColorPalette").default;
      const allColors = flattenColorPalette(theme("colors"));
      const entries: [string, unknown][] = Object.entries(allColors);
      const newVars = Object.fromEntries(
        entries.map(([key, val]) => [
          `--${key}`,
          String(val),
        ])
      );
      addBase({ ":root": newVars as Record<string, string> });
    },
  ],
};
export default config;

