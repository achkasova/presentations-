/** @type {import('tailwindcss').Config} */
const config = {
  content: [
    "./app/**/*.{js,ts,jsx,tsx,mdx}",
    "./components/**/*.{js,ts,jsx,tsx,mdx}",
  ],
  theme: {
    extend: {
      colors: {
        umbrella: {
          ink: "#282A32",
          muted: "#6B6D75",
          bg: "#F7F7F7",
          accent: "#0050FF",
          line: "#CCCCCE",
          soft: "#F2F6FF",
          blueSoft: "#CFDEFF",
        },
      },
      fontFamily: {
        sans: ["Graphik LC", "Arial", "Helvetica", "sans-serif"],
        display: ["ALS Sector", "Graphik LC", "Arial", "sans-serif"],
      },
      boxShadow: {
        panel: "0 16px 48px rgba(40, 42, 50, 0.08)",
        card: "0 10px 28px rgba(40, 42, 50, 0.07)",
      },
    },
  },
  plugins: [],
};

module.exports = config;
