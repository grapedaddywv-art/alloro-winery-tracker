/** @type {import('tailwindcss').Config} */
export default {
  content: ["./index.html", "./src/**/*.{js,jsx}"],
  theme: {
    extend: {
      colors: {
        // Matches the actual Grape Daddy Winery Solutions logo — cream background, ink-black
        // linework. Used in place of the app's old emerald palette; amber/gold stays as the one
        // accent color of warmth against this otherwise neutral ink-and-cream base.
        ink: {
          50: "#FAF8F4",
          100: "#F4EFE6",
          200: "#E6DFD1",
          300: "#D1C7B3",
          400: "#A69C87",
          500: "#7D7364",
          600: "#5C5347",
          700: "#403A31",
          800: "#2A251F",
          900: "#1D1915",
          950: "#141210",
        },
      },
    },
  },
  plugins: [],
};
