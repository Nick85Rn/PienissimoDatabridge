/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      fontFamily: {
        sans: ['Inter', 'sans-serif'],
        mono: ['"IBM Plex Mono"', 'monospace'],
      },
      colors: {
        paper: {
          DEFAULT: '#F7F4ED',
          dark: '#EFE9DA',
        },
        ink: {
          DEFAULT: '#1C2B24',
          soft: '#3A4A42',
        },
        stamp: {
          DEFAULT: '#2F6B4F',
          light: '#E4EEE7',
          dark: '#1F4A36',
          bright: '#5EB58C',
        },
        rust: {
          DEFAULT: '#B5482A',
          light: '#F5E4DD',
          dark: '#8A3520',
        },
      },
      backgroundImage: {
        'paper-grain': "radial-gradient(circle at 1px 1px, rgba(28,43,36,0.035) 1px, transparent 0)",
      },
    },
  },
  plugins: [],
}