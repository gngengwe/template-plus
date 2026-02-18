/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{js,jsx}'],
  theme: {
    extend: {
      colors: {
        cloud: '#f8fafc',
        paper: '#ffffff',
        line: '#dbe3ee',
        ink: '#0f172a',
        inkSoft: '#475569',
        accent: '#0ea5e9',
        accent2: '#2563eb',
      },
      boxShadow: {
        glow: '0 18px 45px rgba(30, 64, 175, 0.12)',
        soft: '0 6px 24px rgba(15, 23, 42, 0.08)',
      },
    },
  },
  plugins: [],
}

