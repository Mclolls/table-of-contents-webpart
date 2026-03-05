// Tell TypeScript how to import CSS/SCSS modules.
// Place this file somewhere in your compiled sources (e.g. src/custom.d.ts).
// Ensure tsconfig.json includes the directory (usually "src") so TS picks this up.

declare module '*.module.scss' {
  const classes: { [key: string]: string };
  export default classes;
}

declare module '*.module.css' {
  const classes: { [key: string]: string };
  export default classes;
}

// Optional: allow plain .scss imports (non-module) if you import global scss
declare module '*.scss';
declare module '*.css';