// src/config.js

export const BACKEND_URL =
  process.env.NODE_ENV === "production"
    ? "https://autocad-boq-webapp.onrender.com"   // Render URL
    : "http://localhost:8000";                    // Local FastAPI
