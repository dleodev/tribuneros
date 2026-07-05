import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
// IMPORTANTE: importar el shim ANTES de App para que window.storage exista
// cuando corre el useEffect de seeding del artefacto.
import "./storage.js";
import App from "./App.jsx";
import Gate from "./Gate.jsx";

createRoot(document.getElementById("root")).render(
  <StrictMode>
    <Gate>
      <App />
    </Gate>
  </StrictMode>
);
