import React from "react";
import { createRoot } from "react-dom/client";

// The app calls window.storage.get/set, which is provided natively inside
// Claude artifacts. When running standalone (GitHub Pages, local, etc.) we
// back it with localStorage instead, so data still persists across reloads
// in this browser.
if (!window.storage) {
  const PREFIX = "packbuilder:";
  window.storage = {
    async get(key, shared) {
      const raw = window.localStorage.getItem(PREFIX + key);
      if (raw === null) throw new Error("key not found: " + key);
      return { key, value: raw, shared: !!shared };
    },
    async set(key, value, shared) {
      window.localStorage.setItem(PREFIX + key, value);
      return { key, value, shared: !!shared };
    },
    async delete(key, shared) {
      window.localStorage.removeItem(PREFIX + key);
      return { key, deleted: true, shared: !!shared };
    },
    async list(prefix, shared) {
      const p = PREFIX + (prefix || "");
      const keys = Object.keys(window.localStorage)
        .filter(k => k.startsWith(p))
        .map(k => k.slice(PREFIX.length));
      return { keys, prefix: prefix || undefined, shared: !!shared };
    }
  };
}

const rootEl = document.getElementById("root");

try {
  const res = await fetch("js/app.jsx");
  if (!res.ok) throw new Error("Could not load js/app.jsx (" + res.status + ")");
  const raw = await res.text();
  const source = 'import React from "react";\n' + raw;
  const transformed = Babel.transform(source, { presets: [["react", { runtime: "classic" }]] }).code;
  const blob = new Blob([transformed], { type: "text/javascript" });
  const url = URL.createObjectURL(blob);
  const mod = await import(url);
  const App = mod.default;
  createRoot(rootEl).render(React.createElement(App));
} catch (err) {
  console.error(err);
  rootEl.innerHTML = '<div class="tw-boot-msg">FAILED TO LOAD APP<br/>' + String(err.message || err) + '</div>';
}
