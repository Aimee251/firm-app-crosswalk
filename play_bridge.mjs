// Node bridge for Google Play — three modes, no paid API needed.
//
//   node play_bridge.mjs search    "<company name>"        [num]
//   node play_bridge.mjs developer "<devId>"               [num]
//   node play_bridge.mjs app       "<packageId>"
//
// Prints JSON to stdout (array for search/developer, object for app).
// Setup (once, in this folder):  npm install google-play-scraper

import * as mod from "google-play-scraper";
const gplay = mod.default ?? mod;  // works for both ESM default and namespace

const cmd = process.argv[2];
const arg = process.argv[3];
const num = parseInt(process.argv[4] || "200", 10);

if (!cmd || !arg) {
  console.error('usage: node play_bridge.mjs <search|developer|app> "<value>" [num]');
  process.exit(2);
}

try {
  let out;
  if (cmd === "search") {
    out = await gplay.search({ term: arg, num: Math.min(num, 30),
                               country: "us", lang: "en" });
  } else if (cmd === "developer") {
    out = await gplay.developer({ devId: arg, num, country: "us", lang: "en" });
  } else if (cmd === "app") {
    out = await gplay.app({ appId: arg, country: "us", lang: "en" });
  } else {
    console.error("unknown command: " + cmd);
    process.exit(2);
  }
  process.stdout.write(JSON.stringify(out));
} catch (e) {
  console.error(String(e && e.message ? e.message : e));
  process.exit(1);
}
