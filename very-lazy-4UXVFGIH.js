import {
  UniverActionRecorderPlugin
} from "./chunk-SPACWQFH.js";
import {
  UniverSheetsFindReplacePlugin
} from "./chunk-VHKQKHRR.js";
import {
  UniverSheetsSortUIPlugin
} from "./chunk-TVJJCC4D.js";
import {
  UniverUniscriptPlugin
} from "./chunk-7RB2IIKG.js";
import "./chunk-GMDX6E2J.js";
import "./chunk-A3DRKKMY.js";
import "./chunk-WLDOIN2T.js";
import {
  UniverSheetsCrosshairHighlightPlugin
} from "./chunk-ZRQKYQMN.js";
import {
  UniverSheetsHyperLinkUIPlugin
} from "./chunk-AN63EBF6.js";
import "./chunk-HKH3GJ5O.js";
import {
  UniverDebuggerPlugin
} from "./chunk-6Z6V4ZJE.js";
import {
  UniverWatermarkPlugin
} from "./chunk-IOYWP7BT.js";
import "./chunk-UY6HPM6Z.js";
import "./chunk-AURPDNMS.js";
import "./chunk-SHS2QU7R.js";
import "./chunk-VWBPVFU7.js";
import "./chunk-NDZAFLT3.js";
import "./chunk-JUXP4HVV.js";
import "./chunk-66S2PCKH.js";
import "./chunk-BJIQQXAX.js";
import "./chunk-T3S76CBR.js";
import "./chunk-42RBW3MT.js";
import "./chunk-KC3DSKIX.js";
import "./chunk-33LYH36L.js";
import "./chunk-EQ2B2W73.js";
import "./chunk-DO7PIA5W.js";

// src/sheets-no-worker/very-lazy.ts
var IS_E2E = false;
function getVeryLazyPlugins() {
  const plugins = [
    [UniverActionRecorderPlugin],
    [UniverSheetsHyperLinkUIPlugin],
    [UniverSheetsSortUIPlugin],
    [UniverSheetsCrosshairHighlightPlugin],
    [UniverSheetsFindReplacePlugin],
    [UniverWatermarkPlugin]
  ];
  if (!IS_E2E) {
    plugins.push([UniverDebuggerPlugin]);
    plugins.push([UniverUniscriptPlugin, {
      getWorkerUrl(_, label) {
        if (label === "json") {
          return "/vs/language/json/json.worker.js";
        }
        if (label === "css" || label === "scss" || label === "less") {
          return "/vs/language/css/css.worker.js";
        }
        if (label === "html" || label === "handlebars" || label === "razor") {
          return "/vs/language/html/html.worker.js";
        }
        if (label === "typescript" || label === "javascript") {
          return "/vs/language/typescript/ts.worker.js";
        }
        return "/vs/editor/editor.worker.js";
      }
    }]);
  }
  return plugins;
}
export {
  getVeryLazyPlugins as default
};
