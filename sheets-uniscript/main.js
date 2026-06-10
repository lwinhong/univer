import {
  UniverUniscriptPlugin
} from "../chunk-7RB2IIKG.js";
import "../chunk-GMDX6E2J.js";
import "../chunk-A3DRKKMY.js";
import "../chunk-WLDOIN2T.js";
import {
  UniverSheetsNumfmtUIPlugin
} from "../chunk-ONNWGSNN.js";
import {
  UniverDebuggerPlugin
} from "../chunk-6Z6V4ZJE.js";
import "../chunk-IOYWP7BT.js";
import "../chunk-UY6HPM6Z.js";
import {
  UniverSheetsNumfmtPlugin
} from "../chunk-S7WCECRJ.js";
import {
  UniverSheetsUIPlugin
} from "../chunk-SHS2QU7R.js";
import {
  UNISCRIT_WORKBOOK_DATA_DEMO
} from "../chunk-VWBPVFU7.js";
import {
  UniverDocsPlugin,
  UniverDocsUIPlugin
} from "../chunk-NDZAFLT3.js";
import "../chunk-LI6UXASZ.js";
import {
  UniverUIPlugin
} from "../chunk-JUXP4HVV.js";
import "../chunk-66S2PCKH.js";
import {
  zh_CN_default
} from "../chunk-7GI54JNG.js";
import "../chunk-N6FCNQNY.js";
import "../chunk-75OJL7X3.js";
import {
  UniverSheetsFormulaPlugin
} from "../chunk-T3S76CBR.js";
import {
  UniverFormulaEnginePlugin,
  UniverSheetsPlugin
} from "../chunk-42RBW3MT.js";
import {
  UniverRenderEnginePlugin
} from "../chunk-KC3DSKIX.js";
import {
  Univer
} from "../chunk-33LYH36L.js";
import "../chunk-EQ2B2W73.js";
import "../chunk-DO7PIA5W.js";

// src/sheets-uniscript/main.ts
var IS_E2E = false;
var univer = new Univer({
  locale: "zhCN" /* ZH_CN */,
  locales: {
    ["zhCN" /* ZH_CN */]: zh_CN_default
  },
  logLevel: 4 /* VERBOSE */
});
univer.registerPlugin(UniverRenderEnginePlugin);
univer.registerPlugin(UniverUIPlugin, {
  container: "app",
  ribbonType: "classic"
});
univer.registerPlugin(UniverDocsPlugin);
univer.registerPlugin(UniverDocsUIPlugin);
univer.registerPlugin(UniverSheetsPlugin);
univer.registerPlugin(UniverSheetsUIPlugin);
univer.registerPlugin(UniverSheetsNumfmtPlugin);
univer.registerPlugin(UniverSheetsNumfmtUIPlugin);
univer.registerPlugin(UniverFormulaEnginePlugin);
univer.registerPlugin(UniverSheetsFormulaPlugin);
univer.registerPlugin(UniverUniscriptPlugin, {
  getWorkerUrl(_, label) {
    if (label === "typescript" || label === "javascript") {
      return "/vs/language/typescript/ts.worker.js";
    }
    return "/vs/editor/editor.worker.js";
  }
});
if (IS_E2E) {
  univer.registerPlugin(UniverDebuggerPlugin, {
    fab: false,
    performanceMonitor: {
      enabled: false
    }
  });
}
univer.createUnit(2 /* UNIVER_SHEET */, UNISCRIT_WORKBOOK_DATA_DEMO);
window.univer = univer;
