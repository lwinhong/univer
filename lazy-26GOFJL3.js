import {
  UniverSheetsConditionalFormattingUIPlugin,
  UniverSheetsDataValidationUIPlugin,
  UniverSheetsFilterUIPlugin
} from "./chunk-TVNXSKS3.js";
import "./chunk-PF2GOS5P.js";
import {
  UniverSheetsDrawingUIPlugin
} from "./chunk-UY6HPM6Z.js";
import "./chunk-AURPDNMS.js";
import "./chunk-S7WCECRJ.js";
import "./chunk-SHS2QU7R.js";
import "./chunk-NDZAFLT3.js";
import "./chunk-JUXP4HVV.js";
import "./chunk-BJIQQXAX.js";
import "./chunk-T3S76CBR.js";
import "./chunk-42RBW3MT.js";
import "./chunk-KC3DSKIX.js";
import "./chunk-33LYH36L.js";
import "./chunk-EQ2B2W73.js";
import "./chunk-DO7PIA5W.js";

// src/sheets-multi-units/lazy.ts
function getLazyPlugins() {
  return [
    [UniverSheetsDataValidationUIPlugin],
    [UniverSheetsConditionalFormattingUIPlugin],
    [UniverSheetsFilterUIPlugin, { useRemoteFilterValuesGenerator: false }],
    [UniverSheetsDrawingUIPlugin]
  ];
}
export {
  getLazyPlugins as default
};
