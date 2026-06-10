import {
  UniverSheetsNoteUIPlugin,
  UniverSheetsTableUIPlugin
} from "./chunk-CNKSQRZZ.js";
import {
  UniverSheetsConditionalFormattingUIPlugin,
  UniverSheetsDataValidationUIPlugin,
  UniverSheetsFilterUIPlugin
} from "./chunk-TVNXSKS3.js";
import {
  UniverSheetsThreadCommentUIPlugin
} from "./chunk-LF4L4X3J.js";
import {
  UniverDocsMentionUIPlugin
} from "./chunk-7LIU4RCV.js";
import {
  UniverThreadCommentUIPlugin
} from "./chunk-RCRW6VC3.js";
import {
  UniverSheetsNumfmtUIPlugin
} from "./chunk-ONNWGSNN.js";
import "./chunk-DH4VAHA6.js";
import "./chunk-FFH3E73X.js";
import "./chunk-HKH3GJ5O.js";
import "./chunk-PF2GOS5P.js";
import "./chunk-I2H6IZXB.js";
import {
  UniverSheetsDrawingUIPlugin
} from "./chunk-UY6HPM6Z.js";
import {
  UniverSheetsFormulaUIPlugin
} from "./chunk-AURPDNMS.js";
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

// src/sheets-no-worker/lazy.ts
function getLazyPlugins() {
  return [
    [UniverDocsMentionUIPlugin],
    [UniverSheetsNumfmtUIPlugin],
    [UniverThreadCommentUIPlugin],
    [UniverSheetsThreadCommentUIPlugin],
    [UniverSheetsNoteUIPlugin],
    [UniverSheetsTableUIPlugin],
    [UniverSheetsFormulaUIPlugin],
    [UniverSheetsDataValidationUIPlugin],
    [UniverSheetsConditionalFormattingUIPlugin],
    [UniverSheetsFilterUIPlugin, { useRemoteFilterValuesGenerator: false }],
    [UniverSheetsDrawingUIPlugin]
  ];
}
export {
  getLazyPlugins as default
};
