import {
  GlobalRangeSelectorService
} from "./chunk-AURPDNMS.js";
import {
  FUniver
} from "./chunk-66S2PCKH.js";

// ../packages/sheets-formula-ui/src/facade/f-univer.ts
var FUniverSheetsFormulaUIMixin = class extends FUniver {
  showRangeSelectorDialog(opts) {
    const globalRangeSelectorService = this._injector.get(GlobalRangeSelectorService);
    return globalRangeSelectorService.showRangeSelectorDialog(opts);
  }
};
FUniver.extend(FUniverSheetsFormulaUIMixin);
