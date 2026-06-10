import {
  CROSSHAIR_HIGHLIGHT_COLORS,
  DisableCrosshairHighlightOperation,
  EnableCrosshairHighlightOperation,
  FindReplaceModel,
  FindReplaceState,
  IFindReplaceService,
  SetCrosshairHighlightColorOperation,
  SheetsCrosshairHighlightService,
  ToggleCrosshairHighlightOperation,
  createInitFindReplaceState
} from "./chunk-ZRQKYQMN.js";
import {
  SheetsThreadCommentModel
} from "./chunk-FFH3E73X.js";
import {
  SheetsHyperLinkParserService,
  SheetsHyperLinkResolverService
} from "./chunk-AN63EBF6.js";
import {
  SortRangeCommand
} from "./chunk-HKH3GJ5O.js";
import {
  AddCfCommand,
  CFNumberOperator,
  CFTimePeriodOperator,
  CFValueType,
  ClearRangeCfCommand,
  ClearWorksheetCfCommand,
  ConditionalFormattingRuleModel,
  DeleteCfCommand,
  IIconSetType,
  MoveCfCommand,
  SetCfCommand,
  createCfId,
  iconMap
} from "./chunk-PF2GOS5P.js";
import {
  AddCommentCommand,
  DeleteCommentCommand,
  DeleteCommentTreeCommand,
  ResolveCommentCommand,
  ThreadCommentModel,
  UpdateCommentCommand,
  getDT
} from "./chunk-I2H6IZXB.js";
import {
  IBatchSaveImagesService,
  ISheetDrawingService,
  InsertSheetDrawingCommand,
  RemoveSheetDrawingCommand,
  SetDrawingApplyMutation,
  SetDrawingArrangeCommand,
  SetSheetDrawingCommand,
  SheetCanvasFloatDomManagerService,
  SheetDrawingAnchorType,
  SheetDrawingUpdateController,
  transformToAxisAlignPosition,
  transformToDrawingPosition
} from "./chunk-UY6HPM6Z.js";
import {
  transformComponentKey
} from "./chunk-7CXBZBEW.js";
import {
  FRange,
  FWorkbook,
  FWorksheet
} from "./chunk-BKEMWLI4.js";
import {
  AddSheetDataValidationCommand,
  ClearRangeDataValidationCommand,
  DataValidationModel,
  IDrawingManagerService,
  RemoveSheetAllDataValidationCommand,
  RemoveSheetDataValidationCommand,
  SetDrawingSelectedOperation,
  SheetDataValidationModel,
  SheetsDataValidationValidatorService,
  UpdateSheetDataValidationOptionsCommand,
  UpdateSheetDataValidationRangeCommand,
  UpdateSheetDataValidationSettingCommand,
  getImageSize,
  getRuleOptions
} from "./chunk-NDZAFLT3.js";
import {
  ComponentManager
} from "./chunk-JUXP4HVV.js";
import {
  FBase,
  FEnum,
  FEventName,
  FUniver
} from "./chunk-66S2PCKH.js";
import {
  ClearSheetsFilterCriteriaCommand,
  CustomFilterOperator,
  RemoveSheetFilterCommand,
  SetSheetFilterRangeCommand,
  SetSheetsFilterCriteriaCommand,
  SheetsFilterService
} from "./chunk-BJIQQXAX.js";
import {
  SheetSkeletonService,
  convertPositionCellToSheetOverGrid,
  convertPositionSheetOverGridToAbsolute,
  deserializeRangeWithSheet,
  serializeListOptions,
  serializeRangeToRefString
} from "./chunk-42RBW3MT.js";
import {
  IRenderManagerService,
  getCurrentTypeOfRenderer,
  getGroupState,
  transformObjectOutOfGroup
} from "./chunk-KC3DSKIX.js";
import {
  CanceledError,
  ColorKit,
  Disposable,
  DrawingTypeEnum,
  ICommandService,
  IURLImageService,
  IUndoRedoService,
  IUniverInstanceService,
  ImageSourceType,
  Inject,
  Injector,
  Range,
  Rectangle,
  RichTextBuilder,
  RichTextValue,
  Tools,
  UserManagerService,
  filter,
  firstValueFrom,
  generateRandomId,
  toDisposable
} from "./chunk-33LYH36L.js";
import {
  __decorateClass,
  __decorateParam,
  __publicField
} from "./chunk-DO7PIA5W.js";

// ../packages/sheets-data-validation/src/facade/f-data-validation-builder.ts
var FDataValidationBuilder = class _FDataValidationBuilder {
  constructor(rule) {
    __publicField(this, "_rule");
    this._rule = rule != null ? rule : {
      uid: generateRandomId(),
      ranges: void 0,
      type: "custom" /* CUSTOM */
    };
  }
  /**
   * Builds an FDataValidation instance based on the _rule property of the current class
   * @returns {FDataValidation} A new instance of the FDataValidation class
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a new data validation rule that requires a number between 1 and 10 for the range A1:B10
   * const fRange = fWorksheet.getRange('A1:B10');
   * const rule = univerAPI.newDataValidation()
   *   .requireNumberBetween(1, 10)
   *   .setOptions({
   *     allowBlank: true,
   *     showErrorMessage: true,
   *     error: 'Please enter a number between 1 and 10'
   *   })
   *   .build();
   * fRange.setDataValidation(rule);
   * ```
   */
  build() {
    return new FDataValidation(this._rule);
  }
  /**
   * Creates a duplicate of the current DataValidationBuilder object
   * @returns {FDataValidationBuilder} A new instance of the DataValidationBuilder class
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a new data validation rule that requires a number between 1 and 10 for the range A1:B10
   * const fRange = fWorksheet.getRange('A1:B10');
   * const builder = univerAPI.newDataValidation()
   *   .requireNumberBetween(1, 10)
   *   .setOptions({
   *     allowBlank: true,
   *     showErrorMessage: true,
   *     error: 'Please enter a number between 1 and 10'
   *   });
   * fRange.setDataValidation(builder.build());
   *
   * // Copy the builder applied to the new range F1:G10
   * const newRange = fWorksheet.getRange('F1:G10');
   * const copyBuilder = builder.copy();
   * newRange.setDataValidation(copyBuilder.build());
   * ```
   */
  copy() {
    return new _FDataValidationBuilder({
      ...this._rule,
      uid: generateRandomId()
    });
  }
  /**
   * Determines whether invalid data is allowed
   * @returns {boolean} True if invalid data is allowed, False otherwise
   * @example
   * ```typescript
   * const builder = univerAPI.newDataValidation().requireNumberBetween(1, 10);
   * console.log(builder.getAllowInvalid());
   * ```
   */
  getAllowInvalid() {
    return this._rule.errorStyle !== 1 /* STOP */;
  }
  /**
   * Gets the data validation type of the rule
   * @returns {DataValidationType | string} The data validation type
   * @example
   * ```typescript
   * const builder = univerAPI.newDataValidation();
   * console.log(builder.getCriteriaType()); // custom
   *
   * builder.requireNumberBetween(1, 10);
   * console.log(builder.getCriteriaType()); // decimal
   *
   * builder.requireValueInList(['Yes', 'No']);
   * console.log(builder.getCriteriaType()); // list
   * ```
   */
  getCriteriaType() {
    return this._rule.type;
  }
  /**
   * Gets the values used for criteria evaluation
   * @returns {[string | undefined, string | undefined, string | undefined]} An array containing the operator, formula1, and formula2 values
   * @example
   * ```typescript
   * const builder = univerAPI.newDataValidation().requireNumberBetween(1, 10);
   * const [operator, formula1, formula2] = builder.getCriteriaValues();
   * console.log(operator, formula1, formula2); // between 1 10
   *
   * builder.requireValueInList(['Yes', 'No']);
   * console.log(builder.getCriteriaValues()); // undefined Yes,No undefined
   * ```
   */
  getCriteriaValues() {
    return [this._rule.operator, this._rule.formula1, this._rule.formula2];
  }
  /**
   * Gets the help text information, which is used to provide users with guidance and support
   * @returns {string | undefined} Returns the help text information. If there is no error message, it returns an undefined value
   * @example
   * ```typescript
   * const builder = univerAPI.newDataValidation().setOptions({
   *   showErrorMessage: true,
   *   error: 'Please enter a valid value'
   * });
   * console.log(builder.getHelpText()); // 'Please enter a valid value'
   * ```
   */
  getHelpText() {
    return this._rule.error;
  }
  /**
   * Sets the data validation rule to require that the input is a boolean value; this value is rendered as a checkbox.
   * @param {string} [checkedValue] - The value assigned to a checked box.
   * @param {string} [uncheckedValue] - The value assigned to an unchecked box.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Set the data validation for cell A1:A10 to require a checkbox with default 1 and 0 values
   * const fRange = fWorksheet.getRange('A1:A10');
   * const rule = univerAPI.newDataValidation()
   *   .requireCheckbox()
   *   .build();
   * fRange.setDataValidation(rule);
   *
   * // Set the data validation for cell B1:B10 to require a checkbox with 'Yes' and 'No' values
   * const fRange2 = fWorksheet.getRange('B1:B10');
   * const rule2 = univerAPI.newDataValidation()
   *   .requireCheckbox('Yes', 'No')
   *   .build();
   * fRange2.setDataValidation(rule2);
   * ```
   */
  requireCheckbox(checkedValue, uncheckedValue) {
    this._rule.type = "checkbox" /* CHECKBOX */;
    this._rule.formula1 = checkedValue;
    this._rule.formula2 = uncheckedValue;
    return this;
  }
  /**
   * Set the data validation type to DATE and configure the validation rules to be after a specific date.
   * @param {Date} date - The latest unacceptable date.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Set some date values in the range A1:B2
   * const fRange = fWorksheet.getRange('A1:B2');
   * fRange.setValues([
   *   ['2024-01-01', '2024-12-31'],
   *   ['2025-01-01', '2025-12-31']
   * ]);
   *
   * // Create a data validation rule that requires a date after 2025-01-01
   * const rule = univerAPI.newDataValidation()
   *   .requireDateAfter(new Date('2025-01-01'))
   *   .build();
   * fRange.setDataValidation(rule);
   *
   * // Get the validation status of the range
   * const status = await fRange.getValidatorStatus();
   * console.log(status); // [['invalid', 'invalid', 'invalid', 'valid']]
   * ```
   */
  requireDateAfter(date) {
    this._rule.type = "date" /* DATE */;
    this._rule.formula1 = date.toLocaleDateString();
    this._rule.operator = "greaterThan" /* GREATER_THAN */;
    return this;
  }
  /**
   * Set the data validation type to DATE and configure the validation rules to be before a specific date.
   * @param {Date} date - The earliest unacceptable date.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Set some date values in the range A1:B2
   * const fRange = fWorksheet.getRange('A1:B2');
   * fRange.setValues([
   *   ['2024-01-01', '2024-12-31'],
   *   ['2025-01-01', '2025-12-31']
   * ]);
   *
   * // Create a data validation rule that requires a date before 2025-01-01
   * const rule = univerAPI.newDataValidation()
   *   .requireDateBefore(new Date('2025-01-01'))
   *   .build();
   * fRange.setDataValidation(rule);
   *
   * // Get the validation status of the range
   * const status = await fRange.getValidatorStatus();
   * console.log(status); // [['valid', 'valid', 'invalid', 'invalid']]
   * ```
   */
  requireDateBefore(date) {
    this._rule.type = "date" /* DATE */;
    this._rule.formula1 = date.toLocaleDateString();
    this._rule.formula2 = void 0;
    this._rule.operator = "lessThan" /* LESS_THAN */;
    return this;
  }
  /**
   * Set the data validation type to DATE and configure the validation rules to be within a specific date range.
   * @param {Date} start - The earliest acceptable date.
   * @param {Date} end - The latest acceptable date.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Set some date values in the range A1:B2
   * const fRange = fWorksheet.getRange('A1:B2');
   * fRange.setValues([
   *   ['2024-01-01', '2024-12-31'],
   *   ['2025-01-01', '2025-12-31']
   * ]);
   *
   * // Create a data validation rule that requires a date between 2024-06-01 and 2025-06-01
   * const rule = univerAPI.newDataValidation()
   *   .requireDateBetween(new Date('2024-06-01'), new Date('2025-06-01'))
   *   .build();
   * fRange.setDataValidation(rule);
   *
   * // Get the validation status of the range
   * const status = await fRange.getValidatorStatus();
   * console.log(status); // [['invalid', 'valid', 'valid', 'invalid']]
   * ```
   */
  requireDateBetween(start, end) {
    this._rule.type = "date" /* DATE */;
    this._rule.formula1 = start.toLocaleDateString();
    this._rule.formula2 = end.toLocaleDateString();
    this._rule.operator = "between" /* BETWEEN */;
    return this;
  }
  /**
   * Set the data validation type to DATE and configure the validation rules to be equal to a specific date.
   * @param {Date} date - The sole acceptable date.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Set some date values in the range A1:B2
   * const fRange = fWorksheet.getRange('A1:B2');
   * fRange.setValues([
   *   ['2024-01-01', '2024-12-31'],
   *   ['2025-01-01', '2025-12-31']
   * ]);
   *
   * // Create a data validation rule that requires a date equal to 2025-01-01
   * const rule = univerAPI.newDataValidation()
   *   .requireDateEqualTo(new Date('2025-01-01'))
   *   .build();
   * fRange.setDataValidation(rule);
   *
   * // Get the validation status of the cell A2
   * const status = await fWorksheet.getRange('A2').getValidatorStatus();
   * console.log(status?.[0]?.[0]); // 'valid'
   *
   * // Get the validation status of the cell B2
   * const status2 = await fWorksheet.getRange('B2').getValidatorStatus();
   * console.log(status2?.[0]?.[0]); // 'invalid'
   * ```
   */
  requireDateEqualTo(date) {
    this._rule.type = "date" /* DATE */;
    this._rule.formula1 = date.toLocaleDateString();
    this._rule.formula2 = void 0;
    this._rule.operator = "equal" /* EQUAL */;
    return this;
  }
  /**
   * Set the data validation type to DATE and configure the validation rules to be not within a specific date range.
   * @param {Date} start - The earliest unacceptable date.
   * @param {Date} end - The latest unacceptable date.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Set some date values in the range A1:B2
   * const fRange = fWorksheet.getRange('A1:B2');
   * fRange.setValues([
   *   ['2024-01-01', '2024-12-31'],
   *   ['2025-01-01', '2025-12-31']
   * ]);
   *
   * // Create a data validation rule that requires a date not between 2024-06-01 and 2025-06-01
   * const rule = univerAPI.newDataValidation()
   *   .requireDateNotBetween(new Date('2024-06-01'), new Date('2025-06-01'))
   *   .build();
   * fRange.setDataValidation(rule);
   *
   * // Get the validation status of the range
   * const status = await fRange.getValidatorStatus();
   * console.log(status); // [['valid', 'invalid', 'invalid', 'valid']]
   * ```
   */
  requireDateNotBetween(start, end) {
    this._rule.type = "date" /* DATE */;
    this._rule.formula1 = start.toLocaleDateString();
    this._rule.formula2 = end.toLocaleDateString();
    this._rule.operator = "notBetween" /* NOT_BETWEEN */;
    return this;
  }
  /**
   * Set the data validation type to DATE and configure the validation rules to be on or after a specific date.
   * @param {Date} date - The earliest acceptable date.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Set some date values in the range A1:B2
   * const fRange = fWorksheet.getRange('A1:B2');
   * fRange.setValues([
   *   ['2024-01-01', '2024-12-31'],
   *   ['2025-01-01', '2025-12-31']
   * ]);
   *
   * // Create a data validation rule that requires a date on or after 2025-01-01
   * const rule = univerAPI.newDataValidation()
   *   .requireDateOnOrAfter(new Date('2025-01-01'))
   *   .build();
   * fRange.setDataValidation(rule);
   *
   * // Get the validation status of the range
   * const status = await fRange.getValidatorStatus();
   * console.log(status); // [['invalid', 'invalid', 'valid', 'valid']]
   * ```
   */
  requireDateOnOrAfter(date) {
    this._rule.type = "date" /* DATE */;
    this._rule.formula1 = date.toLocaleDateString();
    this._rule.formula2 = void 0;
    this._rule.operator = "greaterThanOrEqual" /* GREATER_THAN_OR_EQUAL */;
    return this;
  }
  /**
   * Set the data validation type to DATE and configure the validation rules to be on or before a specific date.
   * @param {Date} date - The latest acceptable date.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Set some date values in the range A1:B2
   * const fRange = fWorksheet.getRange('A1:B2');
   * fRange.setValues([
   *   ['2024-01-01', '2024-12-31'],
   *   ['2025-01-01', '2025-12-31']
   * ]);
   *
   * // Create a data validation rule that requires a date on or before 2025-01-01
   * const rule = univerAPI.newDataValidation()
   *   .requireDateOnOrBefore(new Date('2025-01-01'))
   *   .build();
   * fRange.setDataValidation(rule);
   *
   * // Get the validation status of the range
   * const status = await fRange.getValidatorStatus();
   * console.log(status); // [['valid', 'valid', 'valid', 'invalid']]
   * ```
   */
  requireDateOnOrBefore(date) {
    this._rule.type = "date" /* DATE */;
    this._rule.formula1 = date.toLocaleDateString();
    this._rule.formula2 = void 0;
    this._rule.operator = "lessThanOrEqual" /* LESS_THAN_OR_EQUAL */;
    return this;
  }
  /**
   * Sets the data validation rule to require that the given formula evaluates to `true`.
   * @param {string} formula - The formula string that needs to be satisfied, formula result should be TRUE or FALSE, and references range will relative offset.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Set some values in the range A1:B2 and C1:D2
   * const cell = fWorksheet.getRange('A1:B2');
   * cell.setValues([
   *   [4, 3],
   *   [2, 1]
   * ]);
   * const fRange = fWorksheet.getRange('C1:D2');
   * fRange.setValues([
   *   [1, 2],
   *   [3, 4]
   * ]);
   *
   * // Create a data validation rule that requires the formula '=A1>2' to be satisfied
   * const rule = univerAPI.newDataValidation()
   *   .requireFormulaSatisfied('=A1>2')
   *   .setOptions({
   *     showErrorMessage: true,
   *     error: 'Please enter a value equal to A1'
   *   })
   *   .build();
   * fRange.setDataValidation(rule);
   *
   * // Get the validation status of the range
   * const status = await fRange.getValidatorStatus();
   * console.log(status); // [['valid', 'valid', 'invalid', 'invalid']]
   * ```
   */
  requireFormulaSatisfied(formula) {
    this._rule.type = "custom" /* CUSTOM */;
    this._rule.formula1 = formula;
    this._rule.formula2 = void 0;
    return this;
  }
  /**
   * Sets the data validation rule to require a number that falls between, or is either of, two specified numbers.
   * @param {number} start - The lowest acceptable value.
   * @param {number} end - The highest acceptable value.
   * @param {boolean} [isInteger] - Indicates whether the required number is an integer.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a new data validation rule that requires a number between 1 and 10 for the range A1:B10
   * const fRange = fWorksheet.getRange('A1:B10');
   * const rule = univerAPI.newDataValidation()
   *   .requireNumberBetween(1, 10)
   *   .setOptions({
   *     allowBlank: false,
   *     showErrorMessage: true,
   *     error: 'Please enter a number between 1 and 10'
   *   })
   *   .build();
   * fRange.setDataValidation(rule);
   * ```
   */
  requireNumberBetween(start, end, isInteger) {
    this._rule.formula1 = `${start}`;
    this._rule.formula2 = `${end}`;
    this._rule.operator = "between" /* BETWEEN */;
    this._rule.type = isInteger ? "whole" /* WHOLE */ : "decimal" /* DECIMAL */;
    return this;
  }
  /**
   * Sets the data validation rule to require a number equal to the given value.
   * @param {number} num - The sole acceptable value.
   * @param {boolean} [isInteger] - Indicates whether the required number is an integer.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a new data validation rule that requires a number equal to 10 for the range A1:B10
   * const fRange = fWorksheet.getRange('A1:B10');
   * const rule = univerAPI.newDataValidation()
   *   .requireNumberEqualTo(10)
   *   .setOptions({
   *     allowBlank: false,
   *     showErrorMessage: true,
   *     error: 'Please enter a number equal to 10'
   *   })
   *   .build();
   * fRange.setDataValidation(rule);
   * ```
   */
  requireNumberEqualTo(num, isInteger) {
    this._rule.formula1 = `${num}`;
    this._rule.formula2 = void 0;
    this._rule.operator = "equal" /* EQUAL */;
    this._rule.type = isInteger ? "whole" /* WHOLE */ : "decimal" /* DECIMAL */;
    return this;
  }
  /**
   * Sets the data validation rule to require a number greater than the given value.
   * @param {number} num - The highest unacceptable value.
   * @param {boolean} [isInteger] - Indicates whether the required number is an integer.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a new data validation rule that requires a number greater than 10 for the range A1:B10
   * const fRange = fWorksheet.getRange('A1:B10');
   * const rule = univerAPI.newDataValidation()
   *   .requireNumberGreaterThan(10)
   *   .setOptions({
   *     allowBlank: false,
   *     showErrorMessage: true,
   *     error: 'Please enter a number greater than 10'
   *   })
   *   .build();
   * fRange.setDataValidation(rule);
   * ```
   */
  requireNumberGreaterThan(num, isInteger) {
    this._rule.formula1 = `${num}`;
    this._rule.formula2 = void 0;
    this._rule.operator = "greaterThan" /* GREATER_THAN */;
    this._rule.type = isInteger ? "whole" /* WHOLE */ : "decimal" /* DECIMAL */;
    return this;
  }
  /**
   * Sets the data validation rule to require a number greater than or equal to the given value.
   * @param {number} num - The lowest acceptable value.
   * @param {boolean} [isInteger] - Indicates whether the required number is an integer.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a new data validation rule that requires a number greater than 10 or equal to 10 for the range A1:B10
   * const fRange = fWorksheet.getRange('A1:B10');
   * const rule = univerAPI.newDataValidation()
   *   .requireNumberGreaterThanOrEqualTo(10)
   *   .setOptions({
   *     allowBlank: false,
   *     showErrorMessage: true,
   *     error: 'Please enter a number greater than 10 or equal to 10'
   *   })
   *   .build();
   * fRange.setDataValidation(rule);
   * ```
   */
  requireNumberGreaterThanOrEqualTo(num, isInteger) {
    this._rule.formula1 = `${num}`;
    this._rule.formula2 = void 0;
    this._rule.operator = "greaterThanOrEqual" /* GREATER_THAN_OR_EQUAL */;
    this._rule.type = isInteger ? "whole" /* WHOLE */ : "decimal" /* DECIMAL */;
    return this;
  }
  /**
   * Sets the data validation rule to require a number less than the given value.
   * @param {number} num - The lowest unacceptable value.
   * @param {boolean} [isInteger] - Indicates whether the required number is an integer.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a new data validation rule that requires a number less than 10 for the range A1:B10
   * const fRange = fWorksheet.getRange('A1:B10');
   * const rule = univerAPI.newDataValidation()
   *   .requireNumberLessThan(10)
   *   .setOptions({
   *     allowBlank: false,
   *     showErrorMessage: true,
   *     error: 'Please enter a number less than 10'
   *   })
   *   .build();
   * fRange.setDataValidation(rule);
   * ```
   */
  requireNumberLessThan(num, isInteger) {
    this._rule.formula1 = `${num}`;
    this._rule.formula2 = void 0;
    this._rule.operator = "lessThan" /* LESS_THAN */;
    this._rule.type = isInteger ? "whole" /* WHOLE */ : "decimal" /* DECIMAL */;
    return this;
  }
  /**
   * Sets the data validation rule to require a number less than or equal to the given value.
   * @param {number} num - The highest acceptable value.
   * @param {boolean} [isInteger] - Indicates whether the required number is an integer.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a new data validation rule that requires a number less than 10 or equal to 10 for the range A1:B10
   * const fRange = fWorksheet.getRange('A1:B10');
   * const rule = univerAPI.newDataValidation()
   *   .requireNumberLessThanOrEqualTo(10)
   *   .setOptions({
   *     allowBlank: false,
   *     showErrorMessage: true,
   *     error: 'Please enter a number less than 10 or equal to 10'
   *   })
   *   .build();
   * fRange.setDataValidation(rule);
   * ```
   */
  requireNumberLessThanOrEqualTo(num, isInteger) {
    this._rule.formula1 = `${num}`;
    this._rule.formula2 = void 0;
    this._rule.operator = "lessThanOrEqual" /* LESS_THAN_OR_EQUAL */;
    this._rule.type = isInteger ? "whole" /* WHOLE */ : "decimal" /* DECIMAL */;
    return this;
  }
  /**
   * Sets the data validation rule to require a number that does not fall between, and is neither of, two specified numbers.
   * @param {number} start - The lowest unacceptable value.
   * @param {number} end - The highest unacceptable value.
   * @param {boolean} [isInteger] - Optional parameter, indicating whether the number to be verified is an integer.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a new data validation rule that requires a number not between 1 and 10 for the range A1:B10
   * const fRange = fWorksheet.getRange('A1:B10');
   * const rule = univerAPI.newDataValidation()
   *   .requireNumberNotBetween(1, 10)
   *   .setOptions({
   *     allowBlank: false,
   *     showErrorMessage: true,
   *     error: 'Please enter a number not between 1 and 10'
   *   })
   *   .build();
   * fRange.setDataValidation(rule);
   * ```
   */
  requireNumberNotBetween(start, end, isInteger) {
    this._rule.formula1 = `${start}`;
    this._rule.formula2 = `${end}`;
    this._rule.operator = "notBetween" /* NOT_BETWEEN */;
    this._rule.type = isInteger ? "whole" /* WHOLE */ : "decimal" /* DECIMAL */;
    return this;
  }
  /**
   * Sets the data validation rule to require a number not equal to the given value.
   * @param {number} num - The sole unacceptable value.
   * @param {boolean} [isInteger] - Indicates whether the required number is an integer.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a new data validation rule that requires a number not equal to 10 for the range A1:B10
   * const fRange = fWorksheet.getRange('A1:B10');
   * const rule = univerAPI.newDataValidation()
   *   .requireNumberNotEqualTo(10)
   *   .setOptions({
   *     allowBlank: false,
   *     showErrorMessage: true,
   *     error: 'Please enter a number not equal to 10'
   *   })
   *   .build();
   * fRange.setDataValidation(rule);
   * ```
   */
  requireNumberNotEqualTo(num, isInteger) {
    this._rule.formula1 = `${num}`;
    this._rule.formula2 = void 0;
    this._rule.operator = "notEqual" /* NOT_EQUAL */;
    this._rule.type = isInteger ? "whole" /* WHOLE */ : "decimal" /* DECIMAL */;
    return this;
  }
  /**
   * Sets a data validation rule that requires the user to enter a value from a list of specific values.
   * The list can be displayed in a dropdown, and the user can choose multiple values according to the settings.
   * @param {string[]} values - An array of acceptable values.
   * @param {boolean} [multiple] - Optional parameter indicating whether the user can select multiple values.
   * @param {boolean} [showDropdown] - Optional parameter indicating whether to display the list in a dropdown.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a new data validation rule that requires the user to enter a value from the list ['Yes', 'No'] for the range A1:B10
   * const fRange = fWorksheet.getRange('A1:B10');
   * const rule = univerAPI.newDataValidation()
   *   .requireValueInList(['Yes', 'No'])
   *   .setOptions({
   *     allowBlank: true,
   *     showErrorMessage: true,
   *     error: 'Please enter a value from the list'
   *   })
   *   .build();
   * fRange.setDataValidation(rule);
   * ```
   */
  requireValueInList(values, multiple, showDropdown) {
    this._rule.type = multiple ? "listMultiple" /* LIST_MULTIPLE */ : "list" /* LIST */;
    this._rule.formula1 = serializeListOptions(values);
    this._rule.formula2 = void 0;
    this._rule.showDropDown = showDropdown != null ? showDropdown : true;
    return this;
  }
  /**
   * Sets a data validation rule that requires the user to enter a value within a specific range.
   * The range is defined by an FRange object, which contains the unit ID, sheet name, and cell range.
   * @param {FRange} range - An FRange object representing the range of values that the user can enter.
   * @param {boolean} [multiple] - Optional parameter indicating whether the user can select multiple values.
   * @param {boolean} [showDropdown] - Optional parameter indicating whether to display the list in a dropdown.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Set the values in the range B1:B2
   * const fRange = fWorksheet.getRange('B1:B2');
   * fRange.setValues([
   *   ['Yes'],
   *   ['No']
   * ]);
   *
   * // Create a new data validation rule that requires the user to enter a value from the range B1:B2 for the range A1:A10
   * const rule = univerAPI.newDataValidation()
   *   .requireValueInRange(fRange)
   *   .setOptions({
   *     allowBlank: false,
   *     showErrorMessage: true,
   *     error: 'Please enter a value from the list'
   *   })
   *   .build();
   * const cell = fWorksheet.getRange('A1');
   * cell.setDataValidation(rule);
   * ```
   */
  requireValueInRange(range, multiple, showDropdown) {
    this._rule.type = multiple ? "listMultiple" /* LIST_MULTIPLE */ : "list" /* LIST */;
    this._rule.formula1 = `=${serializeRangeToRefString({
      unitId: range.getUnitId(),
      sheetName: range.getSheetName(),
      range: range.getRange()
    })}`;
    this._rule.formula2 = void 0;
    this._rule.showDropDown = showDropdown != null ? showDropdown : true;
    return this;
  }
  /**
   * Sets whether to allow invalid data and configures the error style.
   * If invalid data is not allowed, the error style will be set to STOP, indicating that data entry must stop upon encountering an error.
   * If invalid data is allowed, the error style will be set to WARNING, indicating that a warning will be displayed when invalid data is entered, but data entry can continue.
   * @param {boolean} allowInvalidData - Whether to allow invalid data.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Set the data validation for cell A1:B2 to allow invalid data, so A1:B2 will display a warning when invalid data is entered
   * const fRange = fWorksheet.getRange('A1:B2');
   * const rule = univerAPI.newDataValidation()
   *   .requireValueInList(['Yes', 'No'])
   *   .setAllowInvalid(true)
   *   .build();
   * fRange.setDataValidation(rule);
   *
   * // Set the data validation for cell C1:D2 to not allow invalid data, so C1:D2 will stop data entry when invalid data is entered
   * const fRange2 = fWorksheet.getRange('C1:D2');
   * const rule2 = univerAPI.newDataValidation()
   *   .requireValueInList(['Yes', 'No'])
   *   .setAllowInvalid(false)
   *   .build();
   * fRange2.setDataValidation(rule2);
   * ```
   */
  setAllowInvalid(allowInvalidData) {
    this._rule.errorStyle = !allowInvalidData ? 1 /* STOP */ : 2 /* WARNING */;
    return this;
  }
  /**
   * Sets whether to allow blank values.
   * @param {boolean} allowBlank - Whether to allow blank values.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * // Assume current sheet is empty data
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Set the data validation for cell A1:B2 to allow blank values
   * const fRange = fWorksheet.getRange('A1:B2');
   * const rule = univerAPI.newDataValidation()
   *   .requireValueInList(['Yes', 'No'])
   *   .setAllowBlank(true)
   *   .build();
   * fRange.setDataValidation(rule);
   *
   * // Set the data validation for cell C1:D2 to not allow blank values
   * const fRange2 = fWorksheet.getRange('C1:D2');
   * const rule2 = univerAPI.newDataValidation()
   *   .requireValueInList(['Yes', 'No'])
   *   .setAllowBlank(false)
   *   .build();
   * fRange2.setDataValidation(rule2);
   * ```
   */
  setAllowBlank(allowBlank) {
    this._rule.allowBlank = allowBlank;
    return this;
  }
  /**
   * Sets the options for the data validation rule.
   * @param {Partial<IDataValidationRuleOptions>} options - The options to set for the data validation rule.
   * @returns {FDataValidationBuilder} The current instance for method chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a new data validation rule that requires the user to enter a value from the list ['Yes', 'No'] for the range A1:B10
   * const fRange = fWorksheet.getRange('A1:B10');
   * const rule = univerAPI.newDataValidation()
   *   .requireValueInList(['Yes', 'No'])
   *   .setOptions({
   *     allowBlank: true,
   *     showErrorMessage: true,
   *     error: 'Please enter a value from the list'
   *   })
   *   .build();
   * fRange.setDataValidation(rule);
   * ```
   */
  setOptions(options) {
    Object.assign(this._rule, options);
    return this;
  }
};

// ../packages/sheets-data-validation/src/facade/f-data-validation.ts
var FDataValidation = class {
  constructor(rule, worksheet, _injector) {
    __publicField(this, "rule");
    __publicField(this, "_worksheet");
    __publicField(this, "_injector");
    this._injector = _injector;
    this.rule = rule;
    this._worksheet = worksheet;
  }
  /**
   * Gets whether invalid data is allowed based on the error style value
   * @returns {boolean} true if invalid data is allowed, false otherwise
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const rules = fWorksheet.getDataValidations();
   * rules.forEach((rule) => {
   *   console.log(rule, rule.getAllowInvalid());
   * });
   * ```
   */
  getAllowInvalid() {
    return this.rule.errorStyle !== 1 /* STOP */;
  }
  /**
   * Gets the data validation type of the rule
   * @returns {DataValidationType | string} The data validation type
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const rules = fWorksheet.getDataValidations();
   * rules.forEach((rule) => {
   *   console.log(rule, rule.getCriteriaType());
   * });
   * ```
   */
  getCriteriaType() {
    return this.rule.type;
  }
  /**
   * Gets the values used for criteria evaluation
   * @returns {[string | undefined, string | undefined, string | undefined]} An array containing the operator, formula1, and formula2 values
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const rules = fWorksheet.getDataValidations();
   * rules.forEach((rule) => {
   *   console.log(rule);
   *   const criteriaValues = rule.getCriteriaValues();
   *   const [operator, formula1, formula2] = criteriaValues;
   *   console.log(operator, formula1, formula2);
   * });
   * ```
   */
  getCriteriaValues() {
    return [this.rule.operator, this.rule.formula1, this.rule.formula2];
  }
  /**
   * Gets the help text information, which is used to provide users with guidance and support
   * @returns {string | undefined} Returns the help text information. If there is no error message, it returns an undefined value
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const fRange = fWorksheet.getRange('A1:B10');
   * const rule = univerAPI.newDataValidation()
   *   .requireNumberBetween(1, 10)
   *   .setOptions({
   *     allowBlank: true,
   *     showErrorMessage: true,
   *     error: 'Please enter a number between 1 and 10'
   *   })
   *   .build();
   * fRange.setDataValidation(rule);
   * console.log(fRange.getDataValidation().getHelpText()); // 'Please enter a number between 1 and 10'
   * ```
   */
  getHelpText() {
    return this.rule.error;
  }
  /**
   * Creates a new instance of FDataValidationBuilder using the current rule object
   * @returns {FDataValidationBuilder} A new FDataValidationBuilder instance with the same rule configuration
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const fRange = fWorksheet.getRange('A1:B10');
   * const rule = univerAPI.newDataValidation()
   *   .requireNumberBetween(1, 10)
   *   .setOptions({
   *     allowBlank: true,
   *     showErrorMessage: true,
   *     error: 'Please enter a number between 1 and 10'
   *   })
   *   .build();
   * fRange.setDataValidation(rule);
   *
   * const builder = fRange.getDataValidation().copy();
   * const newRule = builder
   *   .requireNumberBetween(1, 5)
   *   .setOptions({
   *     error: 'Please enter a number between 1 and 5'
   *   })
   *   .build();
   * fRange.setDataValidation(newRule);
   * ```
   */
  copy() {
    return new FDataValidationBuilder(this.rule);
  }
  /**
   * Gets whether the data validation rule is applied to the worksheet
   * @returns {boolean} true if the rule is applied, false otherwise
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const rules = fWorksheet.getDataValidations();
   * rules.forEach((rule) => {
   *   console.log(rule, rule.getApplied());
   * });
   *
   * const fRange = fWorksheet.getRange('A1:B10');
   * console.log(fRange.getDataValidation()?.getApplied());
   * ```
   */
  getApplied() {
    if (!this._worksheet) {
      return false;
    }
    const dataValidationModel = this._injector.get(DataValidationModel);
    const currentRule = dataValidationModel.getRuleById(this._worksheet.getUnitId(), this._worksheet.getSheetId(), this.rule.uid);
    if (currentRule && currentRule.ranges.length) {
      return true;
    }
    return false;
  }
  /**
   * Gets the ranges to which the data validation rule is applied
   * @returns {FRange[]} An array of FRange objects representing the ranges to which the data validation rule is applied
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const rules = fWorksheet.getDataValidations();
   * rules.forEach((rule) => {
   *   console.log(rule);
   *   const ranges = rule.getRanges();
   *   ranges.forEach((range) => {
   *     console.log(range.getA1Notation());
   *   });
   * });
   * ```
   */
  getRanges() {
    if (!this.getApplied()) {
      return [];
    }
    const workbook = this._injector.get(IUniverInstanceService).getUnit(this._worksheet.getUnitId());
    return this.rule.ranges.map((range) => this._injector.createInstance(FRange, workbook, this._worksheet, range));
  }
  /**
   * Gets the unit ID of the worksheet
   * @returns {string | undefined} The unit ID of the worksheet
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const fRange = fWorksheet.getRange('A1:B10');
   * console.log(fRange.getDataValidation().getUnitId());
   * ```
   */
  getUnitId() {
    var _a;
    return (_a = this._worksheet) == null ? void 0 : _a.getUnitId();
  }
  /**
   * Gets the sheet ID of the worksheet
   * @returns {string | undefined} The sheet ID of the worksheet
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const fRange = fWorksheet.getRange('A1:B10');
   * console.log(fRange.getDataValidation().getSheetId());
   * ```
   */
  getSheetId() {
    var _a;
    return (_a = this._worksheet) == null ? void 0 : _a.getSheetId();
  }
  /**
   * Set Criteria for the data validation rule
   * @param {DataValidationType} type - The type of data validation criteria
   * @param {[DataValidationOperator, string, string]} values - An array containing the operator, formula1, and formula2 values
   * @param {boolean} [allowBlank] - Whether to allow blank values
   * @returns {FDataValidation} The current instance for method chaining
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a new data validation rule that requires a number equal to 20 for the range A1:B10
   * const fRange = fWorksheet.getRange('A1:B10');
   * const rule = univerAPI.newDataValidation()
   *   .requireNumberEqualTo(20)
   *   .build();
   * fRange.setDataValidation(rule);
   *
   * // Change the rule criteria to require a number between 1 and 10
   * fRange.getDataValidation().setCriteria(
   *   univerAPI.Enum.DataValidationType.DECIMAL,
   *   [univerAPI.Enum.DataValidationOperator.BETWEEN, '1', '10']
   * );
   * ```
   */
  setCriteria(type, values, allowBlank = true) {
    if (this.getApplied()) {
      const commandService = this._injector.get(ICommandService);
      const res = commandService.syncExecuteCommand(UpdateSheetDataValidationSettingCommand.id, {
        unitId: this.getUnitId(),
        subUnitId: this.getSheetId(),
        ruleId: this.rule.uid,
        setting: {
          operator: values[0],
          formula1: values[1],
          formula2: values[2],
          type: this.rule.type,
          allowBlank
        }
      });
      if (!res) {
        throw new Error("setCriteria failed");
      }
    }
    this.rule.operator = values[0];
    this.rule.formula1 = values[1];
    this.rule.formula2 = values[2];
    this.rule.type = type;
    this.rule.allowBlank = allowBlank;
    return this;
  }
  /**
   * Set the options for the data validation rule
   * @param {Partial<IDataValidationRuleOptions>} options - The options to set for the data validation rule
   * @returns {FDataValidation} The current instance for method chaining
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a new data validation rule that requires a number equal to 20 for the range A1:B10
   * const fRange = fWorksheet.getRange('A1:B10');
   * const rule = univerAPI.newDataValidation()
   *   .requireNumberEqualTo(20)
   *   .build();
   * fRange.setDataValidation(rule);
   *
   * // Supplement the rule with additional options
   * fRange.getDataValidation().setOptions({
   *   allowBlank: true,
   *   showErrorMessage: true,
   *   error: 'Please enter a valid value'
   * });
   * ```
   */
  setOptions(options) {
    if (this.getApplied()) {
      const commandService = this._injector.get(ICommandService);
      const res = commandService.syncExecuteCommand(UpdateSheetDataValidationOptionsCommand.id, {
        unitId: this.getUnitId(),
        subUnitId: this.getSheetId(),
        ruleId: this.rule.uid,
        options: {
          ...getRuleOptions(this.rule),
          ...options
        }
      });
      if (!res) {
        throw new Error("setOptions failed");
      }
    }
    Object.assign(this.rule, options);
    return this;
  }
  /**
   * Set the ranges to the data validation rule
   * @param {FRange[]} ranges - New ranges array
   * @returns {FDataValidation} The current instance for method chaining
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a new data validation rule that requires a number equal to 20 for the range A1:B10
   * const fRange = fWorksheet.getRange('A1:B10');
   * const rule = univerAPI.newDataValidation()
   *   .requireNumberEqualTo(20)
   *   .build();
   * fRange.setDataValidation(rule);
   *
   * // Change the range to C1:D10
   * const newRuleRange = fWorksheet.getRange('C1:D10');
   * fRange.getDataValidation().setRanges([newRuleRange]);
   * ```
   */
  setRanges(ranges) {
    if (this.getApplied()) {
      const commandService = this._injector.get(ICommandService);
      const res = commandService.syncExecuteCommand(UpdateSheetDataValidationRangeCommand.id, {
        unitId: this.getUnitId(),
        subUnitId: this.getSheetId(),
        ruleId: this.rule.uid,
        ranges: ranges.map((range) => range.getRange())
      });
      if (!res) {
        throw new Error("setRanges failed");
      }
    }
    this.rule.ranges = ranges.map((range) => range.getRange());
    return this;
  }
  /**
   * Delete the data validation rule from the worksheet
   * @returns {boolean} true if the rule is deleted successfully, false otherwise
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a new data validation rule that requires a number equal to 20 for the range A1:B10
   * const fRange = fWorksheet.getRange('A1:B10');
   * const rule = univerAPI.newDataValidation()
   *   .requireNumberEqualTo(20)
   *   .build();
   * fRange.setDataValidation(rule);
   *
   * // Delete the data validation rule
   * fRange.getDataValidation().delete();
   * ```
   */
  delete() {
    if (!this.getApplied()) {
      return false;
    }
    const commandService = this._injector.get(ICommandService);
    return commandService.syncExecuteCommand(RemoveSheetDataValidationCommand.id, {
      unitId: this.getUnitId(),
      subUnitId: this.getSheetId(),
      ruleId: this.rule.uid
    });
  }
};

// ../packages/sheets-data-validation/src/facade/f-range.ts
var FRangeSheetsDataValidationMixin = class extends FRange {
  setDataValidation(rule) {
    if (!rule) {
      this._commandService.syncExecuteCommand(ClearRangeDataValidationCommand.id, {
        unitId: this._workbook.getUnitId(),
        subUnitId: this._worksheet.getSheetId(),
        ranges: [this._range]
      });
      return this;
    }
    const params = {
      unitId: this._workbook.getUnitId(),
      subUnitId: this._worksheet.getSheetId(),
      rule: {
        ...rule.rule,
        ranges: [this._range]
      }
    };
    this._commandService.syncExecuteCommand(AddSheetDataValidationCommand.id, params);
    return this;
  }
  getDataValidation() {
    const validatorService = this._injector.get(SheetsDataValidationValidatorService);
    const rule = validatorService.getDataValidation(
      this._workbook.getUnitId(),
      this._worksheet.getSheetId(),
      [this._range]
    );
    if (rule) {
      return new FDataValidation(rule, this._worksheet, this._injector);
    }
    return rule;
  }
  getDataValidations() {
    const validatorService = this._injector.get(SheetsDataValidationValidatorService);
    return validatorService.getDataValidations(
      this._workbook.getUnitId(),
      this._worksheet.getSheetId(),
      [this._range]
    ).map((rule) => new FDataValidation(rule, this._worksheet, this._injector));
  }
  async getValidatorStatus() {
    const validatorService = this._injector.get(SheetsDataValidationValidatorService);
    return validatorService.validatorRanges(
      this._workbook.getUnitId(),
      this._worksheet.getSheetId(),
      [this._range]
    );
  }
  async getDataValidationErrorAsync() {
    const unitId = this._workbook.getUnitId();
    const sheetId = this._worksheet.getSheetId();
    return this._collectValidationErrorsForRange(unitId, sheetId, [this._range]);
  }
  async _collectValidationErrorsForRange(unitId, sheetId, ranges) {
    if (!ranges.length) {
      return [];
    }
    const validatorService = this._injector.get(SheetsDataValidationValidatorService);
    const worksheet = this._worksheet;
    const sheetName = worksheet.getName();
    const errors = [];
    for (const range of ranges) {
      const promises = [];
      for (let row = range.startRow; row <= range.endRow; row++) {
        for (let col = range.startColumn; col <= range.endColumn; col++) {
          promises.push((async () => {
            var _a;
            try {
              const status = await validatorService.validatorCell(unitId, sheetId, row, col);
              if (status !== "valid" /* VALID */) {
                const dataValidationModel = this._injector.get(SheetDataValidationModel);
                const rule = dataValidationModel.getRuleByLocation(unitId, sheetId, row, col);
                if (rule) {
                  const cellValue = ((_a = worksheet.getCell(row, col)) == null ? void 0 : _a.v) || null;
                  const error = this._createDataValidationError(
                    sheetName,
                    row,
                    col,
                    rule,
                    cellValue
                  );
                  errors.push(error);
                }
              }
            } catch (e) {
              console.warn(`Failed to validate cell [${row}, ${col}]:`, e);
            }
          })());
        }
      }
      await Promise.all(promises);
    }
    return errors;
  }
  _createDataValidationError(sheetName, row, column, rule, inputValue) {
    return {
      sheetName,
      row,
      column,
      ruleId: rule.uid,
      inputValue,
      rule
    };
  }
};
FRange.extend(FRangeSheetsDataValidationMixin);

// ../packages/sheets-data-validation/src/facade/f-univer.ts
var FUniverSheetsDataValidationMixin = class extends FUniver {
  /**
   * @deprecated use `univerAPI.newDataValidation()` as instead.
   * @returns {FDataValidationBuilder} A new instance of the FDataValidationBuilder class
   */
  static newDataValidation() {
    return new FDataValidationBuilder();
  }
  newDataValidation() {
    return new FDataValidationBuilder();
  }
  /**
   * @ignore
   */
  // eslint-disable-next-line max-lines-per-function
  _initialize(injector) {
    const commandService = injector.get(ICommandService);
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.SheetDataValidationChanged,
        () => {
          if (!injector.has(SheetDataValidationModel)) return { dispose: () => {
          } };
          const sheetDataValidationModel = injector.get(SheetDataValidationModel);
          return sheetDataValidationModel.ruleChange$.subscribe((ruleChange) => {
            const { unitId, subUnitId, rule, oldRule, type } = ruleChange;
            const target = this.getSheetCommandTarget({ unitId, subUnitId });
            if (!target) return;
            const { workbook, worksheet } = target;
            const fRule = new FDataValidation(rule, worksheet.getSheet(), this._injector);
            const eventParams = {
              origin: ruleChange,
              worksheet,
              workbook,
              changeType: type,
              oldRule,
              rule: fRule
            };
            this.fireEvent(this.Event.SheetDataValidationChanged, eventParams);
          });
        }
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.SheetDataValidatorStatusChanged,
        () => {
          if (!injector.has(SheetDataValidationModel)) return { dispose: () => {
          } };
          const sheetDataValidationModel = injector.get(SheetDataValidationModel);
          return sheetDataValidationModel.validStatusChange$.subscribe((statusChange) => {
            const { unitId, subUnitId, ruleId, status, row, col } = statusChange;
            const target = this.getSheetCommandTarget({ unitId, subUnitId });
            if (!target) return;
            const { workbook, worksheet } = target;
            const rule = worksheet.getDataValidation(ruleId);
            if (!rule) return;
            const eventParams = {
              workbook,
              worksheet,
              row,
              column: col,
              rule,
              status
            };
            this.fireEvent(this.Event.SheetDataValidatorStatusChanged, eventParams);
          });
        }
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeSheetDataValidationAdd,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          if (commandInfo.id === AddSheetDataValidationCommand.id) {
            const params = commandInfo.params;
            const target = this.getSheetCommandTarget(params);
            if (!target) return;
            const { workbook, worksheet } = target;
            const { rule } = params;
            const eventParams = {
              worksheet,
              workbook,
              rule
            };
            this.fireEvent(this.Event.BeforeSheetDataValidationAdd, eventParams);
            if (eventParams.cancel) {
              throw new CanceledError();
            }
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeSheetDataValidationCriteriaUpdate,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          if (commandInfo.id === UpdateSheetDataValidationSettingCommand.id) {
            const params = commandInfo.params;
            const target = this.getSheetCommandTarget(params);
            if (!target) return;
            const { workbook, worksheet } = target;
            const { ruleId, setting: newCriteria } = params;
            const rule = worksheet.getDataValidation(ruleId);
            if (!rule) return;
            const eventParams = {
              worksheet,
              workbook,
              rule,
              ruleId,
              newCriteria
            };
            this.fireEvent(this.Event.BeforeSheetDataValidationCriteriaUpdate, eventParams);
            if (eventParams.cancel) {
              throw new CanceledError();
            }
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeSheetDataValidationRangeUpdate,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          if (commandInfo.id === UpdateSheetDataValidationRangeCommand.id) {
            const params = commandInfo.params;
            const target = this.getSheetCommandTarget(params);
            if (!target) return;
            const { workbook, worksheet } = target;
            const { ruleId, ranges: newRanges } = params;
            const rule = worksheet.getDataValidation(ruleId);
            if (!rule) return;
            const eventParams = {
              worksheet,
              workbook,
              rule,
              ruleId,
              newRanges
            };
            this.fireEvent(this.Event.BeforeSheetDataValidationRangeUpdate, eventParams);
            if (eventParams.cancel) {
              throw new CanceledError();
            }
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeSheetDataValidationOptionsUpdate,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          if (commandInfo.id === UpdateSheetDataValidationOptionsCommand.id) {
            const params = commandInfo.params;
            const target = this.getSheetCommandTarget(params);
            if (!target) return;
            const { workbook, worksheet } = target;
            const { ruleId, options: newOptions } = params;
            const rule = worksheet.getDataValidation(ruleId);
            if (!rule) return;
            const eventParams = {
              worksheet,
              workbook,
              rule,
              ruleId,
              newOptions
            };
            this.fireEvent(this.Event.BeforeSheetDataValidationOptionsUpdate, eventParams);
            if (eventParams.cancel) {
              throw new CanceledError();
            }
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeSheetDataValidationDelete,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          if (commandInfo.id === RemoveSheetDataValidationCommand.id) {
            const params = commandInfo.params;
            const target = this.getSheetCommandTarget(params);
            if (!target) return;
            const { workbook, worksheet } = target;
            const { ruleId } = params;
            const rule = worksheet.getDataValidation(ruleId);
            if (!rule) return;
            const eventParams = {
              worksheet,
              workbook,
              rule,
              ruleId
            };
            this.fireEvent(this.Event.BeforeSheetDataValidationDelete, eventParams);
            if (eventParams.cancel) {
              throw new CanceledError();
            }
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeSheetDataValidationDeleteAll,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          if (commandInfo.id === RemoveSheetAllDataValidationCommand.id) {
            const params = commandInfo.params;
            const target = this.getSheetCommandTarget(params);
            if (!target) return;
            const { workbook, worksheet } = target;
            const rules = worksheet.getDataValidations();
            const eventParams = {
              worksheet,
              workbook,
              rules
            };
            this.fireEvent(this.Event.BeforeSheetDataValidationDeleteAll, eventParams);
            if (eventParams.cancel) {
              throw new CanceledError();
            }
          }
        })
      )
    );
  }
};
FUniver.extend(FUniverSheetsDataValidationMixin);

// ../packages/sheets-data-validation/src/facade/f-workbook.ts
var FWorkbookSheetsDataValidationMixin = class extends FWorkbook {
  _initialize() {
    Object.defineProperty(this, "_dataValidationModel", {
      get() {
        return this._injector.get(SheetDataValidationModel);
      }
    });
  }
  getValidatorStatus() {
    const validatorService = this._injector.get(SheetsDataValidationValidatorService);
    return validatorService.validatorWorkbook(this._workbook.getUnitId());
  }
  async getAllDataValidationErrorAsync() {
    const unitId = this._workbook.getUnitId();
    const sheetIds = this._dataValidationModel.getSubUnitIds(unitId);
    const allErrors = [];
    for (const sheetId of sheetIds) {
      const sheetErrors = await this._collectValidationErrorsForSheet(unitId, sheetId);
      allErrors.push(...sheetErrors);
    }
    return allErrors;
  }
  async _collectValidationErrorsForSheet(unitId, sheetId) {
    const rules = this._dataValidationModel.getRules(unitId, sheetId);
    if (!rules.length) {
      return [];
    }
    const allRanges = rules.flatMap((rule) => rule.ranges);
    return this._collectValidationErrorsForRange(unitId, sheetId, allRanges);
  }
  async _collectValidationErrorsForRange(unitId, sheetId, ranges) {
    if (!ranges.length) {
      return [];
    }
    const validatorService = this._injector.get(SheetsDataValidationValidatorService);
    const workbook = this._workbook;
    const worksheet = workbook.getSheetBySheetId(sheetId);
    if (!worksheet) {
      throw new Error(`Cannot find worksheet with sheetId: ${sheetId}`);
    }
    const sheetName = worksheet.getName();
    const errors = [];
    for (const range of ranges) {
      const promises = [];
      for (let row = range.startRow; row <= range.endRow; row++) {
        for (let col = range.startColumn; col <= range.endColumn; col++) {
          promises.push((async () => {
            var _a;
            try {
              const status = await validatorService.validatorCell(unitId, sheetId, row, col);
              if (status !== "valid" /* VALID */) {
                const rule = this._dataValidationModel.getRuleByLocation(unitId, sheetId, row, col);
                if (rule) {
                  const cellValue = ((_a = worksheet.getCell(row, col)) == null ? void 0 : _a.v) || null;
                  const error = this._createDataValidationError(
                    sheetName,
                    row,
                    col,
                    rule,
                    cellValue
                  );
                  errors.push(error);
                }
              }
            } catch (e) {
              console.warn(`Failed to validate cell [${row}, ${col}]:`, e);
            }
          })());
        }
      }
      await Promise.all(promises);
    }
    return errors;
  }
  _createDataValidationError(sheetName, row, column, rule, inputValue) {
    return {
      sheetName,
      row,
      column,
      ruleId: rule.uid,
      inputValue,
      rule
    };
  }
};
FWorkbook.extend(FWorkbookSheetsDataValidationMixin);

// ../packages/sheets-data-validation/src/facade/f-worksheet.ts
var FWorksheetDataValidationMixin = class extends FWorksheet {
  getDataValidations() {
    const dataValidationModel = this._injector.get(DataValidationModel);
    return dataValidationModel.getRules(this._workbook.getUnitId(), this._worksheet.getSheetId()).map((rule) => new FDataValidation(rule, this._worksheet, this._injector));
  }
  getValidatorStatusAsync() {
    const validatorService = this._injector.get(SheetsDataValidationValidatorService);
    return validatorService.validatorWorksheet(
      this._workbook.getUnitId(),
      this._worksheet.getSheetId()
    );
  }
  getDataValidation(ruleId) {
    const dataValidationModel = this._injector.get(DataValidationModel);
    const rule = dataValidationModel.getRuleById(this._workbook.getUnitId(), this._worksheet.getSheetId(), ruleId);
    if (rule) {
      return new FDataValidation(rule, this._worksheet, this._injector);
    }
    return null;
  }
  async getAllDataValidationErrorAsync() {
    const unitId = this._workbook.getUnitId();
    const sheetId = this._worksheet.getSheetId();
    return this._collectValidationErrorsForSheet(unitId, sheetId);
  }
  async _collectValidationErrorsForSheet(unitId, sheetId) {
    const dataValidationModel = this._injector.get(DataValidationModel);
    const rules = dataValidationModel.getRules(unitId, sheetId);
    if (!rules.length) {
      return [];
    }
    const allRanges = rules.flatMap((rule) => rule.ranges);
    return this._collectValidationErrorsForRange(unitId, sheetId, allRanges);
  }
  async _collectValidationErrorsForRange(unitId, sheetId, ranges) {
    if (!ranges.length) {
      return [];
    }
    const validatorService = this._injector.get(SheetsDataValidationValidatorService);
    const worksheet = this._worksheet;
    const sheetName = worksheet.getName();
    const errors = [];
    for (const range of ranges) {
      const promises = [];
      for (let row = range.startRow; row <= range.endRow; row++) {
        for (let col = range.startColumn; col <= range.endColumn; col++) {
          promises.push((async () => {
            var _a;
            try {
              const status = await validatorService.validatorCell(unitId, sheetId, row, col);
              if (status !== "valid" /* VALID */) {
                const dataValidationModel = this._injector.get(SheetDataValidationModel);
                const rule = dataValidationModel.getRuleByLocation(unitId, sheetId, row, col);
                if (rule) {
                  const cellValue = ((_a = worksheet.getCell(row, col)) == null ? void 0 : _a.v) || null;
                  const error = this._createDataValidationError(
                    sheetName,
                    row,
                    col,
                    rule,
                    cellValue
                  );
                  errors.push(error);
                }
              }
            } catch (e) {
              console.warn(`Failed to validate cell [${row}, ${col}]:`, e);
            }
          })());
        }
      }
      await Promise.all(promises);
    }
    return errors;
  }
  _createDataValidationError(sheetName, row, column, rule, inputValue) {
    return {
      sheetName,
      row,
      column,
      ruleId: rule.uid,
      inputValue,
      rule
    };
  }
};
FWorksheet.extend(FWorksheetDataValidationMixin);

// ../packages/sheets-data-validation/src/facade/f-event.ts
var FSheetsDataValidationEventNameMixin = class extends FEventName {
  get SheetDataValidationChanged() {
    return "SheetDataValidationChanged";
  }
  get SheetDataValidatorStatusChanged() {
    return "SheetDataValidatorStatusChanged";
  }
  get BeforeSheetDataValidationAdd() {
    return "BeforeSheetDataValidationAdd";
  }
  get BeforeSheetDataValidationDelete() {
    return "BeforeSheetDataValidationDelete";
  }
  get BeforeSheetDataValidationDeleteAll() {
    return "BeforeSheetDataValidationDeleteAll";
  }
  get BeforeSheetDataValidationCriteriaUpdate() {
    return "BeforeSheetDataValidationCriteriaUpdate";
  }
  get BeforeSheetDataValidationRangeUpdate() {
    return "BeforeSheetDataValidationRangeUpdate";
  }
  get BeforeSheetDataValidationOptionsUpdate() {
    return "BeforeSheetDataValidationOptionsUpdate";
  }
};
FEventName.extend(FSheetsDataValidationEventNameMixin);

// ../packages/sheets-filter/src/facade/f-univer.ts
var FUniverSheetsFilterMixin = class extends FUniver {
  /**
   * @ignore
   */
  _initialize(injector) {
    const commandService = injector.get(ICommandService);
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.SheetBeforeRangeFilter,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          if (commandInfo.id === SetSheetsFilterCriteriaCommand.id) {
            this._beforeRangeFilter(commandInfo);
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.SheetBeforeRangeFilterClear,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          if (commandInfo.id === ClearSheetsFilterCriteriaCommand.id) {
            this._beforeRangeFilterClear();
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.SheetRangeFiltered,
        () => commandService.onCommandExecuted((commandInfo) => {
          if (commandInfo.id === SetSheetsFilterCriteriaCommand.id) {
            this._onRangeFiltered(commandInfo);
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.SheetRangeFilterCleared,
        () => commandService.onCommandExecuted((commandInfo) => {
          if (commandInfo.id === ClearSheetsFilterCriteriaCommand.id) {
            this._onRangeFilterCleared();
          }
        })
      )
    );
  }
  _beforeRangeFilter(commandInfo) {
    const params = commandInfo.params;
    const fWorkbook = this.getWorkbook(params.unitId);
    const eventParams = {
      workbook: fWorkbook,
      worksheet: fWorkbook.getSheetBySheetId(params.subUnitId),
      col: params.col,
      criteria: params.criteria
    };
    this.fireEvent(this.Event.SheetBeforeRangeFilter, eventParams);
    if (eventParams.cancel) {
      throw new Error("SetSheetsFilterCriteriaCommand canceled.");
    }
  }
  _onRangeFiltered(commandInfo) {
    const params = commandInfo.params;
    const fWorkbook = this.getWorkbook(params.unitId);
    const eventParams = {
      workbook: fWorkbook,
      worksheet: fWorkbook.getSheetBySheetId(params.subUnitId),
      col: params.col,
      criteria: params.criteria
    };
    this.fireEvent(this.Event.SheetRangeFiltered, eventParams);
  }
  _beforeRangeFilterClear() {
    const fWorkbook = this.getActiveWorkbook();
    if (!fWorkbook) return;
    const eventParams = {
      workbook: fWorkbook,
      worksheet: fWorkbook.getActiveSheet()
    };
    this.fireEvent(this.Event.SheetBeforeRangeFilterClear, eventParams);
    if (eventParams.cancel) {
      throw new Error("SetSheetsFilterCriteriaCommand canceled.");
    }
  }
  _onRangeFilterCleared() {
    const fWorkbook = this.getActiveWorkbook();
    if (!fWorkbook) return;
    const eventParams = {
      workbook: fWorkbook,
      worksheet: fWorkbook.getActiveSheet()
    };
    this.fireEvent(this.Event.SheetRangeFilterCleared, eventParams);
  }
};
FUniver.extend(FUniverSheetsFilterMixin);

// ../packages/sheets-filter/src/facade/f-filter.ts
var FFilter = class {
  constructor(_workbook, _worksheet, _filterModel, _injector, _commandSrv) {
    __publicField(this, "_workbook", _workbook);
    __publicField(this, "_worksheet", _worksheet);
    __publicField(this, "_filterModel", _filterModel);
    __publicField(this, "_injector", _injector);
    __publicField(this, "_commandSrv", _commandSrv);
  }
  /**
   * Get the filtered out rows by this filter.
   * @returns {number[]} Filtered out rows by this filter.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Set some values of the range C1:F10
   * const fRange = fWorksheet.getRange('C1:F10');
   * fRange.setValues([
   *   [1, 2, 3, 4],
   *   [2, 3, 4, 5],
   *   [3, 4, 5, 6],
   *   [4, 5, 6, 7],
   *   [5, 6, 7, 8],
   *   [6, 7, 8, 9],
   *   [7, 8, 9, 10],
   *   [8, 9, 10, 11],
   *   [9, 10, 11, 12],
   *   [10, 11, 12, 13],
   * ]);
   *
   * // Create a filter on the range C1:F10
   * let fFilter = fRange.createFilter();
   *
   * // If the filter already exists, remove it and create a new one
   * if (!fFilter) {
   *   fRange.getFilter().remove();
   *   fFilter = fRange.createFilter();
   * }
   *
   * // Set the filter criteria of the column C, filter out the rows that are not 1, 5, 9
   * const column = fWorksheet.getRange('C:C').getColumn();
   * fFilter.setColumnFilterCriteria(column, {
   *   colId: 0,
   *   filters: {
   *     filters: ['1', '5', '9'],
   *   },
   * });
   *
   * // Get the filtered out rows
   * console.log(fFilter.getFilteredOutRows()); // [1, 2, 3, 5, 6, 7, 9]
   * ```
   */
  getFilteredOutRows() {
    return Array.from(this._filterModel.filteredOutRows).sort();
  }
  /**
   * Get the filter criteria of a column.
   * @param {number} column - The column index.
   * @returns {Nullable<IFilterColumn>} The filter criteria of the column.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Set some values of the range C1:F10
   * const fRange = fWorksheet.getRange('C1:F10');
   * fRange.setValues([
   *   [1, 2, 3, 4],
   *   [2, 3, 4, 5],
   *   [3, 4, 5, 6],
   *   [4, 5, 6, 7],
   *   [5, 6, 7, 8],
   *   [6, 7, 8, 9],
   *   [7, 8, 9, 10],
   *   [8, 9, 10, 11],
   *   [9, 10, 11, 12],
   *   [10, 11, 12, 13],
   * ]);
   *
   * // Create a filter on the range C1:F10
   * let fFilter = fRange.createFilter();
   *
   * // If the filter already exists, remove it and create a new one
   * if (!fFilter) {
   *   fRange.getFilter().remove();
   *   fFilter = fRange.createFilter();
   * }
   *
   * // Set the filter criteria of the column C, filter out the rows that are not 1, 5, 9
   * const column = fWorksheet.getRange('C:C').getColumn();
   * fFilter.setColumnFilterCriteria(column, {
   *   colId: 0,
   *   filters: {
   *     filters: ['1', '5', '9'],
   *   },
   * });
   *
   * // Print the filter criteria of the column C and D
   * console.log(fFilter.getColumnFilterCriteria(column)); // { colId: 0, filters: { filters: ['1', '5', '9'] } }
   * console.log(fFilter.getColumnFilterCriteria(column + 1)); // undefined
   * ```
   */
  getColumnFilterCriteria(column) {
    var _a;
    return (_a = this._filterModel.getFilterColumn(column)) == null ? void 0 : _a.getColumnData();
  }
  /**
   * Clear the filter criteria of a column.
   * @param {number} column - The column index.
   * @returns {FFilter} The FFilter instance for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Set some values of the range C1:F10
   * const fRange = fWorksheet.getRange('C1:F10');
   * fRange.setValues([
   *   [1, 2, 3, 4],
   *   [2, 3, 4, 5],
   *   [3, 4, 5, 6],
   *   [4, 5, 6, 7],
   *   [5, 6, 7, 8],
   *   [6, 7, 8, 9],
   *   [7, 8, 9, 10],
   *   [8, 9, 10, 11],
   *   [9, 10, 11, 12],
   *   [10, 11, 12, 13],
   * ]);
   *
   * // Create a filter on the range C1:F10
   * let fFilter = fRange.createFilter();
   *
   * // If the filter already exists, remove it and create a new one
   * if (!fFilter) {
   *   fRange.getFilter().remove();
   *   fFilter = fRange.createFilter();
   * }
   *
   * // Set the filter criteria of the column C, filter out the rows that are not 1, 5, 9
   * const column = fWorksheet.getRange('C:C').getColumn();
   * fFilter.setColumnFilterCriteria(column, {
   *   colId: 0,
   *   filters: {
   *     filters: ['1', '5', '9'],
   *   },
   * });
   *
   * // Clear the filter criteria of the column C after 3 seconds
   * setTimeout(() => {
   *   fFilter.removeColumnFilterCriteria(column);
   * }, 3000);
   * ```
   */
  removeColumnFilterCriteria(column) {
    this._commandSrv.syncExecuteCommand(SetSheetsFilterCriteriaCommand.id, {
      unitId: this._workbook.getUnitId(),
      subUnitId: this._worksheet.getSheetId(),
      col: column,
      criteria: null
    });
    return this;
  }
  /**
   * Set the filter criteria of a column.
   * @param {number} column - The column index.
   * @param {ISetSheetsFilterCriteriaCommandParams['criteria']} criteria - The new filter criteria.
   * @returns {FFilter} The FFilter instance for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Set some values of the range C1:F10
   * const fRange = fWorksheet.getRange('C1:F10');
   * fRange.setValues([
   *   [1, 2, 3, 4],
   *   [2, 3, 4, 5],
   *   [3, 4, 5, 6],
   *   [4, 5, 6, 7],
   *   [5, 6, 7, 8],
   *   [6, 7, 8, 9],
   *   [7, 8, 9, 10],
   *   [8, 9, 10, 11],
   *   [9, 10, 11, 12],
   *   [10, 11, 12, 13],
   * ]);
   *
   * // Create a filter on the range C1:F10
   * let fFilter = fRange.createFilter();
   *
   * // If the filter already exists, remove it and create a new one
   * if (!fFilter) {
   *   fRange.getFilter().remove();
   *   fFilter = fRange.createFilter();
   * }
   *
   * // Set the filter criteria of the column C, filter out the rows that are not 1, 5, 9
   * const column = fWorksheet.getRange('C:C').getColumn();
   * fFilter.setColumnFilterCriteria(column, {
   *   colId: 0,
   *   filters: {
   *     filters: ['1', '5', '9'],
   *   },
   * });
   * ```
   */
  setColumnFilterCriteria(column, criteria) {
    this._commandSrv.syncExecuteCommand(SetSheetsFilterCriteriaCommand.id, {
      unitId: this._workbook.getUnitId(),
      subUnitId: this._worksheet.getSheetId(),
      col: column,
      criteria
    });
    return this;
  }
  /**
   * Get the range of the filter.
   * @returns {FRange} The range of the filter.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const fFilter = fWorksheet.getFilter();
   * console.log(fFilter?.getRange().getA1Notation());
   * ```
   */
  getRange() {
    const range = this._filterModel.getRange();
    return this._injector.createInstance(FRange, this._workbook, this._worksheet, range);
  }
  /**
   * Remove the filter criteria of all columns.
   * @returns {FFilter} The FFilter instance for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Set some values of the range C1:F10
   * const fRange = fWorksheet.getRange('C1:F10');
   * fRange.setValues([
   *   [1, 2, 3, 4],
   *   [2, 3, 4, 5],
   *   [3, 4, 5, 6],
   *   [4, 5, 6, 7],
   *   [5, 6, 7, 8],
   *   [6, 7, 8, 9],
   *   [7, 8, 9, 10],
   *   [8, 9, 10, 11],
   *   [9, 10, 11, 12],
   *   [10, 11, 12, 13],
   * ]);
   *
   * // Create a filter on the range C1:F10
   * let fFilter = fRange.createFilter();
   *
   * // If the filter already exists, remove it and create a new one
   * if (!fFilter) {
   *   fRange.getFilter().remove();
   *   fFilter = fRange.createFilter();
   * }
   *
   * // Set the filter criteria of the column C, filter out the rows that are not 1, 5, 9
   * const column = fWorksheet.getRange('C:C').getColumn();
   * fFilter.setColumnFilterCriteria(column, {
   *   colId: 0,
   *   filters: {
   *     filters: ['1', '5', '9'],
   *   },
   * });
   *
   * // Clear the filter criteria of all columns after 3 seconds
   * setTimeout(() => {
   *   fFilter.removeFilterCriteria();
   * }, 3000);
   * ```
   */
  removeFilterCriteria() {
    this._commandSrv.syncExecuteCommand(ClearSheetsFilterCriteriaCommand.id);
    return this;
  }
  /**
   * Remove the filter from the worksheet.
   * @returns {boolean} True if the filter is removed successfully; otherwise, false.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const fRange = fWorksheet.getRange('A1:D14');
   * let fFilter = fRange.createFilter();
   *
   * // If the worksheet already has a filter, remove it and create a new filter.
   * if (!fFilter) {
   *   fWorksheet.getFilter().remove();
   *   fFilter = fRange.createFilter();
   * }
   * console.log(fFilter);
   * ```
   */
  remove() {
    return this._commandSrv.syncExecuteCommand(RemoveSheetFilterCommand.id, {
      unitId: this._workbook.getUnitId(),
      subUnitId: this._worksheet.getSheetId()
    });
  }
};
FFilter = __decorateClass([
  __decorateParam(3, Inject(Injector)),
  __decorateParam(4, ICommandService)
], FFilter);

// ../packages/sheets-filter/src/facade/f-range.ts
var FRangeSheetsFilterMixin = class extends FRange {
  createFilter() {
    if (this._getFilterModel()) return null;
    const success = this._commandService.syncExecuteCommand(SetSheetFilterRangeCommand.id, {
      unitId: this._workbook.getUnitId(),
      subUnitId: this._worksheet.getSheetId(),
      range: this._range
    });
    if (!success) return null;
    return this.getFilter();
  }
  /**
   * Get the filter for the current range's worksheet.
   * @returns {FFilter | null} The interface class to handle the filter. If the worksheet does not have a filter,
   * this method would return `null`.
   */
  getFilter() {
    const filterModel = this._getFilterModel();
    if (!filterModel) return null;
    return this._injector.createInstance(FFilter, this._workbook, this._worksheet, filterModel);
  }
  _getFilterModel() {
    return this._injector.get(SheetsFilterService).getFilterModel(
      this._workbook.getUnitId(),
      this._worksheet.getSheetId()
    );
  }
};
FRange.extend(FRangeSheetsFilterMixin);

// ../packages/sheets-filter/src/facade/f-worksheet.ts
var FWorksheetFilterMixin = class extends FWorksheet {
  getFilter() {
    const filterModel = this._getFilterModel();
    if (!filterModel) return null;
    return this._injector.createInstance(FFilter, this._workbook, this._worksheet, filterModel);
  }
  _getFilterModel() {
    return this._injector.get(SheetsFilterService).getFilterModel(
      this._workbook.getUnitId(),
      this._worksheet.getSheetId()
    );
  }
};
FWorksheet.extend(FWorksheetFilterMixin);

// ../packages/sheets-filter/src/facade/f-enum.ts
var FSheetsFilterEnumMixin = class extends FEnum {
  get CustomFilterOperator() {
    return CustomFilterOperator;
  }
};
FEnum.extend(FSheetsFilterEnumMixin);

// ../packages/sheets-filter/src/facade/f-event.ts
var FSheetsFilterEventNameMixin = class extends FEventName {
  get SheetBeforeRangeFilter() {
    return "SheetBeforeRangeFilter";
  }
  get SheetRangeFiltered() {
    return "SheetRangeFiltered";
  }
  get SheetRangeFilterCleared() {
    return "SheetRangeFilterCleared";
  }
  get SheetBeforeRangeFilterClear() {
    return "SheetBeforeRangeFilterClear";
  }
};
FEventName.extend(FSheetsFilterEventNameMixin);

// ../packages/sheets-hyper-link-ui/src/facade/f-workbook.ts
var FWorkbookHyperlinkUIMixin = class extends FWorkbook {
  // TODO: this should be migrated back to hyperlink ui plugin
  navigateToSheetHyperlink(hyperlink) {
    const parserService = this._injector.get(SheetsHyperLinkParserService);
    const resolverService = this._injector.get(SheetsHyperLinkResolverService);
    const info = parserService.parseHyperLink(hyperlink);
    resolverService.navigate(info);
  }
};
FWorkbook.extend(FWorkbookHyperlinkUIMixin);

// ../packages/sheets-thread-comment/src/facade/f-thread-comment.ts
var FTheadCommentItem = class _FTheadCommentItem {
  constructor(comment) {
    __publicField(this, "_comment", {
      id: generateRandomId(),
      ref: "",
      threadId: "",
      dT: "",
      personId: "",
      text: RichTextBuilder.newEmptyData().body,
      attachments: [],
      unitId: "",
      subUnitId: ""
    });
    if (comment) {
      this._comment = comment;
    }
  }
  /**
   * Create a new FTheadCommentItem
   * @param {IThreadComment|undefined} comment The comment
   * @returns {FTheadCommentItem} A new instance of FTheadCommentItem
   * @example
   * ```ts
   * const commentBuilder = univerAPI.newTheadComment();
   * console.log(commentBuilder);
   * ```
   */
  static create(comment) {
    return new _FTheadCommentItem(comment);
  }
  /**
   * Get the person id of the comment
   * @returns {string} The person id of the comment
   * @example
   * ```ts
   * const commentBuilder = univerAPI.newTheadComment();
   * console.log(commentBuilder.personId);
   * ```
   */
  get personId() {
    return this._comment.personId;
  }
  /**
   * Get the date time of the comment
   * @returns {string} The date time of the comment
   * @example
   * ```ts
   * const commentBuilder = univerAPI.newTheadComment();
   * console.log(commentBuilder.dateTime);
   * ```
   */
  get dateTime() {
    return this._comment.dT;
  }
  /**
   * Get the content of the comment
   * @returns {RichTextValue} The content of the comment
   * @example
   * ```ts
   * const commentBuilder = univerAPI.newTheadComment();
   * console.log(commentBuilder.content);
   * ```
   */
  get content() {
    return RichTextValue.createByBody(this._comment.text);
  }
  /**
   * Get the id of the comment
   * @returns {string} The id of the comment
   * @example
   * ```ts
   * const commentBuilder = univerAPI.newTheadComment();
   * console.log(commentBuilder.id);
   * ```
   */
  get id() {
    return this._comment.id;
  }
  /**
   * Get the thread id of the comment
   * @returns {string} The thread id of the comment
   * @example
   * ```ts
   * const commentBuilder = univerAPI.newTheadComment();
   * console.log(commentBuilder.threadId);
   * ```
   */
  get threadId() {
    return this._comment.threadId;
  }
  /**
   * Copy the comment
   * @returns {FTheadCommentBuilder} The comment builder
   * @example
   * ```ts
   * const commentBuilder = univerAPI.newTheadComment();
   * const newCommentBuilder = commentBuilder.copy();
   * console.log(newCommentBuilder);
   * ```
   */
  copy() {
    return FTheadCommentBuilder.create(Tools.deepClone(this._comment));
  }
};
var FTheadCommentBuilder = class _FTheadCommentBuilder extends FTheadCommentItem {
  static create(comment) {
    return new _FTheadCommentBuilder(comment);
  }
  /**
   * Set the content of the comment
   * @param {IDocumentBody | RichTextValue} content The content of the comment
   * @returns {FTheadCommentBuilder} The comment builder for chaining
   * @example
   * ```ts
   * // Create a new comment
   * const richText = univerAPI.newRichText().insertText('hello univer');
   * const commentBuilder = univerAPI.newTheadComment()
   *   .setContent(richText);
   * console.log(commentBuilder.content);
   *
   * // Add the comment to the cell A1
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const cell = fWorksheet.getRange('A1');
   * const result = await cell.addCommentAsync(commentBuilder);
   * console.log(result);
   * ```
   */
  setContent(content) {
    if (content instanceof RichTextValue) {
      this._comment.text = content.getData().body;
    } else {
      this._comment.text = content;
    }
    return this;
  }
  /**
   * Set the person id of the comment
   * @param {string} userId The person id of the comment
   * @returns {FTheadCommentBuilder} The comment builder for chaining
   * @example
   * ```ts
   * // Create a new comment
   * const richText = univerAPI.newRichText().insertText('hello univer');
   * const commentBuilder = univerAPI.newTheadComment()
   *   .setContent(richText)
   *   .setPersonId('mock-user-id');
   * console.log(commentBuilder.personId);
   *
   * // Add the comment to the cell A1
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const cell = fWorksheet.getRange('A1');
   * const result = await cell.addCommentAsync(commentBuilder);
   * console.log(result);
   * ```
   */
  setPersonId(userId) {
    this._comment.personId = userId;
    return this;
  }
  /**
   * Set the date time of the comment
   * @param {Date} date The date time of the comment
   * @returns {FTheadCommentBuilder} The comment builder for chaining
   * @example
   * ```ts
   * // Create a new comment
   * const richText = univerAPI.newRichText().insertText('hello univer');
   * const commentBuilder = univerAPI.newTheadComment()
   *   .setContent(richText)
   *   .setDateTime(new Date('2025-02-21 14:22:22'));
   * console.log(commentBuilder.dateTime);
   *
   * // Add the comment to the cell A1
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const cell = fWorksheet.getRange('A1');
   * const result = await cell.addCommentAsync(commentBuilder);
   * console.log(result);
   * ```
   */
  setDateTime(date) {
    this._comment.dT = getDT(date);
    return this;
  }
  /**
   * Set the id of the comment
   * @param {string} id The id of the comment
   * @returns {FTheadCommentBuilder} The comment builder for chaining
   * @example
   * ```ts
   * // Create a new comment
   * const richText = univerAPI.newRichText().insertText('hello univer');
   * const commentBuilder = univerAPI.newTheadComment()
   *   .setContent(richText)
   *   .setId('mock-comment-id');
   * console.log(commentBuilder.id);
   *
   * // Add the comment to the cell A1
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const cell = fWorksheet.getRange('A1');
   * const result = await cell.addCommentAsync(commentBuilder);
   * console.log(result);
   * ```
   */
  setId(id) {
    this._comment.id = id;
    return this;
  }
  /**
   * Set the thread id of the comment
   * @param {string} threadId The thread id of the comment
   * @returns {FTheadCommentBuilder} The comment builder
   * @example
   * ```ts
   * // Create a new comment
   * const richText = univerAPI.newRichText().insertText('hello univer');
   * const commentBuilder = univerAPI.newTheadComment()
   *   .setContent(richText)
   *   .setThreadId('mock-thread-id');
   * console.log(commentBuilder.threadId);
   *
   * // Add the comment to the cell A1
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const cell = fWorksheet.getRange('A1');
   * const result = await cell.addCommentAsync(commentBuilder);
   * console.log(result);
   * ```
   */
  setThreadId(threadId) {
    this._comment.threadId = threadId;
    return this;
  }
  /**
   * Build the comment
   * @returns {IThreadComment} The comment
   * @example
   * ```ts
   * const richText = univerAPI.newRichText().insertText('hello univer');
   * const comment = univerAPI.newTheadComment()
   *   .setContent(richText)
   *   .setPersonId('mock-user-id')
   *   .setDateTime(new Date('2025-02-21 14:22:22'))
   *   .setId('mock-comment-id')
   *   .setThreadId('mock-thread-id')
   *   .build();
   * console.log(comment);
   * ```
   */
  build() {
    return this._comment;
  }
};
var FThreadComment = class {
  /**
   * @ignore
   */
  constructor(_thread, _parent, _injector, _commandService, _univerInstanceService, _threadCommentModel, _userManagerService) {
    __publicField(this, "_thread", _thread);
    __publicField(this, "_parent", _parent);
    __publicField(this, "_injector", _injector);
    __publicField(this, "_commandService", _commandService);
    __publicField(this, "_univerInstanceService", _univerInstanceService);
    __publicField(this, "_threadCommentModel", _threadCommentModel);
    __publicField(this, "_userManagerService", _userManagerService);
  }
  _getRef() {
    var _a;
    const ref = ((_a = this._parent) == null ? void 0 : _a.ref) || this._thread.ref;
    const range = deserializeRangeWithSheet(ref);
    return range.range;
  }
  /**
   * Whether the comment is a root comment
   * @returns {boolean} Whether the comment is a root comment
   * @example
   * ```ts
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const comments = fWorksheet.getComments();
   * comments.forEach((comment) => {
   *   console.log(comment.getIsRoot());
   * });
   * ```
   */
  getIsRoot() {
    return !this._parent;
  }
  /**
   * Get the comment data
   * @returns {IBaseComment} The comment data
   * @example
   * ```ts
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const comments = fWorksheet.getComments();
   * comments.forEach((comment) => {
   *   console.log(comment.getCommentData());
   * });
   * ```
   */
  getCommentData() {
    const { children, ...comment } = this._thread;
    return comment;
  }
  /**
   * Get the replies of the comment
   * @returns {FThreadComment[]} the replies of the comment
   * @example
   * ```ts
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const comments = fWorksheet.getComments();
   * comments.forEach((comment) => {
   *   if (comment.getIsRoot()) {
   *     const replies = comment.getReplies();
   *     replies.forEach((reply) => {
   *       console.log(reply.getCommentData());
   *     });
   *   }
   * });
   * ```
   */
  getReplies() {
    var _a;
    const range = this._getRef();
    const comments = this._threadCommentModel.getCommentWithChildren(this._thread.unitId, this._thread.subUnitId, range.startRow, range.startColumn);
    return (_a = comments == null ? void 0 : comments.children) == null ? void 0 : _a.map((child) => this._injector.createInstance(FThreadComment, child));
  }
  /**
   * Get the range of the comment
   * @returns {FRange | null} The range of the comment
   * @example
   * ```ts
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const comments = fWorksheet.getComments();
   * comments.forEach((comment) => {
   *   console.log(comment.getRange().getA1Notation());
   * });
   * ```
   */
  getRange() {
    const workbook = this._univerInstanceService.getUnit(this._thread.unitId, 2 /* UNIVER_SHEET */);
    if (!workbook) {
      return null;
    }
    const worksheet = workbook.getSheetBySheetId(this._thread.subUnitId);
    if (!worksheet) {
      return null;
    }
    const range = this._getRef();
    return this._injector.createInstance(FRange, workbook, worksheet, range);
  }
  /**
   * @deprecated use `getRichText` as instead
   */
  getContent() {
    return this._thread.text;
  }
  /**
   * Get the rich text of the comment
   * @returns {RichTextValue} The rich text of the comment
   * @example
   * ```ts
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const comments = fWorksheet.getComments();
   * comments.forEach((comment) => {
   *   console.log(comment.getRichText());
   * });
   * ```
   */
  getRichText() {
    const body = this._thread.text;
    return RichTextValue.create({ body, documentStyle: {}, id: "d" });
  }
  /**
   * Delete the comment and it's replies
   * @returns {Promise<boolean>} Whether the comment is deleted successfully
   * @example
   * ```ts
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const comments = fWorksheet.getComments();
   *
   * // Delete the first comment
   * const result = await comments[0]?.deleteAsync();
   * console.log(result);
   * ```
   */
  deleteAsync() {
    return this._commandService.executeCommand(
      this.getIsRoot() ? DeleteCommentTreeCommand.id : DeleteCommentCommand.id,
      {
        commentId: this._thread.id,
        unitId: this._thread.unitId,
        subUnitId: this._thread.subUnitId
      }
    );
  }
  /**
   * @deprecated use `deleteAsync` as instead.
   */
  delete() {
    return this.deleteAsync();
  }
  /**
   * @deprecated use `updateAsync` as instead
   */
  async update(content) {
    return this.updateAsync(content);
  }
  /**
   * Update the comment content
   * @param {IDocumentBody | RichTextValue} content The new content of the comment
   * @returns {Promise<boolean>} Whether the comment is updated successfully
   * @example
   * ```ts
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a new comment
   * const richText = univerAPI.newRichText().insertText('hello univer');
   * const commentBuilder = univerAPI.newTheadComment()
   *   .setContent(richText)
   *   .setId('mock-comment-id');
   * const cell = fWorksheet.getRange('A1');
   * await cell.addCommentAsync(commentBuilder);
   *
   * // Update the comment after 3 seconds
   * setTimeout(async () => {
   *   const comment = fWorksheet.getCommentById('mock-comment-id');
   *   const newRichText = univerAPI.newRichText().insertText('Hello Univer AI');
   *   const result = await comment.updateAsync(newRichText);
   *   console.log(result);
   * }, 3000);
   * ```
   */
  async updateAsync(content) {
    const body = content instanceof RichTextValue ? content.getData().body : content;
    const dt = getDT();
    const res = await this._commandService.executeCommand(
      UpdateCommentCommand.id,
      {
        unitId: this._thread.unitId,
        subUnitId: this._thread.subUnitId,
        payload: {
          commentId: this._thread.id,
          text: body,
          updated: true,
          updateT: dt
        }
      }
    );
    return res;
  }
  /**
   * @deprecated use `resolveAsync` as instead
   */
  resolve(resolved) {
    return this.resolveAsync(resolved);
  }
  /**
   * Resolve the comment
   * @param {boolean} resolved Whether the comment is resolved
   * @returns {Promise<boolean>} Set the comment to resolved or not operation result
   * @example
   * ```ts
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a new comment
   * const richText = univerAPI.newRichText().insertText('hello univer');
   * const commentBuilder = univerAPI.newTheadComment()
   *   .setContent(richText)
   *   .setId('mock-comment-id');
   * const cell = fWorksheet.getRange('A1');
   * await cell.addCommentAsync(commentBuilder);
   *
   * // Resolve the comment after 3 seconds
   * setTimeout(async () => {
   *   const comment = fWorksheet.getCommentById('mock-comment-id');
   *   const result = await comment.resolveAsync(true);
   *   console.log(result);
   * }, 3000);
   * ```
   */
  resolveAsync(resolved) {
    return this._commandService.executeCommand(
      ResolveCommentCommand.id,
      {
        unitId: this._thread.unitId,
        subUnitId: this._thread.subUnitId,
        commentId: this._thread.id,
        resolved: resolved != null ? resolved : !this._thread.resolved
      }
    );
  }
  /**
   * Reply to the comment
   * @param {FTheadCommentBuilder} comment The comment to reply to
   * @returns {Promise<boolean>} Whether the comment is replied successfully
   * @example
   * ```ts
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a new comment
   * const richText = univerAPI.newRichText().insertText('hello univer');
   * const commentBuilder = univerAPI.newTheadComment()
   *   .setContent(richText)
   *   .setId('mock-comment-id');
   * const cell = fWorksheet.getRange('A1');
   * await cell.addCommentAsync(commentBuilder);
   *
   * // Reply to the comment
   * const replyText = univerAPI.newRichText().insertText('Hello Univer AI');
   * const reply = univerAPI.newTheadComment().setContent(replyText);
   * const comment = fWorksheet.getCommentById('mock-comment-id');
   * const result = await comment.replyAsync(reply);
   * console.log(result);
   * ```
   */
  replyAsync(comment) {
    var _a;
    const commentData = comment.build();
    return this._commandService.executeCommand(
      AddCommentCommand.id,
      {
        unitId: this._thread.unitId,
        subUnitId: this._thread.subUnitId,
        comment: {
          id: generateRandomId(),
          parentId: this._thread.id,
          threadId: this._thread.threadId,
          ref: ((_a = this._parent) == null ? void 0 : _a.ref) || this._thread.ref,
          unitId: this._thread.unitId,
          subUnitId: this._thread.subUnitId,
          text: commentData.text,
          attachments: commentData.attachments,
          dT: commentData.dT || getDT(),
          personId: commentData.personId || this._userManagerService.getCurrentUser().userID
        }
      }
    );
  }
};
FThreadComment = __decorateClass([
  __decorateParam(2, Inject(Injector)),
  __decorateParam(3, ICommandService),
  __decorateParam(4, IUniverInstanceService),
  __decorateParam(5, Inject(SheetsThreadCommentModel)),
  __decorateParam(6, Inject(UserManagerService))
], FThreadComment);

// ../packages/sheets-thread-comment/src/facade/f-range.ts
var FRangeSheetsThreadCommentMixin = class extends FRange {
  getComment() {
    const injector = this._injector;
    const sheetsTheadCommentModel = injector.get(SheetsThreadCommentModel);
    const unitId = this._workbook.getUnitId();
    const sheetId = this._worksheet.getSheetId();
    const commentId = sheetsTheadCommentModel.getByLocation(unitId, sheetId, this._range.startRow, this._range.startColumn);
    if (!commentId) {
      return null;
    }
    const comment = sheetsTheadCommentModel.getComment(unitId, sheetId, commentId);
    if (comment) {
      return this._injector.createInstance(FThreadComment, comment);
    }
    return null;
  }
  getComments() {
    const injector = this._injector;
    const sheetsTheadCommentModel = injector.get(SheetsThreadCommentModel);
    const unitId = this._workbook.getUnitId();
    const sheetId = this._worksheet.getSheetId();
    const comments = [];
    Range.foreach(this._range, (row, col) => {
      const commentId = sheetsTheadCommentModel.getByLocation(unitId, sheetId, row, col);
      if (commentId) {
        const comment = sheetsTheadCommentModel.getComment(unitId, sheetId, commentId);
        if (comment) {
          comments.push(this._injector.createInstance(FThreadComment, comment));
        }
      }
    });
    return comments;
  }
  addComment(content) {
    var _a;
    const injector = this._injector;
    const currentComment = (_a = this.getComment()) == null ? void 0 : _a.getCommentData();
    const commentService = injector.get(ICommandService);
    const userService = injector.get(UserManagerService);
    const unitId = this._workbook.getUnitId();
    const sheetId = this._worksheet.getSheetId();
    const refStr = `${Tools.chatAtABC(this._range.startColumn)}${this._range.startRow + 1}`;
    const currentUser = userService.getCurrentUser();
    const commentData = content instanceof FTheadCommentBuilder ? content.build() : { text: content };
    return commentService.executeCommand(AddCommentCommand.id, {
      unitId,
      subUnitId: sheetId,
      comment: {
        text: commentData.text,
        dT: commentData.dT || getDT(),
        attachments: [],
        id: commentData.id || generateRandomId(),
        ref: refStr,
        personId: commentData.personId || currentUser.userID,
        parentId: currentComment == null ? void 0 : currentComment.id,
        unitId,
        subUnitId: sheetId,
        threadId: (currentComment == null ? void 0 : currentComment.threadId) || generateRandomId()
      }
    });
  }
  clearComment() {
    var _a;
    const injector = this._injector;
    const currentComment = (_a = this.getComment()) == null ? void 0 : _a.getCommentData();
    const commentService = injector.get(ICommandService);
    const unitId = this._workbook.getUnitId();
    const sheetId = this._worksheet.getSheetId();
    if (currentComment) {
      return commentService.executeCommand(DeleteCommentTreeCommand.id, {
        unitId,
        subUnitId: sheetId,
        threadId: currentComment.threadId,
        commentId: currentComment.id
      });
    }
    return Promise.resolve(true);
  }
  clearComments() {
    const comments = this.getComments();
    const promises = comments.map((comment) => comment.deleteAsync());
    return Promise.all(promises).then(() => true);
  }
  addCommentAsync(content) {
    return this.addComment(content);
  }
  clearCommentAsync() {
    return this.clearComment();
  }
  clearCommentsAsync() {
    return this.clearComments();
  }
};
FRange.extend(FRangeSheetsThreadCommentMixin);

// ../packages/sheets-thread-comment/src/facade/f-workbook.ts
var FWorkbookSheetsThreadCommentMixin = class extends FWorkbook {
  /**
   * @ignore
   */
  _initialize() {
    Object.defineProperty(this, "_threadCommentModel", {
      get() {
        return this._injector.get(ThreadCommentModel);
      }
    });
  }
  getComments() {
    return this._threadCommentModel.getUnit(this._workbook.getUnitId()).map((i) => this._injector.createInstance(FThreadComment, i.root));
  }
  clearComments() {
    const comments = this.getComments();
    const promises = comments.map((comment) => comment.deleteAsync());
    return Promise.all(promises).then(() => true);
  }
  /**
   * @param callback
   * @deprecated
   */
  onThreadCommentChange(callback) {
    return toDisposable(this._threadCommentModel.commentUpdate$.pipe(filter((change) => change.unitId === this._workbook.getUnitId())).subscribe(callback));
  }
  /**
   * @param callback
   * @deprecated
   */
  onBeforeAddThreadComment(callback) {
    return toDisposable(this._commandService.beforeCommandExecuted((commandInfo, options) => {
      const params = commandInfo.params;
      if (commandInfo.id === AddCommentCommand.id) {
        if (params.unitId !== this._workbook.getUnitId()) {
          return;
        }
        if (callback(params, options) === false) {
          throw new Error("Command is stopped by the hook onBeforeAddThreadComment");
        }
      }
    }));
  }
  /**
   * @param callback
   * @deprecated
   */
  onBeforeUpdateThreadComment(callback) {
    return toDisposable(this._commandService.beforeCommandExecuted((commandInfo, options) => {
      const params = commandInfo.params;
      if (commandInfo.id === UpdateCommentCommand.id) {
        if (params.unitId !== this._workbook.getUnitId()) {
          return;
        }
        if (callback(params, options) === false) {
          throw new Error("Command is stopped by the hook onBeforeUpdateThreadComment");
        }
      }
    }));
  }
  /**
   * @param callback
   * @deprecated
   */
  onBeforeDeleteThreadComment(callback) {
    return toDisposable(this._commandService.beforeCommandExecuted((commandInfo, options) => {
      const params = commandInfo.params;
      if (commandInfo.id === DeleteCommentCommand.id || commandInfo.id === DeleteCommentTreeCommand.id) {
        if (params.unitId !== this._workbook.getUnitId()) {
          return;
        }
        if (callback(params, options) === false) {
          throw new Error("Command is stopped by the hook onBeforeDeleteThreadComment");
        }
      }
    }));
  }
};
FWorkbook.extend(FWorkbookSheetsThreadCommentMixin);

// ../packages/sheets-thread-comment/src/facade/f-worksheet.ts
var FWorksheetCommentMixin = class extends FWorksheet {
  getComments() {
    const sheetsTheadCommentModel = this._injector.get(SheetsThreadCommentModel);
    const comments = sheetsTheadCommentModel.getSubUnitAll(this._workbook.getUnitId(), this._worksheet.getSheetId());
    return comments.map((comment) => this._injector.createInstance(FThreadComment, comment));
  }
  clearComments() {
    const comments = this.getComments();
    const promises = comments.map((comment) => comment.deleteAsync());
    return Promise.all(promises).then(() => true);
  }
  /**
   * Subscribe to comment events.
   * @param callback Callback function, param contains comment info and target cell.
   */
  onCommented(callback) {
    const commandService = this._injector.get(ICommandService);
    return commandService.onCommandExecuted((command) => {
      if (command.id === AddCommentCommand.id) {
        const params = command.params;
        callback(params);
      }
    });
  }
  getCommentById(commentId) {
    const sheetsTheadCommentModel = this._injector.get(SheetsThreadCommentModel);
    const comment = sheetsTheadCommentModel.getComment(this._workbook.getUnitId(), this._worksheet.getSheetId(), commentId);
    if (comment) {
      return this._injector.createInstance(FThreadComment, comment);
    }
  }
};
FWorksheet.extend(FWorksheetCommentMixin);

// ../packages/sheets-thread-comment/src/facade/f-event.ts
var FSheetsThreadCommentEventNameMixin = class extends FEventName {
  get CommentAdded() {
    return "CommentAdded";
  }
  get BeforeCommentAdd() {
    return "BeforeCommentAdd";
  }
  get CommentUpdated() {
    return "CommentUpdated";
  }
  get BeforeCommentUpdate() {
    return "BeforeCommentUpdate";
  }
  get CommentDeleted() {
    return "CommentDeleted";
  }
  get BeforeCommentDelete() {
    return "BeforeCommentDelete";
  }
  get CommentResolved() {
    return "CommentResolved";
  }
  get BeforeCommentResolve() {
    return "BeforeCommentResolve";
  }
};
FEventName.extend(FSheetsThreadCommentEventNameMixin);

// ../packages/sheets-thread-comment/src/facade/f-univer.ts
var FUniverSheetsThreadCommentMixin = class extends FUniver {
  // eslint-disable-next-line max-lines-per-function
  _initialize(injector) {
    const commandService = injector.get(ICommandService);
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.CommentAdded,
        () => commandService.onCommandExecuted((commandInfo) => {
          var _a, _b, _c, _d;
          if (commandInfo.id !== AddCommentCommand.id) return;
          const params = commandInfo.params;
          const target = this.getSheetCommandTarget(params);
          if (!target) return;
          const { workbook, worksheet } = target;
          const { comment } = params;
          const threadComment = worksheet.getCommentById(comment.id);
          if (!threadComment) return;
          const eventParams = {
            workbook,
            worksheet,
            row: (_b = (_a = threadComment.getRange()) == null ? void 0 : _a.getRow()) != null ? _b : 0,
            col: (_d = (_c = threadComment.getRange()) == null ? void 0 : _c.getColumn()) != null ? _d : 0,
            comment: threadComment
          };
          this.fireEvent(this.Event.CommentAdded, eventParams);
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.CommentUpdated,
        () => commandService.onCommandExecuted((commandInfo) => {
          var _a, _b, _c, _d;
          if (commandInfo.id !== UpdateCommentCommand.id) return;
          const params = commandInfo.params;
          const target = this.getSheetCommandTarget(params);
          if (!target) return;
          const { workbook, worksheet } = target;
          const { payload } = params;
          const threadComment = worksheet.getCommentById(payload.commentId);
          if (!threadComment) return;
          const eventParams = {
            workbook,
            worksheet,
            row: (_b = (_a = threadComment.getRange()) == null ? void 0 : _a.getRow()) != null ? _b : 0,
            col: (_d = (_c = threadComment.getRange()) == null ? void 0 : _c.getColumn()) != null ? _d : 0,
            comment: threadComment
          };
          this.fireEvent(this.Event.CommentUpdated, eventParams);
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.CommentDeleted,
        () => commandService.onCommandExecuted((commandInfo) => {
          if (commandInfo.id !== DeleteCommentCommand.id && commandInfo.id !== DeleteCommentTreeCommand.id) return;
          const params = commandInfo.params;
          const target = this.getSheetCommandTarget(params);
          if (!target) return;
          const { workbook, worksheet } = target;
          const { commentId } = params;
          const eventParams = {
            workbook,
            worksheet,
            commentId
          };
          this.fireEvent(this.Event.CommentDeleted, eventParams);
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.CommentResolved,
        () => commandService.onCommandExecuted((commandInfo) => {
          var _a, _b, _c, _d;
          if (commandInfo.id !== ResolveCommentCommand.id) return;
          const params = commandInfo.params;
          const target = this.getSheetCommandTarget(params);
          if (!target) return;
          const { workbook, worksheet } = target;
          const { commentId, resolved } = params;
          const threadComment = worksheet.getCommentById(commentId);
          if (!threadComment) return;
          const eventParams = {
            workbook,
            worksheet,
            row: (_b = (_a = threadComment.getRange()) == null ? void 0 : _a.getRow()) != null ? _b : 0,
            col: (_d = (_c = threadComment.getRange()) == null ? void 0 : _c.getColumn()) != null ? _d : 0,
            comment: threadComment,
            resolved
          };
          this.fireEvent(this.Event.CommentResolved, eventParams);
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeCommentAdd,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          if (commandInfo.id !== AddCommentCommand.id) return;
          const params = commandInfo.params;
          const target = this.getSheetCommandTarget(params);
          if (!target) return;
          const { workbook, worksheet } = target;
          const { comment } = params;
          const { range } = deserializeRangeWithSheet(comment.ref);
          const eventParams = {
            workbook,
            worksheet,
            row: range.startRow,
            col: range.startColumn,
            comment: FTheadCommentItem.create(comment)
          };
          this.fireEvent(this.Event.BeforeCommentAdd, eventParams);
          if (eventParams.cancel) {
            throw new CanceledError();
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeCommentUpdate,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          var _a, _b, _c, _d;
          if (commandInfo.id !== UpdateCommentCommand.id) return;
          const params = commandInfo.params;
          const target = this.getSheetCommandTarget(params);
          if (!target) return;
          const { workbook, worksheet } = target;
          const { payload } = params;
          const threadComment = worksheet.getCommentById(payload.commentId);
          if (!threadComment) return;
          const eventParams = {
            workbook,
            worksheet,
            row: (_b = (_a = threadComment.getRange()) == null ? void 0 : _a.getRow()) != null ? _b : 0,
            col: (_d = (_c = threadComment.getRange()) == null ? void 0 : _c.getColumn()) != null ? _d : 0,
            comment: threadComment,
            newContent: RichTextValue.createByBody(payload.text)
          };
          this.fireEvent(this.Event.BeforeCommentUpdate, eventParams);
          if (eventParams.cancel) {
            throw new CanceledError();
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeCommentDelete,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          var _a, _b, _c, _d;
          if (commandInfo.id !== DeleteCommentCommand.id && commandInfo.id !== DeleteCommentTreeCommand.id) return;
          const params = commandInfo.params;
          const target = this.getSheetCommandTarget(params);
          if (!target) return;
          const { workbook, worksheet } = target;
          const { commentId } = params;
          const threadComment = worksheet.getCommentById(commentId);
          if (!threadComment) return;
          const eventParams = {
            workbook,
            worksheet,
            row: (_b = (_a = threadComment.getRange()) == null ? void 0 : _a.getRow()) != null ? _b : 0,
            col: (_d = (_c = threadComment.getRange()) == null ? void 0 : _c.getColumn()) != null ? _d : 0,
            comment: threadComment
          };
          this.fireEvent(this.Event.BeforeCommentDelete, eventParams);
          if (eventParams.cancel) {
            throw new CanceledError();
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeCommentResolve,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          var _a, _b;
          if (commandInfo.id !== ResolveCommentCommand.id) return;
          const params = commandInfo.params;
          const target = this.getSheetCommandTarget(params);
          if (!target) return;
          const { workbook, worksheet } = target;
          const { commentId, resolved } = params;
          const threadComment = worksheet.getCommentById(commentId);
          if (!threadComment) return;
          const eventParams = {
            workbook,
            worksheet,
            row: (_a = threadComment.getRange().getRow()) != null ? _a : 0,
            col: (_b = threadComment.getRange().getColumn()) != null ? _b : 0,
            comment: threadComment,
            resolved
          };
          this.fireEvent(this.Event.BeforeCommentResolve, eventParams);
          if (eventParams.cancel) {
            throw new CanceledError();
          }
        })
      )
    );
  }
  /**
   * @ignore
   */
  newTheadComment(comment) {
    return new FTheadCommentBuilder(comment);
  }
};
FUniver.extend(FUniverSheetsThreadCommentMixin);

// ../packages/sheets-conditional-formatting/src/facade/f-conditional-formatting-builder.ts
var ConditionalFormatRuleBaseBuilder = class _ConditionalFormatRuleBaseBuilder {
  constructor(initRule = {}) {
    __publicField(this, "_rule", {});
    this._rule = initRule;
    this._ensureAttr(this._rule, ["rule"]);
  }
  get _ruleConfig() {
    return this._rule.rule || null;
  }
  _getDefaultConfig(type = "highlightCell" /* highlightCell */) {
    switch (type) {
      case "colorScale" /* colorScale */: {
        return {
          type,
          config: [
            { index: 0, color: new ColorKit("").toRgbString(), value: { type: "min" /* min */ } },
            { index: 0, color: new ColorKit("green").toRgbString(), value: { type: "max" /* max */ } }
          ]
        };
      }
      case "dataBar" /* dataBar */: {
        return {
          type,
          isShowValue: true,
          config: { min: { type: "min" /* min */ }, max: { type: "max" /* max */ }, positiveColor: new ColorKit("green").toRgbString(), nativeColor: new ColorKit("").toRgbString(), isGradient: false }
        };
      }
      case "highlightCell" /* highlightCell */: {
        return {
          type,
          subType: "text" /* text */,
          operator: "containsText" /* containsText */,
          value: "abc",
          style: {}
        };
      }
      case "iconSet" /* iconSet */: {
        return {
          type,
          isShowValue: true,
          config: [{
            operator: "greaterThanOrEqual" /* greaterThanOrEqual */,
            value: { type: "min" /* min */ },
            iconType: "EMPTY_ICON_TYPE" /* empty */,
            iconId: ""
          }, {
            operator: "greaterThanOrEqual" /* greaterThanOrEqual */,
            value: { type: "percentile" /* percentile */, value: 0.5 },
            iconType: "EMPTY_ICON_TYPE" /* empty */,
            iconId: ""
          }, {
            operator: "lessThanOrEqual" /* lessThanOrEqual */,
            value: { type: "max" /* max */ },
            iconType: "EMPTY_ICON_TYPE" /* empty */,
            iconId: ""
          }]
        };
      }
    }
  }
  // eslint-disable-next-line ts/no-explicit-any
  _ensureAttr(obj, keys) {
    keys.reduce((pre, cur) => {
      if (!pre[cur]) {
        pre[cur] = {};
      }
      return pre[cur];
    }, obj);
    return obj;
  }
  /**
   * Constructs a conditional format rule from the settings applied to the builder.
   * @returns {IConditionFormattingRule} The conditional format rule.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with no content in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenCellEmpty()
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  build() {
    var _a;
    if (!this._rule.cfId) {
      this._rule.cfId = createCfId();
    }
    if (!this._rule.ranges) {
      this._rule.ranges = [];
    }
    if (this._rule.stopIfTrue === void 0) {
      this._rule.stopIfTrue = false;
    }
    if (!((_a = this._rule.rule) == null ? void 0 : _a.type)) {
      this._rule.rule.type = "highlightCell" /* highlightCell */;
      this._ensureAttr(this._rule, ["rule", "style"]);
    }
    const defaultConfig = this._getDefaultConfig(this._rule.rule.type);
    const result = { ...this._rule, rule: { ...defaultConfig, ...this._rule.rule } };
    return result;
  }
  /**
   * Deep clone a current builder.
   * @returns {ConditionalFormatRuleBaseBuilder} A new builder with the same settings as the original.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with no content in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const builder = fWorksheet.newConditionalFormattingRule()
   *   .whenCellEmpty()
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()]);
   * fWorksheet.addConditionalFormattingRule(builder.build());
   *
   * // Copy the rule and change the background color to green for the range A1:B2.
   * const newRange = fWorksheet.getRange('A1:B2');
   * const newBuilder = builder.copy()
   *   .setBackground('#00FF00')
   *   .setRanges([newRange.getRange()]);
   * fWorksheet.addConditionalFormattingRule(newBuilder.build());
   * ```
   */
  copy() {
    const newRule = Tools.deepClone(this._rule);
    if (newRule.cfId) {
      newRule.cfId = createCfId();
    }
    return new _ConditionalFormatRuleBaseBuilder(newRule);
  }
  /**
   * Gets the scope of the current conditional format.
   * @returns {IRange[]} The ranges to which the conditional format applies.
   */
  getRanges() {
    return this._rule.ranges || [];
  }
  /**
   * Get the icon set mapping dictionary.
   * @returns {Record<string, string[]>} The icon set mapping dictionary.
   */
  getIconMap() {
    return iconMap;
  }
  /**
   * Create a conditional format ID.
   * @returns {string} The conditional format ID.
   */
  createCfId() {
    return createCfId();
  }
  /**
   * Sets the scope for conditional formatting.
   * @param {IRange[]} ranges - The ranges to which the conditional format applies.
   * @returns {ConditionalFormatRuleBaseBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with no content in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenCellEmpty()
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setRanges(ranges) {
    this._rule.ranges = ranges;
    return this;
  }
};
var ConditionalFormatHighlightRuleBuilder = class _ConditionalFormatHighlightRuleBuilder extends ConditionalFormatRuleBaseBuilder {
  constructor(initConfig = {}) {
    super(initConfig);
    this._ensureAttr(this._rule, ["rule", "style"]);
  }
  /**
   * Deep clone a current builder.
   * @returns {ConditionalFormatHighlightRuleBuilder} A new builder with the same settings as the original.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with no content in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const builder = fWorksheet.newConditionalFormattingRule()
   *   .whenCellEmpty()
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()]);
   * fWorksheet.addConditionalFormattingRule(builder.build());
   *
   * // Copy the rule and change the background color to green for the range A1:B2.
   * const newRange = fWorksheet.getRange('A1:B2');
   * const newBuilder = builder.copy()
   *   .setBackground('#00FF00')
   *   .setRanges([newRange.getRange()]);
   * fWorksheet.addConditionalFormattingRule(newBuilder.build());
   * ```
   */
  copy() {
    const newRule = Tools.deepClone(this._rule);
    if (newRule.cfId) {
      newRule.cfId = createCfId();
    }
    return new _ConditionalFormatHighlightRuleBuilder(newRule);
  }
  /**
   * Set average rule.
   * @param {IAverageHighlightCell['operator']} operator - The operator to use for the average rule.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with greater than average values in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .setAverage(univerAPI.Enum.ConditionFormatNumberOperatorEnum.greaterThan)
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setAverage(operator) {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "highlightCell" /* highlightCell */;
    ruleConfig.subType = "average" /* average */;
    ruleConfig.operator = operator;
    return this;
  }
  /**
   * Set unique values rule.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with unique values in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .setUniqueValues()
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setUniqueValues() {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "highlightCell" /* highlightCell */;
    ruleConfig.subType = "uniqueValues" /* uniqueValues */;
    return this;
  }
  /**
   * Set duplicate values rule.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with duplicate values in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .setDuplicateValues()
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setDuplicateValues() {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "highlightCell" /* highlightCell */;
    ruleConfig.subType = "duplicateValues" /* duplicateValues */;
    return this;
  }
  /**
   * Set rank rule.
   * @param {{ isBottom: boolean, isPercent: boolean, value: number }} config - The rank rule settings.
   * @param {boolean} config.isBottom - Whether to highlight the bottom rank.
   * @param {boolean} config.isPercent - Whether to use a percentage rank.
   * @param {number} config.value - The rank value.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights the bottom 10% of values in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .setRank({ isBottom: true, isPercent: true, value: 10 })
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setRank(config) {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "highlightCell" /* highlightCell */;
    ruleConfig.subType = "rank" /* rank */;
    ruleConfig.isBottom = config.isBottom;
    ruleConfig.isPercent = config.isPercent;
    ruleConfig.value = config.value;
    return this;
  }
  /**
   * Sets the background color for the conditional format rule's format.
   * @param {string} [color] - The background color to set. If not provided, the background color is removed.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with no content in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenCellEmpty()
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setBackground(color) {
    var _a;
    if (((_a = this._ruleConfig) == null ? void 0 : _a.type) === "highlightCell" /* highlightCell */) {
      if (color) {
        this._ensureAttr(this._ruleConfig, ["style", "bg"]);
        const colorKit = new ColorKit(color);
        this._ruleConfig.style.bg.rgb = colorKit.toRgbString();
      } else {
        delete this._ruleConfig.style.bg;
      }
    }
    return this;
  }
  /**
   * Sets text bolding for the conditional format rule's format.
   * @param {boolean} isBold - Whether to bold the text.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that bolds the text for cells with not empty content in the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenCellNotEmpty()
   *   .setBold(true)
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setBold(isBold) {
    var _a;
    if (((_a = this._ruleConfig) == null ? void 0 : _a.type) === "highlightCell" /* highlightCell */) {
      this._ensureAttr(this._ruleConfig, ["style"]);
      this._ruleConfig.style.bl = isBold ? 1 /* TRUE */ : 0 /* FALSE */;
    }
    return this;
  }
  /**
   * Sets the font color for the conditional format rule's format.
   * @param {string} [color] - The font color to set. If not provided, the font color is removed.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that changes the font color to red for cells with not empty content in the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenCellNotEmpty()
   *   .setFontColor('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setFontColor(color) {
    var _a;
    if (((_a = this._ruleConfig) == null ? void 0 : _a.type) === "highlightCell" /* highlightCell */) {
      if (color) {
        const colorKit = new ColorKit(color);
        this._ensureAttr(this._ruleConfig, ["style", "cl"]);
        this._ruleConfig.style.cl.rgb = colorKit.toRgbString();
      } else {
        delete this._ruleConfig.style.cl;
      }
    }
    return this;
  }
  /**
   * Sets text italics for the conditional format rule's format.
   * @param {boolean} isItalic - Whether to italicize the text.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that italicizes the text for cells with not empty content in the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenCellNotEmpty()
   *   .setItalic(true)
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setItalic(isItalic) {
    var _a;
    if (((_a = this._ruleConfig) == null ? void 0 : _a.type) === "highlightCell" /* highlightCell */) {
      this._ensureAttr(this._ruleConfig, ["style"]);
      this._ruleConfig.style.it = isItalic ? 1 /* TRUE */ : 0 /* FALSE */;
    }
    return this;
  }
  /**
   * Sets text strikethrough for the conditional format rule's format.
   * @param {boolean} isStrikethrough - Whether is strikethrough the text.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that set text strikethrough for cells with not empty content in the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenCellNotEmpty()
   *   .setStrikethrough(true)
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setStrikethrough(isStrikethrough) {
    var _a;
    if (((_a = this._ruleConfig) == null ? void 0 : _a.type) === "highlightCell" /* highlightCell */) {
      this._ensureAttr(this._ruleConfig, ["style", "st"]);
      this._ruleConfig.style.st.s = isStrikethrough ? 1 /* TRUE */ : 0 /* FALSE */;
    }
    return this;
  }
  /**
   * Sets text underlining for the conditional format rule's format.
   * @param {boolean} isUnderline - Whether to underline the text.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that underlines the text for cells with not empty content in the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenCellNotEmpty()
   *   .setUnderline(true)
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setUnderline(isUnderline) {
    var _a;
    if (((_a = this._ruleConfig) == null ? void 0 : _a.type) === "highlightCell" /* highlightCell */) {
      this._ensureAttr(this._ruleConfig, ["style", "ul"]);
      this._ruleConfig.style.ul.s = isUnderline ? 1 /* TRUE */ : 0 /* FALSE */;
    }
    return this;
  }
  /**
   * Sets the conditional format rule to trigger when the cell is empty.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with no content in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenCellEmpty()
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenCellEmpty() {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "highlightCell" /* highlightCell */;
    ruleConfig.subType = "text" /* text */;
    ruleConfig.value = "";
    ruleConfig.operator = "equal" /* equal */;
    return this;
  }
  /**
   * Sets the conditional format rule to trigger when the cell is not empty.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that changes the font color to red for cells with not empty content in the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenCellNotEmpty()
   *   .setFontColor('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenCellNotEmpty() {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "highlightCell" /* highlightCell */;
    ruleConfig.subType = "text" /* text */;
    ruleConfig.value = "";
    ruleConfig.operator = "notEqual" /* notEqual */;
    return this;
  }
  /**
   * Sets the conditional format rule to trigger when a time period is met.
   * @param {CFTimePeriodOperator} date - The time period to check for.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with dates in the last 7 days in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenDate(univerAPI.Enum.ConditionFormatTimePeriodOperatorEnum.last7Days)
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenDate(date) {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "highlightCell" /* highlightCell */;
    ruleConfig.subType = "timePeriod" /* timePeriod */;
    ruleConfig.operator = date;
    return this;
  }
  /**
   * Sets the conditional format rule to trigger when that the given formula evaluates to `true`.
   * @param {string} formulaString - A custom formula that evaluates to true if the input is valid. formulaString start with '='.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with values greater than 10 in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenFormulaSatisfied('=A1>10')
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenFormulaSatisfied(formulaString) {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "highlightCell" /* highlightCell */;
    ruleConfig.subType = "formula" /* formula */;
    ruleConfig.value = formulaString;
    return this;
  }
  /**
   * Sets the conditional format rule to trigger when a number falls between, or is either of, two specified values.
   * @param {number} start - The lowest acceptable value.
   * @param {number} end - The highest acceptable value.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with values between 10 and 20 in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenNumberBetween(10, 20)
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenNumberBetween(start, end) {
    const min = Math.min(start, end);
    const max = Math.max(start, end);
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "highlightCell" /* highlightCell */;
    ruleConfig.subType = "number" /* number */;
    ruleConfig.value = [min, max];
    ruleConfig.operator = "between" /* between */;
    return this;
  }
  /**
   * Sets the conditional format rule to trigger when a number is equal to the given value.
   * @param {number} value - The sole acceptable value.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with values equal to 10 in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenNumberEqualTo(10)
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenNumberEqualTo(value) {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "highlightCell" /* highlightCell */;
    ruleConfig.subType = "number" /* number */;
    ruleConfig.value = value;
    ruleConfig.operator = "equal" /* equal */;
    return this;
  }
  /**
   * Sets the conditional format rule to trigger when a number is greater than the given value.
   * @param {number} value - The highest unacceptable value.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with values greater than 10 in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenNumberGreaterThan(10)
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenNumberGreaterThan(value) {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "highlightCell" /* highlightCell */;
    ruleConfig.subType = "number" /* number */;
    ruleConfig.value = value;
    ruleConfig.operator = "greaterThan" /* greaterThan */;
    return this;
  }
  /**
   * Sets the conditional format rule to trigger when a number is greater than or equal to the given value.
   * @param {number} value - The lowest acceptable value.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with values greater than or equal to 10 in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenNumberGreaterThanOrEqualTo(10)
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenNumberGreaterThanOrEqualTo(value) {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "highlightCell" /* highlightCell */;
    ruleConfig.subType = "number" /* number */;
    ruleConfig.value = value;
    ruleConfig.operator = "greaterThanOrEqual" /* greaterThanOrEqual */;
    return this;
  }
  /**
   * Sets the conditional conditional format rule to trigger when a number less than the given value.
   * @param {number} value - The lowest unacceptable value.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with values less than 10 in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenNumberLessThan(10)
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenNumberLessThan(value) {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "highlightCell" /* highlightCell */;
    ruleConfig.subType = "number" /* number */;
    ruleConfig.value = value;
    ruleConfig.operator = "lessThan" /* lessThan */;
    return this;
  }
  /**
   * Sets the conditional format rule to trigger when a number less than or equal to the given value.
   * @param {number} value - The highest acceptable value.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with values less than or equal to 10 in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenNumberLessThanOrEqualTo(10)
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenNumberLessThanOrEqualTo(value) {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "highlightCell" /* highlightCell */;
    ruleConfig.subType = "number" /* number */;
    ruleConfig.value = value;
    ruleConfig.operator = "lessThanOrEqual" /* lessThanOrEqual */;
    return this;
  }
  /**
   * Sets the conditional format rule to trigger when a number does not fall between, and is neither of, two specified values.
   * @param {number} start - The lowest unacceptable value.
   * @param {number} end - The highest unacceptable value.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with values not between 10 and 20 in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenNumberNotBetween(10, 20)
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenNumberNotBetween(start, end) {
    const min = Math.min(start, end);
    const max = Math.max(start, end);
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "highlightCell" /* highlightCell */;
    ruleConfig.subType = "number" /* number */;
    ruleConfig.value = [min, max];
    ruleConfig.operator = "notBetween" /* notBetween */;
    return this;
  }
  /**
   * Sets the conditional format rule to trigger when a number is not equal to the given value.
   * @param {number} value - The sole unacceptable value.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with values not equal to 10 in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenNumberNotEqualTo(10)
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenNumberNotEqualTo(value) {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "highlightCell" /* highlightCell */;
    ruleConfig.subType = "number" /* number */;
    ruleConfig.value = value;
    ruleConfig.operator = "notEqual" /* notEqual */;
    return this;
  }
  /**
   * Sets the conditional format rule to trigger when that the input contains the given value.
   * @param {string} text - The value that the input must contain.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with text containing 'apple' in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenTextContains('apple')
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenTextContains(text) {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "highlightCell" /* highlightCell */;
    ruleConfig.subType = "text" /* text */;
    ruleConfig.value = text;
    ruleConfig.operator = "containsText" /* containsText */;
    return this;
  }
  /**
   * Sets the conditional format rule to trigger when that the input does not contain the given value.
   * @param {string} text - The value that the input must not contain.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with text not containing 'apple' in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenTextDoesNotContain('apple')
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenTextDoesNotContain(text) {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "highlightCell" /* highlightCell */;
    ruleConfig.subType = "text" /* text */;
    ruleConfig.value = text;
    ruleConfig.operator = "notContainsText" /* notContainsText */;
    return this;
  }
  /**
   * Sets the conditional format rule to trigger when that the input ends with the given value.
   * @param {string} text - Text to compare against the end of the string.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with text ending with '.ai' in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenTextEndsWith('.ai')
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenTextEndsWith(text) {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "highlightCell" /* highlightCell */;
    ruleConfig.subType = "text" /* text */;
    ruleConfig.value = text;
    ruleConfig.operator = "endsWith" /* endsWith */;
    return this;
  }
  /**
   * Sets the conditional format rule to trigger when that the input is equal to the given value.
   * @param {string} text - The sole acceptable value.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with text equal to 'apple' in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenTextEqualTo('apple')
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenTextEqualTo(text) {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "highlightCell" /* highlightCell */;
    ruleConfig.subType = "text" /* text */;
    ruleConfig.value = text;
    ruleConfig.operator = "equal" /* equal */;
    return this;
  }
  /**
   * Sets the conditional format rule to trigger when that the input starts with the given value.
   * @param {string} text - Text to compare against the beginning of the string.
   * @returns {ConditionalFormatHighlightRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with text starting with 'https://' in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenTextStartsWith('https://')
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenTextStartsWith(text) {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "highlightCell" /* highlightCell */;
    ruleConfig.subType = "text" /* text */;
    ruleConfig.value = text;
    ruleConfig.operator = "beginsWith" /* beginsWith */;
    return this;
  }
};
var ConditionalFormatDataBarRuleBuilder = class _ConditionalFormatDataBarRuleBuilder extends ConditionalFormatRuleBaseBuilder {
  /**
   * Deep clone a current builder.
   * @returns {ConditionalFormatDataBarRuleBuilder} A new instance of the builder with the same settings as the original.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that adds a data bar to cells with values between -100 and 100 in the range A1:D10.
   * // positive values are green and negative values are red.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const builder = fWorksheet.newConditionalFormattingRule()
   *   .setDataBar({
   *     min: { type: 'num', value: -100 },
   *     max: { type: 'num', value: 100 },
   *     positiveColor: '#00FF00',
   *     nativeColor: '#FF0000',
   *     isShowValue: true
   *   })
   *   .setRanges([fRange.getRange()]);
   * fWorksheet.addConditionalFormattingRule(builder.build());
   *
   * // Copy the rule and apply it to a new range.
   * const newRange = fWorksheet.getRange('F1:F10');
   * const newBuilder = builder.copy()
   *   .setRanges([newRange.getRange()]);
   * fWorksheet.addConditionalFormattingRule(newBuilder.build());
   * ```
   */
  copy() {
    const newRule = Tools.deepClone(this._rule);
    if (newRule.cfId) {
      newRule.cfId = createCfId();
    }
    return new _ConditionalFormatDataBarRuleBuilder(newRule);
  }
  /**
   * Set data bar rule.
   * @param {{
   *         min: IValueConfig;
   *         max: IValueConfig;
   *         isGradient?: boolean;
   *         positiveColor: string;
   *         nativeColor: string;
   *         isShowValue?: boolean;
   *     }} config - The data bar rule settings.
   * @param {IValueConfig} config.min - The minimum value for the data bar.
   * @param {IValueConfig} config.max - The maximum value for the data bar.
   * @param {boolean} [config.isGradient] - Whether the data bar is gradient.
   * @param {string} config.positiveColor - The color for positive values.
   * @param {string} config.nativeColor - The color for negative values.
   * @param {boolean} [config.isShowValue] - Whether to show the value in the cell.
   * @returns {ConditionalFormatDataBarRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that adds a data bar to cells with values between -100 and 100 in the range A1:D10.
   * // positive values are green and negative values are red.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .setDataBar({
   *     min: { type: 'num', value: -100 },
   *     max: { type: 'num', value: 100 },
   *     positiveColor: '#00FF00',
   *     nativeColor: '#FF0000',
   *     isShowValue: true
   *   })
   *  .setRanges([fRange.getRange()])
   * .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setDataBar(config) {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "dataBar" /* dataBar */;
    ruleConfig.isShowValue = !!config.isShowValue;
    ruleConfig.config = {
      min: config.min,
      max: config.max,
      positiveColor: config.positiveColor,
      nativeColor: config.nativeColor,
      isGradient: !!config.isGradient
    };
    return this;
  }
};
var ConditionalFormatColorScaleRuleBuilder = class _ConditionalFormatColorScaleRuleBuilder extends ConditionalFormatRuleBaseBuilder {
  /**
   * Deep clone a current builder.
   * @returns {ConditionalFormatColorScaleRuleBuilder} A new instance of the builder with the same settings as the original.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that adds a color scale to cells with values between 0 and 100 in the range A1:D10.
   * // The color scale is green for 0, yellow for 50, and red for 100.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const builder = fWorksheet.newConditionalFormattingRule()
   *   .setColorScale([
   *     { index: 0, color: '#00FF00', value: { type: 'num', value: 0 } },
   *     { index: 1, color: '#FFFF00', value: { type: 'num', value: 50 } },
   *     { index: 2, color: '#FF0000', value: { type: 'num', value: 100 } }
   *   ])
   *   .setRanges([fRange.getRange()]);
   * fWorksheet.addConditionalFormattingRule(builder.build());
   *
   * // Copy the rule and apply it to a new range.
   * const newRange = fWorksheet.getRange('F1:F10');
   * const newBuilder = builder.copy()
   *   .setRanges([newRange.getRange()]);
   * fWorksheet.addConditionalFormattingRule(newBuilder.build());
   * ```
   */
  copy() {
    const newRule = Tools.deepClone(this._rule);
    if (newRule.cfId) {
      newRule.cfId = createCfId();
    }
    return new _ConditionalFormatColorScaleRuleBuilder(newRule);
  }
  /**
   * Set color scale rule.
   * @param {{ index: number; color: string; value: IValueConfig }[]} config - The color scale rule settings.
   * @param {number} config.index - The index of the color scale configuration.
   * @param {string} config.color - The color corresponding to the index of the color scale configuration.
   * @param {IValueConfig} config.value - The condition value corresponding to the index of the color scale configuration.
   * @returns {ConditionalFormatColorScaleRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that adds a color scale to cells with values between 0 and 100 in the range A1:D10.
   * // The color scale is green for 0, yellow for 50, and red for 100.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .setColorScale([
   *     { index: 0, color: '#00FF00', value: { type: 'num', value: 0 } },
   *     { index: 1, color: '#FFFF00', value: { type: 'num', value: 50 } },
   *     { index: 2, color: '#FF0000', value: { type: 'num', value: 100 } }
   *   ])
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setColorScale(config) {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "colorScale" /* colorScale */;
    ruleConfig.config = config;
    return this;
  }
};
var ConditionalFormatIconSetRuleBuilder = class _ConditionalFormatIconSetRuleBuilder extends ConditionalFormatRuleBaseBuilder {
  /**
   * Deep clone a current builder.
   * @returns {ConditionalFormatIconSetRuleBuilder} A new instance of the builder with the same settings as the original.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a 3-arrow icon set conditional formatting rule in the range A1:D10.
   * // The first arrow is green for values greater than 20.
   * // The second arrow is yellow for values greater than 10.
   * // The third arrow is red for values less than or equal to 10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const builder = fWorksheet.newConditionalFormattingRule()
   *   .setIconSet({
   *     iconConfigs: [
   *       {
   *         iconType: univerAPI.Enum.ConditionFormatIconSetTypeEnum.threeArrows,
   *         iconId: '0',
   *         operator: univerAPI.Enum.ConditionFormatNumberOperatorEnum.greaterThan,
   *         value: {
   *           type: univerAPI.Enum.ConditionFormatValueTypeEnum.num,
   *           value: 20
   *         }
   *       },
   *       {
   *         iconType: univerAPI.Enum.ConditionFormatIconSetTypeEnum.threeArrows,
   *         iconId: '1',
   *         operator: univerAPI.Enum.ConditionFormatNumberOperatorEnum.greaterThan,
   *         value: {
   *           type: univerAPI.Enum.ConditionFormatValueTypeEnum.num,
   *           value: 10
   *         }
   *       },
   *       {
   *         iconType: univerAPI.Enum.ConditionFormatIconSetTypeEnum.threeArrows,
   *         iconId: '2',
   *         operator: univerAPI.Enum.ConditionFormatNumberOperatorEnum.lessThanOrEqual,
   *         value: {
   *           type: univerAPI.Enum.ConditionFormatValueTypeEnum.num,
   *           value: 10
   *         }
   *       }
   *     ],
   *     isShowValue: true,
   *   })
   *   .setRanges([fRange.getRange()]);
   * fWorksheet.addConditionalFormattingRule(builder.build());
   *
   * // Copy the rule and apply it to a new range.
   * const newRange = fWorksheet.getRange('F1:F10');
   * const newBuilder = builder.copy()
   *   .setRanges([newRange.getRange()]);
   * fWorksheet.addConditionalFormattingRule(newBuilder.build());
   * ```
   */
  copy() {
    const newRule = Tools.deepClone(this._rule);
    if (newRule.cfId) {
      newRule.cfId = createCfId();
    }
    return new _ConditionalFormatIconSetRuleBuilder(newRule);
  }
  /**
   * Set up icon set conditional formatting rule.
   * @param {{ iconConfigs: IIconSet['config'], isShowValue: boolean }} config - The icon set conditional formatting rule settings.
   * @param {IIconSet['config']} config.iconConfigs - The icon configurations. iconId property is a string indexing of a group icons.
   * @param {boolean} config.isShowValue - Whether to show the value in the cell.
   * @returns {ConditionalFormatIconSetRuleBuilder} This builder for chaining.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a 3-arrow icon set conditional formatting rule in the range A1:D10.
   * // The first arrow is green for values greater than 20.
   * // The second arrow is yellow for values greater than 10.
   * // The third arrow is red for values less than or equal to 10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const builder = fWorksheet.newConditionalFormattingRule();
   * console.log(builder.getIconMap()); // icons key-value map
   * const rule = builder.setIconSet({
   *     iconConfigs: [
   *       {
   *         iconType: univerAPI.Enum.ConditionFormatIconSetTypeEnum.threeArrows,
   *         iconId: '0',
   *         operator: univerAPI.Enum.ConditionFormatNumberOperatorEnum.greaterThan,
   *         value: {
   *           type: univerAPI.Enum.ConditionFormatValueTypeEnum.num,
   *           value: 20
   *         }
   *       },
   *       {
   *         iconType: univerAPI.Enum.ConditionFormatIconSetTypeEnum.threeArrows,
   *         iconId: '1',
   *         operator: univerAPI.Enum.ConditionFormatNumberOperatorEnum.greaterThan,
   *         value: {
   *           type: univerAPI.Enum.ConditionFormatValueTypeEnum.num,
   *           value: 10
   *         }
   *       },
   *       {
   *         iconType: univerAPI.Enum.ConditionFormatIconSetTypeEnum.threeArrows,
   *         iconId: '2',
   *         operator: univerAPI.Enum.ConditionFormatNumberOperatorEnum.lessThanOrEqual,
   *         value: {
   *           type: univerAPI.Enum.ConditionFormatValueTypeEnum.num,
   *           value: 10
   *         }
   *       }
   *     ],
   *     isShowValue: true,
   *   })
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setIconSet(config) {
    const ruleConfig = this._ruleConfig;
    ruleConfig.type = "iconSet" /* iconSet */;
    ruleConfig.config = config.iconConfigs;
    ruleConfig.isShowValue = config.isShowValue;
    return this;
  }
};
var FConditionalFormattingBuilder = class {
  constructor(_initConfig = {}) {
    __publicField(this, "_initConfig", _initConfig);
  }
  /**
   * Constructs a conditional format rule from the settings applied to the builder.
   * @returns {IConditionFormattingRule} The conditional format rule.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with values greater than 10 in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenNumberGreaterThan(10)
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  build() {
    return new ConditionalFormatRuleBaseBuilder(this._initConfig).build();
  }
  /**
   * Set average rule.
   * @param {IAverageHighlightCell['operator']} operator - The operator to use for the average rule.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with greater than average values in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .setAverage(univerAPI.Enum.ConditionFormatNumberOperatorEnum.greaterThan)
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setAverage(operator) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).setAverage(operator);
  }
  /**
   * Set unique values rule.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with unique values in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .setUniqueValues()
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setUniqueValues() {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).setUniqueValues();
  }
  /**
   * Set duplicate values rule.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with duplicate values in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .setDuplicateValues()
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setDuplicateValues() {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).setDuplicateValues();
  }
  /**
   * Set rank rule.
   * @param {{ isBottom: boolean, isPercent: boolean, value: number }} config - The rank rule settings.
   * @param {boolean} config.isBottom - Whether to highlight the bottom rank.
   * @param {boolean} config.isPercent - Whether to use a percentage rank.
   * @param {number} config.value - The rank value.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights the bottom 10% of values in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .setRank({ isBottom: true, isPercent: true, value: 10 })
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setRank(config) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).setRank(config);
  }
  /**
   * Get the icon set mapping dictionary.
   * @returns {Record<string, string[]>} The icon set mapping dictionary.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * console.log(fWorksheet.newConditionalFormattingRule().getIconMap()); // icons key-value map
   * ```
   */
  getIconMap() {
    return iconMap;
  }
  /**
   * Set up icon set conditional formatting rule.
   * @param {{ iconConfigs: IIconSet['config'], isShowValue: boolean }} config - The icon set conditional formatting rule settings.
   * @param {IIconSet['config']} config.iconConfigs - The icon configurations. iconId property is a string indexing of a group icons.
   * @param {boolean} config.isShowValue - Whether to show the value in the cell.
   * @returns {ConditionalFormatIconSetRuleBuilder} The conditional format icon set rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a 3-arrow icon set conditional formatting rule in the range A1:D10.
   * // The first arrow is green for values greater than 20.
   * // The second arrow is yellow for values greater than 10.
   * // The third arrow is red for values less than or equal to 10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const builder = fWorksheet.newConditionalFormattingRule();
   * console.log(builder.getIconMap()); // icons key-value map
   * const rule = builder.setIconSet({
   *     iconConfigs: [
   *       {
   *         iconType: univerAPI.Enum.ConditionFormatIconSetTypeEnum.threeArrows,
   *         iconId: '0',
   *         operator: univerAPI.Enum.ConditionFormatNumberOperatorEnum.greaterThan,
   *         value: {
   *           type: univerAPI.Enum.ConditionFormatValueTypeEnum.num,
   *           value: 20
   *         }
   *       },
   *       {
   *         iconType: univerAPI.Enum.ConditionFormatIconSetTypeEnum.threeArrows,
   *         iconId: '1',
   *         operator: univerAPI.Enum.ConditionFormatNumberOperatorEnum.greaterThan,
   *         value: {
   *           type: univerAPI.Enum.ConditionFormatValueTypeEnum.num,
   *           value: 10
   *         }
   *       },
   *       {
   *         iconType: univerAPI.Enum.ConditionFormatIconSetTypeEnum.threeArrows,
   *         iconId: '2',
   *         operator: univerAPI.Enum.ConditionFormatNumberOperatorEnum.lessThanOrEqual,
   *         value: {
   *           type: univerAPI.Enum.ConditionFormatValueTypeEnum.num,
   *           value: 10
   *         }
   *       }
   *     ],
   *     isShowValue: true,
   *   })
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setIconSet(config) {
    return new ConditionalFormatIconSetRuleBuilder(this._initConfig).setIconSet(config);
  }
  /**
   * Set color scale rule.
   * @param {{ index: number; color: string; value: IValueConfig }[]} config - The color scale rule settings.
   * @param {number} config.index - The index of the color scale.
   * @param {string} config.color - The color for the color scale.
   * @param {IValueConfig} config.value - The value for the color scale.
   * @returns {ConditionalFormatColorScaleRuleBuilder} The conditional format color scale rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that adds a color scale to cells with values between 0 and 100 in the range A1:D10.
   * // The color scale is green for 0, yellow for 50, and red for 100.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .setColorScale([
   *     { index: 0, color: '#00FF00', value: { type: 'num', value: 0 } },
   *     { index: 1, color: '#FFFF00', value: { type: 'num', value: 50 } },
   *     { index: 2, color: '#FF0000', value: { type: 'num', value: 100 } }
   *   ])
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setColorScale(config) {
    return new ConditionalFormatColorScaleRuleBuilder(this._initConfig).setColorScale(config);
  }
  /**
   * Set data bar rule.
   * @param {{
   *         min: IValueConfig;
   *         max: IValueConfig;
   *         isGradient?: boolean;
   *         positiveColor: string;
   *         nativeColor: string;
   *         isShowValue?: boolean;
   *     }} config - The data bar rule settings.
   * @param {IValueConfig} config.min - The minimum value for the data bar.
   * @param {IValueConfig} config.max - The maximum value for the data bar.
   * @param {boolean} [config.isGradient] - Whether the data bar is gradient.
   * @param {string} config.positiveColor - The color for positive values.
   * @param {string} config.nativeColor - The color for negative values.
   * @param {boolean} [config.isShowValue] - Whether to show the value in the cell.
   * @returns {ConditionalFormatDataBarRuleBuilder} The conditional format data bar rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that adds a data bar to cells with values between -100 and 100 in the range A1:D10.
   * // positive values are green and negative values are red.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .setDataBar({
   *     min: { type: 'num', value: -100 },
   *     max: { type: 'num', value: 100 },
   *     positiveColor: '#00FF00',
   *     nativeColor: '#FF0000',
   *     isShowValue: true
   *   })
   *  .setRanges([fRange.getRange()])
   * .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setDataBar(config) {
    return new ConditionalFormatDataBarRuleBuilder(this._initConfig).setDataBar(config);
  }
  /**
   * Sets the background color for the conditional format rule's format.
   * @param {string} [color] - The background color to set. If not provided, the background color is removed.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with no content in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenCellEmpty()
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setBackground(color) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).setBackground(color);
  }
  /**
   * Sets text bolding for the conditional format rule's format.
   * @param {boolean} isBold - Whether to bold the text.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that bolds the text for cells with not empty content in the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenCellNotEmpty()
   *   .setBold(true)
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setBold(isBold) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).setBold(isBold);
  }
  /**
   * Sets the font color for the conditional format rule's format.
   * @param {string} [color] - The font color to set. If not provided, the font color is removed.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that changes the font color to red for cells with not empty content in the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenCellNotEmpty()
   *   .setFontColor('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setFontColor(color) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).setFontColor(color);
  }
  /**
   * Sets text italics for the conditional format rule's format.
   * @param {boolean} isItalic - Whether to italicize the text.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that italicizes the text for cells with not empty content in the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenCellNotEmpty()
   *   .setItalic(true)
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setItalic(isItalic) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).setItalic(isItalic);
  }
  /**
   * Sets text strikethrough for the conditional format rule's format.
   * @param {boolean} isStrikethrough - Whether is strikethrough the text.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that set text strikethrough for cells with not empty content in the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenCellNotEmpty()
   *   .setStrikethrough(true)
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setStrikethrough(isStrikethrough) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).setStrikethrough(isStrikethrough);
  }
  /**
   * Sets text underlining for the conditional format rule's format.
   * @param {boolean} isUnderline - Whether to underline the text.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that underlines the text for cells with not empty content in the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenCellNotEmpty()
   *   .setUnderline(true)
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  setUnderline(isUnderline) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).setUnderline(isUnderline);
  }
  /**
   * Sets the conditional format rule to trigger when the cell is empty.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with no content in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenCellEmpty()
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenCellEmpty() {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).whenCellEmpty();
  }
  /**
   * Sets the conditional format rule to trigger when the cell is not empty.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that changes the font color to red for cells with not empty content in the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenCellNotEmpty()
   *   .setFontColor('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenCellNotEmpty() {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).whenCellNotEmpty();
  }
  /**
   * Sets the conditional format rule to trigger when a time period is met.
   * @param {CFTimePeriodOperator} date - The time period to check for.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with dates in the last 7 days in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenDate(univerAPI.Enum.ConditionFormatTimePeriodOperatorEnum.last7Days)
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenDate(date) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).whenDate(date);
  }
  /**
   * Sets the conditional format rule to trigger when that the given formula evaluates to `true`.
   * @param {string} formulaString - A custom formula that evaluates to true if the input is valid. formulaString start with '='.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with values greater than 10 in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenFormulaSatisfied('=A1>10')
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenFormulaSatisfied(formulaString) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).whenFormulaSatisfied(formulaString);
  }
  /**
   * Sets the conditional format rule to trigger when a number falls between, or is either of, two specified values.
   * @param {number} start - The lowest acceptable value.
   * @param {number} end - The highest acceptable value.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with values between 10 and 20 in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenNumberBetween(10, 20)
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenNumberBetween(start, end) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).whenNumberBetween(start, end);
  }
  /**
   * Sets the conditional format rule to trigger when a number is equal to the given value.
   * @param {number} value - The sole acceptable value.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with values equal to 10 in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenNumberEqualTo(10)
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenNumberEqualTo(value) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).whenNumberEqualTo(value);
  }
  /**
   * Sets the conditional format rule to trigger when a number is greater than the given value.
   * @param {number} value - The highest unacceptable value.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with values greater than 10 in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenNumberGreaterThan(10)
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenNumberGreaterThan(value) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).whenNumberGreaterThan(value);
  }
  /**
   * Sets the conditional format rule to trigger when a number is greater than or equal to the given value.
   * @param {number} value - The lowest acceptable value.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with values greater than or equal to 10 in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenNumberGreaterThanOrEqualTo(10)
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenNumberGreaterThanOrEqualTo(value) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).whenNumberGreaterThanOrEqualTo(value);
  }
  /**
   * Sets the conditional conditional format rule to trigger when a number less than the given value.
   * @param {number} value - The lowest unacceptable value.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with values less than 10 in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenNumberLessThan(10)
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenNumberLessThan(value) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).whenNumberLessThan(value);
  }
  /**
   * Sets the conditional format rule to trigger when a number less than or equal to the given value.
   * @param {number} value - The highest acceptable value.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with values less than or equal to 10 in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenNumberLessThanOrEqualTo(10)
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenNumberLessThanOrEqualTo(value) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).whenNumberLessThanOrEqualTo(value);
  }
  /**
   * Sets the conditional format rule to trigger when a number does not fall between, and is neither of, two specified values.
   * @param {number} start - The lowest unacceptable value.
   * @param {number} end - The highest unacceptable value.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with values not between 10 and 20 in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenNumberNotBetween(10, 20)
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenNumberNotBetween(start, end) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).whenNumberNotBetween(start, end);
  }
  /**
   * Sets the conditional format rule to trigger when a number is not equal to the given value.
   * @param {number} value - The sole unacceptable value.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with values not equal to 10 in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenNumberNotEqualTo(10)
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenNumberNotEqualTo(value) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).whenNumberNotEqualTo(value);
  }
  /**
   * Sets the conditional format rule to trigger when that the input contains the given value.
   * @param {string} text - The value that the input must contain.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with text containing 'apple' in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenTextContains('apple')
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenTextContains(text) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).whenTextContains(text);
  }
  /**
   * Sets the conditional format rule to trigger when that the input does not contain the given value.
   * @param {string} text - The value that the input must not contain.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with text not containing 'apple' in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenTextDoesNotContain('apple')
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenTextDoesNotContain(text) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).whenTextDoesNotContain(text);
  }
  /**
   * Sets the conditional format rule to trigger when that the input ends with the given value.
   * @param {string} text - Text to compare against the end of the string.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with text ending with '.ai' in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenTextEndsWith('.ai')
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenTextEndsWith(text) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).whenTextEndsWith(text);
  }
  /**
   * Sets the conditional format rule to trigger when that the input is equal to the given value.
   * @param {string} text - The sole acceptable value.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with text equal to 'apple' in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenTextEqualTo('apple')
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenTextEqualTo(text) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).whenTextEqualTo(text);
  }
  /**
   * Sets the conditional format rule to trigger when that the input starts with the given value.
   * @param {string} text - Text to compare against the beginning of the string.
   * @returns {ConditionalFormatHighlightRuleBuilder} The conditional format highlight rule builder.
   * @example
   * ```typescript
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // Create a conditional formatting rule that highlights cells with text starting with 'https://' in red for the range A1:D10.
   * const fRange = fWorksheet.getRange('A1:D10');
   * const rule = fWorksheet.newConditionalFormattingRule()
   *   .whenTextStartsWith('https://')
   *   .setBackground('#FF0000')
   *   .setRanges([fRange.getRange()])
   *   .build();
   * fWorksheet.addConditionalFormattingRule(rule);
   * ```
   */
  whenTextStartsWith(text) {
    return new ConditionalFormatHighlightRuleBuilder(this._initConfig).whenTextStartsWith(text);
  }
};

// ../packages/sheets-conditional-formatting/src/facade/f-range.ts
var FRangeSheetsConditionalFormattingMixin = class extends FRange {
  _getConditionalFormattingRuleModel() {
    return this._injector.get(ConditionalFormattingRuleModel);
  }
  getConditionalFormattingRules() {
    const rules = this._getConditionalFormattingRuleModel().getSubunitRules(this._workbook.getUnitId(), this._worksheet.getSheetId()) || [];
    return [...rules].filter((rule) => rule.ranges.some((range) => Rectangle.intersects(range, this._range)));
  }
  createConditionalFormattingRule() {
    return new FConditionalFormattingBuilder({ ranges: [this._range] });
  }
  addConditionalFormattingRule(rule) {
    const params = {
      unitId: this._workbook.getUnitId(),
      subUnitId: this._worksheet.getSheetId(),
      rule
    };
    this._commandService.syncExecuteCommand(AddCfCommand.id, params);
    return this;
  }
  deleteConditionalFormattingRule(cfId) {
    const params = {
      unitId: this._workbook.getUnitId(),
      subUnitId: this._worksheet.getSheetId(),
      cfId
    };
    this._commandService.syncExecuteCommand(DeleteCfCommand.id, params);
    return this;
  }
  moveConditionalFormattingRule(cfId, toCfId, type = "after") {
    const params = {
      unitId: this._workbook.getUnitId(),
      subUnitId: this._worksheet.getSheetId(),
      start: { id: cfId, type: "self" },
      end: { id: toCfId, type }
    };
    this._commandService.syncExecuteCommand(MoveCfCommand.id, params);
    return this;
  }
  setConditionalFormattingRule(cfId, rule) {
    const params = {
      unitId: this._workbook.getUnitId(),
      subUnitId: this._worksheet.getSheetId(),
      rule,
      cfId
    };
    this._commandService.syncExecuteCommand(SetCfCommand.id, params);
    return this;
  }
  clearConditionalFormatRules() {
    const params = {
      unitId: this._workbook.getUnitId(),
      subUnitId: this._worksheet.getSheetId(),
      ranges: [this._range]
    };
    this._commandService.syncExecuteCommand(ClearRangeCfCommand.id, params);
    return this;
  }
};
FRange.extend(FRangeSheetsConditionalFormattingMixin);

// ../packages/sheets-conditional-formatting/src/facade/f-worksheet.ts
var FWorksheetConditionalFormattingMixin = class extends FWorksheet {
  _getConditionalFormattingRuleModel() {
    return this._injector.get(ConditionalFormattingRuleModel);
  }
  getConditionalFormattingRules() {
    const rules = this._getConditionalFormattingRuleModel().getSubunitRules(this._workbook.getUnitId(), this._worksheet.getSheetId()) || [];
    return [...rules];
  }
  createConditionalFormattingRule() {
    return new FConditionalFormattingBuilder();
  }
  newConditionalFormattingRule() {
    return new FConditionalFormattingBuilder();
  }
  addConditionalFormattingRule(rule) {
    const params = {
      unitId: this._workbook.getUnitId(),
      subUnitId: this._worksheet.getSheetId(),
      rule
    };
    this._commandService.syncExecuteCommand(AddCfCommand.id, params);
    return this;
  }
  deleteConditionalFormattingRule(cfId) {
    const params = {
      unitId: this._workbook.getUnitId(),
      subUnitId: this._worksheet.getSheetId(),
      cfId
    };
    this._commandService.syncExecuteCommand(DeleteCfCommand.id, params);
    return this;
  }
  moveConditionalFormattingRule(cfId, toCfId, type = "after") {
    const params = {
      unitId: this._workbook.getUnitId(),
      subUnitId: this._worksheet.getSheetId(),
      start: { id: cfId, type: "self" },
      end: { id: toCfId, type }
    };
    this._commandService.syncExecuteCommand(MoveCfCommand.id, params);
    return this;
  }
  setConditionalFormattingRule(cfId, rule) {
    const params = {
      unitId: this._workbook.getUnitId(),
      subUnitId: this._worksheet.getSheetId(),
      cfId,
      rule
    };
    this._commandService.syncExecuteCommand(SetCfCommand.id, params);
    return this;
  }
  clearConditionalFormatRules() {
    const params = {
      unitId: this._workbook.getUnitId(),
      subUnitId: this._worksheet.getSheetId()
    };
    this._commandService.syncExecuteCommand(ClearWorksheetCfCommand.id, params);
    return this;
  }
};
FWorksheet.extend(FWorksheetConditionalFormattingMixin);

// ../packages/sheets-conditional-formatting/src/facade/f-enum.ts
var FSheetsConditionalFormattingEnumMixin = class extends FEnum {
  get ConditionFormatNumberOperatorEnum() {
    return CFNumberOperator;
  }
  get ConditionFormatTimePeriodOperatorEnum() {
    return CFTimePeriodOperator;
  }
  get ConditionFormatIconSetTypeEnum() {
    return IIconSetType;
  }
  get ConditionFormatValueTypeEnum() {
    return CFValueType;
  }
};
FEnum.extend(FSheetsConditionalFormattingEnumMixin);

// ../packages/sheets-find-replace/src/facade/f-text-finder.ts
var FTextFinder = class extends Disposable {
  constructor(_initialState, _injector, _univerInstanceService, _findReplaceService) {
    super();
    __publicField(this, "_injector", _injector);
    __publicField(this, "_univerInstanceService", _univerInstanceService);
    __publicField(this, "_findReplaceService", _findReplaceService);
    __publicField(this, "_state", new FindReplaceState());
    __publicField(this, "_model");
    __publicField(this, "_complete");
    const providers = this._findReplaceService.getProviders();
    this._model = this._injector.createInstance(FindReplaceModel, this._state, providers);
    const newState = {
      ...createInitFindReplaceState(),
      ..._initialState
    };
    this._state.changeState(newState);
  }
  findAll() {
    if (!this._state.findCompleted || !this._complete) {
      return [];
    }
    return this._complete.results.map((result) => {
      return this._findMatchToFRange(result);
    });
  }
  findNext() {
    var _a;
    if (!this._state.findCompleted || !this._complete) {
      return null;
    }
    const match = (_a = this._model) == null ? void 0 : _a.moveToNextMatch();
    if (!match) {
      return null;
    }
    return this._findMatchToFRange(match);
  }
  findPrevious() {
    var _a;
    const match = (_a = this._model) == null ? void 0 : _a.moveToPreviousMatch();
    if (!match) {
      return null;
    }
    return this._findMatchToFRange(match);
  }
  getCurrentMatch() {
    var _a;
    if (!this._state.findCompleted || !this._complete) {
      throw new Error("Find operation is not completed.");
    }
    const match = (_a = this._model) == null ? void 0 : _a.currentMatch$.value;
    if (!match) {
      return null;
    }
    return this._findMatchToFRange(match);
  }
  async matchCaseAsync(matchCase) {
    this._state.changeState({ caseSensitive: matchCase, findCompleted: false });
    await firstValueFrom(this._state.stateUpdates$.pipe(filter((state) => state.findCompleted === true)));
    await this.ensureCompleteAsync();
    return this;
  }
  async matchEntireCellAsync(matchEntireCell) {
    this._state.changeState({ matchesTheWholeCell: matchEntireCell, findCompleted: false });
    await firstValueFrom(this._state.stateUpdates$.pipe(filter((state) => state.findCompleted === true)));
    await this.ensureCompleteAsync();
    return this;
  }
  async matchFormulaTextAsync(matchFormulaText) {
    this._state.changeState({ findBy: matchFormulaText ? "formula" /* FORMULA */ : "value" /* VALUE */, findCompleted: false });
    await firstValueFrom(this._state.stateUpdates$.pipe(filter((state) => state.findCompleted === true)));
    await this.ensureCompleteAsync();
    return this;
  }
  async replaceAllWithAsync(replaceText) {
    var _a, _b, _c;
    await this._state.changeState({ replaceRevealed: true, replaceString: replaceText });
    const res = (_c = (_b = await ((_a = this._model) == null ? void 0 : _a.replaceAll())) == null ? void 0 : _b.success) != null ? _c : 0;
    this._state.changeState({ replaceRevealed: false });
    return res;
  }
  async replaceWithAsync(replaceText) {
    var _a;
    await this._state.changeState({ replaceRevealed: true, replaceString: replaceText });
    await ((_a = this._model) == null ? void 0 : _a.replace());
    this._state.changeState({ replaceRevealed: false });
    return true;
  }
  async ensureCompleteAsync() {
    var _a;
    this._complete = await ((_a = this._model) == null ? void 0 : _a.start());
  }
  _findMatchToFRange(match) {
    const { unitId } = match;
    const { subUnitId, range } = match.range;
    const workbook = this._univerInstanceService.getUnit(unitId);
    const worksheet = workbook.getSheetBySheetId(subUnitId);
    return this._injector.createInstance(FRange, workbook, worksheet, range);
  }
};
FTextFinder = __decorateClass([
  __decorateParam(1, Inject(Injector)),
  __decorateParam(2, IUniverInstanceService),
  __decorateParam(3, IFindReplaceService)
], FTextFinder);

// ../packages/sheets-find-replace/src/facade/f-univer.ts
var FUniverSheetsFindReplaceMixin = class extends FUniver {
  async createTextFinderAsync(text) {
    const state = { findString: text };
    const textFinder = this._injector.createInstance(FTextFinder, state);
    await textFinder.ensureCompleteAsync();
    return textFinder;
  }
};
FUniver.extend(FUniverSheetsFindReplaceMixin);

// ../packages/sheets-drawing/src/facade/f-over-grid-image.ts
function convertSheetImageToFOverGridImage(sheetImage, skeleton) {
  const { from, to, flipY = false, flipX = false, angle = 0, skewX = 0, skewY = 0 } = sheetImage.sheetTransform;
  const { column: fromColumn, columnOffset: fromColumnOffset, row: fromRow, rowOffset: fromRowOffset } = from;
  const absolutePosition = convertPositionSheetOverGridToAbsolute(
    sheetImage.unitId,
    sheetImage.subUnitId,
    { from, to },
    skeleton
  );
  const { width, height } = absolutePosition;
  return {
    ...sheetImage,
    column: fromColumn,
    columnOffset: fromColumnOffset,
    row: fromRow,
    rowOffset: fromRowOffset,
    width,
    height,
    flipY,
    flipX,
    angle,
    skewX,
    skewY
  };
}
function convertFOverGridImageToSheetImage(fOverGridImage, sheetSkeletonService) {
  const skeleton = sheetSkeletonService.ensureSkeleton(fOverGridImage.unitId, fOverGridImage.subUnitId);
  if (!skeleton) {
    throw new Error(`Skeleton for unitId ${fOverGridImage.unitId} and subUnitId ${fOverGridImage.subUnitId} not found`);
  }
  const { column: fromColumn, columnOffset: fromColumnOffset, row: fromRow, rowOffset: fromRowOffset, flipY = false, flipX = false, angle = 0, skewX = 0, skewY = 0, width, height } = fOverGridImage;
  const absolutePosition = convertPositionCellToSheetOverGrid(
    fOverGridImage.unitId,
    fOverGridImage.subUnitId,
    { column: fromColumn, columnOffset: fromColumnOffset, row: fromRow, rowOffset: fromRowOffset },
    width,
    height,
    skeleton
  );
  const { sheetTransform, transform } = absolutePosition;
  return {
    ...fOverGridImage,
    sheetTransform: {
      ...sheetTransform,
      flipY,
      flipX,
      angle,
      skewX,
      skewY
    },
    transform: {
      ...transform,
      flipY,
      flipX,
      angle,
      skewX,
      skewY
    },
    axisAlignSheetTransform: transformToAxisAlignPosition(transform, skeleton)
  };
}
var FOverGridImageBuilder = class {
  constructor(unitId, subUnitId, _injector) {
    __publicField(this, "_injector", _injector);
    __publicField(this, "_image");
    this._image = {
      drawingId: generateRandomId(6),
      drawingType: 0 /* DRAWING_IMAGE */,
      imageSourceType: "BASE64" /* BASE64 */,
      source: "",
      unitId,
      subUnitId,
      column: 0,
      columnOffset: 0,
      row: 0,
      rowOffset: 0,
      width: 0,
      height: 0,
      axisAlignSheetTransform: {
        from: {
          column: 0,
          columnOffset: 0,
          row: 0,
          rowOffset: 0
        },
        to: {
          column: 0,
          columnOffset: 0,
          row: 0,
          rowOffset: 0
        }
      }
    };
  }
  /**
   * Set the initial image configuration for the image builder.
   * @param {ISheetImage} image - The image configuration
   * @returns {FOverGridImageBuilder} The `FOverGridImageBuilder` for chaining
   * @example
   * ```ts
   * // create a new image builder and set initial image configuration.
   * // then build `ISheetImage` and insert it into the sheet, position is start from F6 cell.
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const image = await fWorksheet.newOverGridImage()
   *   .setImage({
   *     drawingId: '123456',
   *     drawingType: univerAPI.Enum.DrawingType.DRAWING_IMAGE,
   *     imageSourceType: univerAPI.Enum.ImageSourceType.BASE64,
   *     source: 'https://avatars.githubusercontent.com/u/61444807?s=48&v=4',
   *     unitId: fWorkbook.getId(),
   *     subUnitId: fWorksheet.getSheetId(),
   *   })
   *   .setColumn(5)
   *   .setRow(5)
   *   .buildAsync();
   * fWorksheet.insertImages([image]);
   * ```
   */
  setImage(image) {
    const { unitId, subUnitId } = image;
    const sheetSkeletonService = this._injector.get(SheetSkeletonService);
    const skeleton = sheetSkeletonService.getSkeleton(unitId, subUnitId);
    if (!skeleton) {
      throw new Error(`Skeleton for unitId ${unitId} and subUnitId ${subUnitId} not found`);
    }
    if (image.sheetTransform == null) {
      image.sheetTransform = {
        from: {
          column: 0,
          columnOffset: 0,
          row: 0,
          rowOffset: 0
        },
        to: {
          column: 0,
          columnOffset: 0,
          row: 0,
          rowOffset: 0
        }
      };
    }
    if (image.axisAlignSheetTransform == null) {
      image.axisAlignSheetTransform = {
        from: {
          column: 0,
          columnOffset: 0,
          row: 0,
          rowOffset: 0
        },
        to: {
          column: 0,
          columnOffset: 0,
          row: 0,
          rowOffset: 0
        }
      };
    }
    this._image = convertSheetImageToFOverGridImage(image, skeleton);
    return this;
  }
  setSource(source, sourceType) {
    const sourceTypeVal = sourceType != null ? sourceType : "URL" /* URL */;
    this._image.source = source;
    this._image.imageSourceType = sourceTypeVal;
    return this;
  }
  /**
   * Get the source of the image
   * @returns {string} The source of the image
   * @example
   * ```ts
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const images = fWorksheet.getImages();
   * images.forEach((image) => {
   *   console.log(image, image.toBuilder().getSource());
   * });
   * ```
   */
  getSource() {
    return this._image.source;
  }
  /**
   * Get the source type of the image
   * @returns {ImageSourceType} The source type of the image
   * @example
   * ```ts
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const images = fWorksheet.getImages();
   * images.forEach((image) => {
   *   console.log(image, image.toBuilder().getSourceType());
   * });
   * ```
   */
  getSourceType() {
    return this._image.imageSourceType;
  }
  /**
   * Set the horizontal position of the image
   * @param {number} column - The column index of the image start position, start at 0
   * @returns {FOverGridImageBuilder} The `FOverGridImageBuilder` for chaining
   * @example
   * ```ts
   * // create a new image builder and set image source.
   * // then build `ISheetImage` and insert it into the sheet, position is start from F6 cell.
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const image = await fWorksheet.newOverGridImage()
   *   .setSource('https://avatars.githubusercontent.com/u/61444807?s=48&v=4', univerAPI.Enum.ImageSourceType.URL)
   *   .setColumn(5)
   *   .setRow(5)
   *   .buildAsync();
   * fWorksheet.insertImages([image]);
   * ```
   */
  setColumn(column) {
    this._image.column = column;
    return this;
  }
  /**
   * Set the vertical position of the image
   * @param {number} row - The row index of the image start position, start at 0
   * @returns {FOverGridImageBuilder} The `FOverGridImageBuilder` for chaining
   * @example
   * ```ts
   * // create a new image builder and set image source.
   * // then build `ISheetImage` and insert it into the sheet, position is start from F6 cell.
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const image = await fWorksheet.newOverGridImage()
   *   .setSource('https://avatars.githubusercontent.com/u/61444807?s=48&v=4', univerAPI.Enum.ImageSourceType.URL)
   *   .setColumn(5)
   *   .setRow(5)
   *   .buildAsync();
   * fWorksheet.insertImages([image]);
   * ```
   */
  setRow(row) {
    this._image.row = row;
    return this;
  }
  /**
   * Set the horizontal offset of the image
   * @param {number} offset - The column offset of the image start position, pixel unit
   * @returns {FOverGridImageBuilder} The `FOverGridImageBuilder` for chaining
   * @example
   * ```ts
   * // create a new image builder and set image source.
   * // then build `ISheetImage` and insert it into the sheet, position is start from F6 cell and horizontal offset is 10px.
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const image = await fWorksheet.newOverGridImage()
   *   .setSource('https://avatars.githubusercontent.com/u/61444807?s=48&v=4', univerAPI.Enum.ImageSourceType.URL)
   *   .setColumn(5)
   *   .setRow(5)
   *   .setColumnOffset(10)
   *   .buildAsync();
   * fWorksheet.insertImages([image]);
   * ```
   */
  setColumnOffset(offset) {
    this._image.columnOffset = offset;
    return this;
  }
  /**
   * Set the vertical offset of the image
   * @param {number} offset - The row offset of the image start position, pixel unit
   * @returns {FOverGridImageBuilder} The `FOverGridImageBuilder` for chaining
   * @example
   * ```ts
   * // create a new image builder and set image source.
   * // then build `ISheetImage` and insert it into the sheet, position is start from F6 cell and vertical offset is 10px.
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const image = await fWorksheet.newOverGridImage()
   *   .setSource('https://avatars.githubusercontent.com/u/61444807?s=48&v=4', univerAPI.Enum.ImageSourceType.URL)
   *   .setColumn(5)
   *   .setRow(5)
   *   .setRowOffset(10)
   *   .buildAsync();
   * fWorksheet.insertImages([image]);
   * ```
   */
  setRowOffset(offset) {
    this._image.rowOffset = offset;
    return this;
  }
  /**
   * Set the width of the image
   * @param {number} width - The width of the image, pixel unit
   * @returns {FOverGridImageBuilder} The `FOverGridImageBuilder` for chaining
   * @example
   * ```ts
   * // create a new image builder and set image source.
   * // then build `ISheetImage` and insert it into the sheet, position is start from F6 cell, width is 120px and height is 50px.
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const image = await fWorksheet.newOverGridImage()
   *   .setSource('https://avatars.githubusercontent.com/u/61444807?s=48&v=4', univerAPI.Enum.ImageSourceType.URL)
   *   .setColumn(5)
   *   .setRow(5)
   *   .setWidth(120)
   *   .setHeight(50)
   *   .buildAsync();
   * fWorksheet.insertImages([image]);
   * ```
   */
  setWidth(width) {
    this._image.width = width;
    return this;
  }
  /**
   * Set the height of the image
   * @param {number} height - The height of the image, pixel unit
   * @returns {FOverGridImageBuilder} The `FOverGridImageBuilder` for chaining
   * @example
   * ```ts
   * // create a new image builder and set image source.
   * // then build `ISheetImage` and insert it into the sheet, position is start from F6 cell, width is 120px and height is 50px.
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const image = await fWorksheet.newOverGridImage()
   *   .setSource('https://avatars.githubusercontent.com/u/61444807?s=48&v=4', univerAPI.Enum.ImageSourceType.URL)
   *   .setColumn(5)
   *   .setRow(5)
   *   .setWidth(120)
   *   .setHeight(50)
   *   .buildAsync();
   * fWorksheet.insertImages([image]);
   * ```
   */
  setHeight(height) {
    this._image.height = height;
    return this;
  }
  /**
   * Set the anchor type of the image, whether the position and size change with the cell
   * @param {SheetDrawingAnchorType} anchorType - The anchor type of the image
   * @returns {FOverGridImageBuilder} The `FOverGridImageBuilder` for chaining
   * @example
   * ```ts
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   *
   * // image1 position is start from A6 cell, anchor type is Position.
   * // Only the position of the drawing follows the cell changes. When rows or columns are inserted or deleted, the position of the drawing changes, but the size remains the same.
   * const image1 = await fWorksheet.newOverGridImage()
   *   .setSource('https://avatars.githubusercontent.com/u/61444807?s=48&v=4', univerAPI.Enum.ImageSourceType.URL)
   *   .setColumn(0)
   *   .setRow(5)
   *   .setAnchorType(univerAPI.Enum.SheetDrawingAnchorType.Position)
   *   .buildAsync();
   *
   * // image2 position is start from C6 cell, anchor type is Both.
   * // The size and position of the drawing follow the cell changes. When rows or columns are inserted or deleted, the size and position of the drawing change accordingly.
   * const image2 = await fWorksheet.newOverGridImage()
   *   .setSource('https://avatars.githubusercontent.com/u/61444807?s=48&v=4', univerAPI.Enum.ImageSourceType.URL)
   *   .setColumn(2)
   *   .setRow(5)
   *   .setAnchorType(univerAPI.Enum.SheetDrawingAnchorType.Both)
   *   .buildAsync();
   *
   * // image3 position is start from E6 cell, anchor type is None.
   * // The size and position of the drawing do not follow the cell changes. When rows or columns are inserted or deleted, the position and size of the drawing remain unchanged.
   * const image3 = await fWorksheet.newOverGridImage()
   *   .setSource('https://avatars.githubusercontent.com/u/61444807?s=48&v=4', univerAPI.Enum.ImageSourceType.URL)
   *   .setColumn(4)
   *   .setRow(5)
   *   .setAnchorType(univerAPI.Enum.SheetDrawingAnchorType.None)
   *   .buildAsync();
   *
   * // insert images into the sheet
   * fWorksheet.insertImages([image1, image2, image3]);
   *
   * // after 2 seconds, set the row height of the 5th row to 100px and insert a row before the 5th row.
   * // then observe the position and size changes of the images.
   * setTimeout(() => {
   *   fWorksheet.setRowHeight(5, 100).insertRowBefore(5);
   * }, 2000);
   * ```
   */
  setAnchorType(anchorType) {
    this._image.anchorType = anchorType;
    return this;
  }
  /**
   * Set the cropping region of the image by defining the top edges, thereby displaying the specific part of the image you want.
   * @param {number} top - The number of pixels to crop from the top of the image
   * @returns {FOverGridImageBuilder} The `FOverGridImageBuilder` for chaining
   * @example
   * ```ts
   * // create a new image builder and set image source.
   * // then build `ISheetImage` and insert it into the sheet, position is start from F6 cell, top crop is 10px.
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const image = await fWorksheet.newOverGridImage()
   *   .setSource('https://avatars.githubusercontent.com/u/61444807?s=48&v=4', univerAPI.Enum.ImageSourceType.URL)
   *   .setColumn(5)
   *   .setRow(5)
   *   .setCropTop(10)
   *   .buildAsync();
   * fWorksheet.insertImages([image]);
   * ```
   */
  setCropTop(top) {
    this._initializeSrcRect();
    this._image.srcRect.top = top;
    return this;
  }
  /**
   * Set the cropping region of the image by defining the left edges, thereby displaying the specific part of the image you want.
   * @param {number} left - The number of pixels to crop from the left side of the image
   * @returns {FOverGridImageBuilder} The `FOverGridImageBuilder` for chaining
   * @example
   * ```ts
   * // create a new image builder and set image source.
   * // then build `ISheetImage` and insert it into the sheet, position is start from F6 cell, left crop is 10px.
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const image = await fWorksheet.newOverGridImage()
   *   .setSource('https://avatars.githubusercontent.com/u/61444807?s=48&v=4', univerAPI.Enum.ImageSourceType.URL)
   *   .setColumn(5)
   *   .setRow(5)
   *   .setCropLeft(10)
   *   .buildAsync();
   * fWorksheet.insertImages([image]);
   * ```
   */
  setCropLeft(left) {
    this._initializeSrcRect();
    this._image.srcRect.left = left;
    return this;
  }
  /**
   * Set the cropping region of the image by defining the bottom edges, thereby displaying the specific part of the image you want.
   * @param {number} bottom - The number of pixels to crop from the bottom of the image
   * @returns {FOverGridImageBuilder} The `FOverGridImageBuilder` for chaining
   * @example
   * ```ts
   * // create a new image builder and set image source.
   * // then build `ISheetImage` and insert it into the sheet, position is start from F6 cell, bottom crop is 10px.
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const image = await fWorksheet.newOverGridImage()
   *   .setSource('https://avatars.githubusercontent.com/u/61444807?s=48&v=4', univerAPI.Enum.ImageSourceType.URL)
   *   .setColumn(5)
   *   .setRow(5)
   *   .setCropBottom(10)
   *   .buildAsync();
   * fWorksheet.insertImages([image]);
   * ```
   */
  setCropBottom(bottom) {
    this._initializeSrcRect();
    this._image.srcRect.bottom = bottom;
    return this;
  }
  /**
   * Set the cropping region of the image by defining the right edges, thereby displaying the specific part of the image you want.
   * @param {number} right - The number of pixels to crop from the right side of the image
   * @returns {FOverGridImageBuilder} The `FOverGridImageBuilder` for chaining
   * @example
   * ```ts
   * // create a new image builder and set image source.
   * // then build `ISheetImage` and insert it into the sheet, position is start from F6 cell, right crop is 10px.
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const image = await fWorksheet.newOverGridImage()
   *   .setSource('https://avatars.githubusercontent.com/u/61444807?s=48&v=4', univerAPI.Enum.ImageSourceType.URL)
   *   .setColumn(5)
   *   .setRow(5)
   *   .setCropRight(10)
   *   .buildAsync();
   * fWorksheet.insertImages([image]);
   * ```
   */
  setCropRight(right) {
    this._initializeSrcRect();
    this._image.srcRect.right = right;
    return this;
  }
  _initializeSrcRect() {
    if (this._image.srcRect == null) {
      this._image.srcRect = {
        top: 0,
        left: 0,
        bottom: 0,
        right: 0
      };
    }
  }
  /**
   * Set the rotation angle of the image
   * @param {number} angle - Degree of rotation of the image, for example, 90, 180, 270, etc.
   * @returns {FOverGridImageBuilder} The `FOverGridImageBuilder` for chaining
   * @example
   * ```ts
   * // create a new image builder and set image source.
   * // then build `ISheetImage` and insert it into the sheet, position is start from F6 cell, rotate 90 degrees.
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const image = await fWorksheet.newOverGridImage()
   *   .setSource('https://avatars.githubusercontent.com/u/61444807?s=48&v=4', univerAPI.Enum.ImageSourceType.URL)
   *   .setColumn(5)
   *   .setRow(5)
   *   .setRotate(90)
   *   .buildAsync();
   * fWorksheet.insertImages([image]);
   * ```
   */
  setRotate(angle) {
    this._image.angle = angle;
    return this;
  }
  setUnitId(unitId) {
    this._image.unitId = unitId;
    return this;
  }
  setSubUnitId(subUnitId) {
    this._image.subUnitId = subUnitId;
    return this;
  }
  async buildAsync() {
    const sheetSkeletonService = this._injector.get(SheetSkeletonService);
    if (this._image.width === 0 || this._image.height === 0) {
      const size = await getImageSize(this._image.source);
      const width = size.width;
      const height = size.height;
      if (this._image.width === 0) {
        this._image.width = width;
      }
      if (this._image.height === 0) {
        this._image.height = height;
      }
    }
    return convertFOverGridImageToSheetImage(this._image, sheetSkeletonService);
  }
};
FOverGridImageBuilder = __decorateClass([
  __decorateParam(2, Inject(Injector))
], FOverGridImageBuilder);
var FOverGridImage = class extends FBase {
  constructor(_image, _commandService, _injector) {
    super();
    __publicField(this, "_image", _image);
    __publicField(this, "_commandService", _commandService);
    __publicField(this, "_injector", _injector);
  }
  /**
   * Get the id of the image
   * @returns {string} The id of the image
   * @example
   * ```ts
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const images = fWorksheet.getImages();
   * images.forEach((image) => {
   *   console.log(image, image.getId());
   * });
   * ```
   */
  getId() {
    return this._image.drawingId;
  }
  /**
   * Get the drawing type of the image
   * @returns {DrawingTypeEnum} The drawing type of the image
   * @example
   * ```ts
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const images = fWorksheet.getImages();
   * images.forEach((image) => {
   *   console.log(image, image.getType());
   * });
   * ```
   */
  getType() {
    return this._image.drawingType;
  }
  /**
   * Remove the image from the sheet
   * @returns {boolean} true if the image is removed successfully, otherwise false
   * @example
   * ```ts
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const image = fWorksheet.getImages()[0];
   * const result = image?.remove();
   * console.log(result);
   * ```
   */
  remove() {
    return this._commandService.syncExecuteCommand(RemoveSheetDrawingCommand.id, { unitId: this._image.unitId, drawings: [this._image] });
  }
  /**
   * Convert the image to a FOverGridImageBuilder
   * @returns {FOverGridImageBuilder} The builder FOverGridImageBuilder
   * @example
   * ```ts
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const images = fWorksheet.getImages();
   * images.forEach((image) => {
   *   console.log(image, image.toBuilder().getSource());
   * });
   * ```
   */
  toBuilder() {
    const builder = this._injector.createInstance(FOverGridImageBuilder, this._image.unitId, this._image.subUnitId);
    builder.setImage(this._image);
    return builder;
  }
  setSource(source, sourceType) {
    const sourceTypeVal = sourceType != null ? sourceType : "URL" /* URL */;
    this._image.source = source;
    this._image.imageSourceType = sourceTypeVal;
    return this._commandService.syncExecuteCommand(SetSheetDrawingCommand.id, { unitId: this._image.unitId, drawings: [this._image] });
  }
  async setPositionAsync(row, column, rowOffset, columnOffset) {
    const builder = this.toBuilder();
    builder.setColumn(column);
    builder.setRow(row);
    if (rowOffset != null) {
      builder.setRowOffset(rowOffset);
    }
    if (columnOffset != null) {
      builder.setColumnOffset(columnOffset);
    }
    const param = await builder.buildAsync();
    return this._commandService.syncExecuteCommand(SetSheetDrawingCommand.id, { unitId: this._image.unitId, drawings: [param] });
  }
  /**
   * Set the size of the image
   * @param {number} width - The width of the image, pixel unit
   * @param {number} height - The height of the image, pixel unit
   * @returns {boolean} true if the size is set successfully, otherwise false
   * @example
   * ```ts
   * // set the image width 120px and height 50px
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const image = fWorksheet.getImages()[0];
   * const result = image?.setSizeAsync(120, 50);
   * console.log(result);
   * ```
   */
  async setSizeAsync(width, height) {
    const builder = this.toBuilder();
    builder.setWidth(width);
    builder.setHeight(height);
    const param = await builder.buildAsync();
    return this._commandService.syncExecuteCommand(SetSheetDrawingCommand.id, { unitId: this._image.unitId, drawings: [param] });
  }
  /**
   * Set the cropping region of the image by defining the top, bottom, left, and right edges, thereby displaying the specific part of the image you want.
   * @param {number} top - The number of pixels to crop from the top of the image
   * @param {number} left - The number of pixels to crop from the left side of the image
   * @param {number} bottom - The number of pixels to crop from the bottom of the image
   * @param {number} right - The number of pixels to crop from the right side of the image
   * @returns {boolean} true if the crop is set successfully, otherwise false
   * @example
   * ```ts
   * // set the crop of the image, top 10px, left 10px, bottom 10px, right 10px.
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const image = fWorksheet.getImages()[0];
   * const result = image?.setCrop(10, 10, 10, 10);
   * console.log(result);
   * ```
   */
  setCrop(top, left, bottom, right) {
    if (this._image.srcRect == null) {
      this._image.srcRect = {
        top: 0,
        left: 0,
        bottom: 0,
        right: 0
      };
    }
    if (top != null) {
      this._image.srcRect.top = top;
    }
    if (left != null) {
      this._image.srcRect.left = left;
    }
    if (bottom != null) {
      this._image.srcRect.bottom = bottom;
    }
    if (right != null) {
      this._image.srcRect.right = right;
    }
    return this._commandService.syncExecuteCommand(SetSheetDrawingCommand.id, { unitId: this._image.unitId, drawings: [this._image] });
  }
  /**
   * Set the rotation angle of the image
   * @param {number} angle - Degree of rotation of the image, for example, 90, 180, 270, etc.
   * @returns {boolean} true if the rotation is set successfully, otherwise false
   * @example
   * ```ts
   * // set 90 degrees rotation of the image
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const image = fWorksheet.getImages()[0];
   * const result = image?.setRotate(90);
   * console.log(result);
   * ```
   */
  setRotate(angle) {
    this._image.sheetTransform.angle = angle;
    this._image.transform && (this._image.transform.angle = angle);
    if (this._image.transform) {
      const sheetSkeletonService = this._injector.get(SheetSkeletonService);
      const skeleton = sheetSkeletonService.getSkeleton(this._image.unitId, this._image.subUnitId);
      if (!skeleton) {
        throw new Error(`Skeleton for unitId ${this._image.unitId} and subUnitId ${this._image.subUnitId} not found`);
      }
      this._image.axisAlignSheetTransform && (this._image.axisAlignSheetTransform = transformToAxisAlignPosition(this._image.transform, skeleton));
    }
    return this._commandService.syncExecuteCommand(SetSheetDrawingCommand.id, { unitId: this._image.unitId, drawings: [this._image] });
  }
  /**
   * Move the image layer forward by one level
   * @returns {boolean} true if the image is moved forward successfully, otherwise false
   * @example
   * ```ts
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const image = fWorksheet.getImages()[0];
   * const result = image?.setForward();
   * console.log(result);
   * ```
   */
  setForward() {
    return this._commandService.syncExecuteCommand(SetDrawingArrangeCommand.id, {
      unitId: this._image.unitId,
      subUnitId: this._image.subUnitId,
      drawingIds: [this._image.drawingId],
      arrangeType: 0 /* forward */
    });
  }
  /**
   * Move the image layer backward by one level
   * @returns {boolean} true if the image is moved backward successfully, otherwise false
   * @example
   * ```ts
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const image = fWorksheet.getImages()[0];
   * const result = image?.setBackward();
   * console.log(result);
   * ```
   */
  setBackward() {
    return this._commandService.syncExecuteCommand(SetDrawingArrangeCommand.id, {
      unitId: this._image.unitId,
      subUnitId: this._image.subUnitId,
      drawingIds: [this._image.drawingId],
      arrangeType: 1 /* backward */
    });
  }
  /**
   * Move the image layer to the bottom layer
   * @returns {boolean} true if the image is moved to the bottom layer successfully, otherwise false
   * @example
   * ```ts
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const image = fWorksheet.getImages()[0];
   * const result = image?.setBack();
   * console.log(result);
   * ```
   */
  setBack() {
    return this._commandService.syncExecuteCommand(SetDrawingArrangeCommand.id, {
      unitId: this._image.unitId,
      subUnitId: this._image.subUnitId,
      drawingIds: [this._image.drawingId],
      arrangeType: 3 /* back */
    });
  }
  /**
   * Move the image layer to the top layer
   * @returns {boolean} true if the image is moved to the top layer successfully, otherwise false
   * @example
   * ```ts
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const image = fWorksheet.getImages()[0];
   * const result = image?.setFront();
   * console.log(result);
   * ```
   */
  setFront() {
    return this._commandService.syncExecuteCommand(SetDrawingArrangeCommand.id, {
      unitId: this._image.unitId,
      subUnitId: this._image.subUnitId,
      drawingIds: [this._image.drawingId],
      arrangeType: 2 /* front */
    });
  }
};
FOverGridImage = __decorateClass([
  __decorateParam(1, ICommandService),
  __decorateParam(2, Inject(Injector))
], FOverGridImage);

// ../packages/sheets-drawing/src/facade/f-worksheet.ts
var GROUPABLE_DRAWING_TYPES = /* @__PURE__ */ new Set([
  0 /* DRAWING_IMAGE */,
  1 /* DRAWING_SHAPE */,
  6 /* DRAWING_GROUP */
]);
var FWorksheetDrawingMixin = class extends FWorksheet {
  async insertImage(url, column, row, offsetX, offsetY) {
    const imageBuilder = this.newOverGridImage();
    if (typeof url === "string") {
      imageBuilder.setSource(url);
    } else {
      const blobSource = url.getBlob();
      const base64 = await blobSource.getDataAsString();
      imageBuilder.setSource(base64, "BASE64" /* BASE64 */);
    }
    if (column !== void 0) {
      imageBuilder.setColumn(column);
    } else {
      imageBuilder.setColumn(0);
    }
    if (row !== void 0) {
      imageBuilder.setRow(row);
    } else {
      imageBuilder.setRow(0);
    }
    if (offsetX !== void 0) {
      imageBuilder.setColumnOffset(offsetX);
    } else {
      imageBuilder.setColumnOffset(0);
    }
    if (offsetY !== void 0) {
      imageBuilder.setRowOffset(offsetY);
    } else {
      imageBuilder.setRowOffset(0);
    }
    const param = await imageBuilder.buildAsync();
    return this._commandService.syncExecuteCommand(InsertSheetDrawingCommand.id, { unitId: this._fWorkbook.getId(), drawings: [param] });
  }
  insertImages(sheetImages) {
    const param = sheetImages.map((image) => {
      image.unitId = this._fWorkbook.getId();
      image.subUnitId = this.getSheetId();
      return image;
    });
    this._commandService.syncExecuteCommand(InsertSheetDrawingCommand.id, { unitId: this._fWorkbook.getId(), drawings: param });
    return this;
  }
  deleteImages(sheetImages) {
    const drawings = sheetImages.map((image) => {
      return {
        unitId: this._fWorkbook.getId(),
        drawingId: image.getId(),
        subUnitId: this.getSheetId(),
        drawingType: image.getType()
      };
    });
    this._commandService.syncExecuteCommand(RemoveSheetDrawingCommand.id, { unitId: this._fWorkbook.getId(), drawings });
    return this;
  }
  getImages() {
    const sheetDrawingService = this._injector.get(ISheetDrawingService);
    const drawingData = sheetDrawingService.getDrawingData(this._fWorkbook.getId(), this.getSheetId());
    const images = [];
    for (const drawingId in drawingData) {
      const drawing = drawingData[drawingId];
      if (drawing.drawingType !== 0 /* DRAWING_IMAGE */) {
        continue;
      }
      images.push(this._injector.createInstance(FOverGridImage, drawing));
    }
    return images;
  }
  getImageById(id) {
    const sheetDrawingService = this._injector.get(ISheetDrawingService);
    const drawing = sheetDrawingService.getDrawingByParam({ unitId: this._fWorkbook.getId(), subUnitId: this.getSheetId(), drawingId: id });
    if (drawing && drawing.drawingType === 0 /* DRAWING_IMAGE */) {
      return this._injector.createInstance(FOverGridImage, drawing);
    }
    return null;
  }
  getActiveImages() {
    const sheetDrawingService = this._injector.get(ISheetDrawingService);
    const drawingData = sheetDrawingService.getFocusDrawings();
    const images = [];
    for (const drawingId in drawingData) {
      const drawing = drawingData[drawingId];
      images.push(this._injector.createInstance(FOverGridImage, drawing));
    }
    return images;
  }
  updateImages(sheetImages) {
    this._commandService.syncExecuteCommand(SetSheetDrawingCommand.id, { unitId: this._fWorkbook.getId(), drawings: sheetImages });
    return this;
  }
  newOverGridImage() {
    const unitId = this._fWorkbook.getId();
    const subUnitId = this.getSheetId();
    return this._injector.createInstance(FOverGridImageBuilder, unitId, subUnitId);
  }
  groupDrawings(drawingIds, groupId = generateRandomId(10)) {
    const uniqueDrawingIds = Array.from(new Set(drawingIds));
    if (uniqueDrawingIds.length < 2) return null;
    const unitId = this._fWorkbook.getId();
    const subUnitId = this.getSheetId();
    const sheetDrawingService = this._injector.get(ISheetDrawingService);
    if (sheetDrawingService.getDrawingByParam({ unitId, subUnitId, drawingId: groupId })) return null;
    const drawings = uniqueDrawingIds.map((drawingId) => sheetDrawingService.getDrawingByParam({ unitId, subUnitId, drawingId }));
    if (drawings.some((drawing) => !drawing)) return null;
    if (drawings.some((drawing) => !GROUPABLE_DRAWING_TYPES.has(drawing.drawingType))) return null;
    const validDrawings = drawings;
    const groupTransform = getGroupState(0, 0, validDrawings.map((drawing) => drawing.transform || {}));
    const groupParam = {
      unitId,
      subUnitId,
      drawingId: groupId,
      drawingType: 6 /* DRAWING_GROUP */,
      transform: groupTransform,
      groupBaseBound: {
        left: groupTransform.left,
        top: groupTransform.top,
        width: groupTransform.width,
        height: groupTransform.height
      }
    };
    const children = validDrawings.map((drawing) => ({
      ...drawing,
      groupId
    }));
    const result = this._applyGroupDrawingOperation([{
      parent: groupParam,
      children
    }], 4 /* GROUP */);
    return result ? groupId : null;
  }
  ungroupDrawings(groupIds) {
    const unitId = this._fWorkbook.getId();
    const subUnitId = this.getSheetId();
    const sheetDrawingService = this._injector.get(ISheetDrawingService);
    const groupParams = [];
    for (const groupId of groupIds) {
      const groupDrawing = sheetDrawingService.getDrawingByParam({ unitId, subUnitId, drawingId: groupId });
      if (!groupDrawing || groupDrawing.drawingType !== 6 /* DRAWING_GROUP */) continue;
      const groupTransform = groupDrawing.transform || { width: 0, height: 0 };
      const children = sheetDrawingService.getDrawingsByGroup({ unitId, subUnitId, drawingId: groupId }).map((drawing) => {
        const newTransform = transformObjectOutOfGroup(
          drawing.transform || {},
          groupTransform,
          groupTransform.width || 0,
          groupTransform.height || 0,
          groupDrawing.groupBaseBound
        );
        return {
          ...drawing,
          transform: {
            ...drawing.transform,
            ...newTransform
          },
          groupId: void 0
        };
      });
      if (children.length > 0) {
        groupParams.push({
          parent: groupDrawing,
          children
        });
      }
    }
    return this._applyGroupDrawingOperation(groupParams, 5 /* UNGROUP */);
  }
  getDrawingGroupChildren(groupId, recursive = false) {
    const unitId = this._fWorkbook.getId();
    const subUnitId = this.getSheetId();
    const sheetDrawingService = this._injector.get(ISheetDrawingService);
    if (!recursive) {
      return sheetDrawingService.getDrawingsByGroup({ unitId, subUnitId, drawingId: groupId });
    }
    const nested = sheetDrawingService.getDrawingsByGroupNested({ unitId, subUnitId, drawingId: groupId });
    if (!nested) return [];
    return [
      ...nested.flatChildren || [],
      ...nested.groups.filter((group) => group.drawingId !== groupId)
    ];
  }
  getDrawingParentGroup(drawingId) {
    const unitId = this._fWorkbook.getId();
    const subUnitId = this.getSheetId();
    const sheetDrawingService = this._injector.get(ISheetDrawingService);
    const drawing = sheetDrawingService.getDrawingByParam({ unitId, subUnitId, drawingId });
    if (!(drawing == null ? void 0 : drawing.groupId)) return null;
    const groupDrawing = sheetDrawingService.getDrawingByParam({ unitId, subUnitId, drawingId: drawing.groupId });
    if (!groupDrawing || groupDrawing.drawingType !== 6 /* DRAWING_GROUP */) return null;
    return groupDrawing;
  }
  isDrawingGrouped(drawingId) {
    return this.getDrawingParentGroup(drawingId) !== null;
  }
  _applyGroupDrawingOperation(groupParams, type) {
    if (groupParams.length === 0) return false;
    const sheetDrawingService = this._injector.get(ISheetDrawingService);
    const commandService = this._commandService;
    const undoRedoService = this._injector.get(IUndoRedoService);
    const jsonOp = type === 4 /* GROUP */ ? sheetDrawingService.getGroupDrawingOp(groupParams) : sheetDrawingService.getUngroupDrawingOp(groupParams);
    const { unitId, subUnitId, undo, redo, objects } = jsonOp;
    const result = commandService.syncExecuteCommand(SetDrawingApplyMutation.id, { op: redo, unitId, subUnitId, objects, type });
    if (result) {
      const inverseType = type === 4 /* GROUP */ ? 5 /* UNGROUP */ : 4 /* GROUP */;
      undoRedoService.pushUndoRedo({
        unitID: unitId,
        undoMutations: [
          { id: SetDrawingApplyMutation.id, params: { op: undo, unitId, subUnitId, objects: this._invertGroupOperationObjects(objects, type), type: inverseType } }
        ],
        redoMutations: [
          { id: SetDrawingApplyMutation.id, params: { op: redo, unitId, subUnitId, objects, type } }
        ]
      });
      return true;
    }
    return false;
  }
  _invertGroupOperationObjects(groupParams, type) {
    return type === 4 /* GROUP */ ? groupParams.map((groupParam) => {
      const { parent, children } = groupParam;
      const groupTransform = parent.transform || { width: 0, height: 0 };
      return {
        parent: {
          ...parent,
          transform: {
            left: 0,
            top: 0
          }
        },
        children: children.map((child) => ({
          ...child,
          transform: transformObjectOutOfGroup(
            child.transform || {},
            groupTransform,
            groupTransform.width || 0,
            groupTransform.height || 0,
            parent.groupBaseBound
          ),
          groupId: void 0
        }))
      };
    }) : groupParams.map((groupParam) => {
      const { parent, children } = groupParam;
      const groupTransform = getGroupState(0, 0, children.map((child) => child.transform || {}));
      return {
        parent: {
          ...parent,
          transform: groupTransform
        },
        children: children.map((child) => ({
          ...child,
          groupId: parent.drawingId
        }))
      };
    });
  }
};
FWorksheet.extend(FWorksheetDrawingMixin);

// ../packages/sheets-drawing/src/facade/f-enum.ts
var FSheetsDrawingEnumMixin = class extends FEnum {
  get DrawingType() {
    return DrawingTypeEnum;
  }
  get ImageSourceType() {
    return ImageSourceType;
  }
  get SheetDrawingAnchorType() {
    return SheetDrawingAnchorType;
  }
};
FEnum.extend(FSheetsDrawingEnumMixin);

// ../packages/sheets-drawing/src/facade/f-event.ts
var FSheetsDrawingEventNameMixin = class extends FEventName {
  get BeforeFloatDomAdd() {
    return "BeforeFloatDomAdd";
  }
  get FloatDomAdded() {
    return "FloatDomAdded";
  }
  get BeforeFloatDomUpdate() {
    return "BeforeFloatDomUpdate";
  }
  get FloatDomUpdated() {
    return "FloatDomUpdated";
  }
  get BeforeFloatDomDelete() {
    return "BeforeFloatDomDelete";
  }
  get FloatDomDeleted() {
    return "FloatDomDeleted";
  }
  get BeforeOverGridImageChange() {
    return "BeforeOverGridImageChange";
  }
  get OverGridImageChanged() {
    return "OverGridImageChanged";
  }
  get BeforeOverGridImageInsert() {
    return "BeforeOverGridImageInsert";
  }
  get OverGridImageInserted() {
    return "OverGridImageInserted";
  }
  get BeforeOverGridImageRemove() {
    return "BeforeOverGridImageRemove";
  }
  get OverGridImageRemoved() {
    return "OverGridImageRemoved";
  }
  get BeforeOverGridImageSelect() {
    return "BeforeOverGridImageSelect";
  }
  get OverGridImageSelected() {
    return "OverGridImageSelected";
  }
};
FEventName.extend(FSheetsDrawingEventNameMixin);

// ../packages/sheets-drawing/src/facade/f-univer.ts
var FUniverSheetsDrawingMixin = class extends FUniver {
  /**
   * @ignore
   */
  // eslint-disable-next-line max-lines-per-function
  _initialize(injector) {
    const commandService = injector.get(ICommandService);
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeOverGridImageInsert,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          if (commandInfo.id !== InsertSheetDrawingCommand.id) return;
          const params = commandInfo.params;
          const workbook = this.getActiveWorkbook();
          if (workbook == null || params == null) {
            return;
          }
          const { drawings } = params;
          const eventParams = {
            workbook,
            insertImageParams: drawings
          };
          this.fireEvent(this.Event.BeforeOverGridImageInsert, eventParams);
          if (eventParams.cancel) {
            throw new CanceledError();
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeOverGridImageRemove,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          if (commandInfo.id !== RemoveSheetDrawingCommand.id) return;
          const params = commandInfo.params;
          const workbook = this.getActiveWorkbook();
          if (workbook == null || params == null) {
            return;
          }
          const drawingManagerService = injector.get(IDrawingManagerService);
          const { drawings } = params;
          const willRemoveDrawings = drawings.map((drawing) => {
            return drawingManagerService.getDrawingByParam(drawing);
          });
          const eventParams = {
            workbook,
            images: this._createFOverGridImage(willRemoveDrawings)
          };
          this.fireEvent(this.Event.BeforeOverGridImageRemove, eventParams);
          if (eventParams.cancel) {
            throw new CanceledError();
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeOverGridImageChange,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          if (commandInfo.id !== SetSheetDrawingCommand.id) return;
          const params = commandInfo.params;
          const workbook = this.getActiveWorkbook();
          if (workbook == null || params == null) {
            return;
          }
          const { drawings } = params;
          const drawingManagerService = injector.get(IDrawingManagerService);
          const images = [];
          drawings.forEach((drawing) => {
            const image = drawingManagerService.getDrawingByParam(drawing);
            if (image == null) {
              return;
            }
            images.push({
              changeParam: drawing,
              image: this._injector.createInstance(FOverGridImage, image)
            });
          });
          const eventParams = {
            workbook,
            images
          };
          this.fireEvent(this.Event.BeforeOverGridImageChange, eventParams);
          if (eventParams.cancel) {
            drawingManagerService.updateNotification(drawings);
            throw new CanceledError();
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeOverGridImageSelect,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          if (commandInfo.id !== SetDrawingSelectedOperation.id) return;
          const drawings = commandInfo.params;
          const workbook = this.getActiveWorkbook();
          if (workbook == null) {
            return;
          }
          const drawingManagerService = injector.get(IDrawingManagerService);
          const oldSelectedDrawings = drawingManagerService.getFocusDrawings();
          const selectedDrawings = drawings.map((drawing) => {
            return drawingManagerService.getDrawingByParam(drawing);
          });
          const eventParams = {
            workbook,
            selectedImages: this._createFOverGridImage(selectedDrawings),
            oldSelectedImages: this._createFOverGridImage(oldSelectedDrawings)
          };
          this.fireEvent(this.Event.BeforeOverGridImageSelect, eventParams);
          if (eventParams.cancel) {
            throw new CanceledError();
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.OverGridImageInserted,
        () => commandService.onCommandExecuted((commandInfo) => {
          if (commandInfo.id !== InsertSheetDrawingCommand.id) return;
          const params = commandInfo.params;
          const workbook = this.getActiveWorkbook();
          if (workbook == null || params == null) {
            return;
          }
          const { drawings } = params;
          const eventParams = {
            workbook,
            images: this._createFOverGridImage(drawings)
          };
          this.fireEvent(this.Event.OverGridImageInserted, eventParams);
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.OverGridImageRemoved,
        () => commandService.onCommandExecuted((commandInfo) => {
          if (commandInfo.id !== RemoveSheetDrawingCommand.id) return;
          const params = commandInfo.params;
          const workbook = this.getActiveWorkbook();
          if (workbook == null || params == null) {
            return;
          }
          const { drawings } = params;
          const eventParams = {
            workbook,
            removeImageParams: drawings
          };
          this.fireEvent(this.Event.OverGridImageRemoved, eventParams);
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.OverGridImageChanged,
        () => commandService.onCommandExecuted((commandInfo) => {
          if (commandInfo.id !== SetSheetDrawingCommand.id) return;
          const params = commandInfo.params;
          const workbook = this.getActiveWorkbook();
          if (workbook == null || params == null) {
            return;
          }
          const { drawings } = params;
          const drawingManagerService = injector.get(IDrawingManagerService);
          const images = drawings.map((drawing) => {
            return this._injector.createInstance(FOverGridImage, drawingManagerService.getDrawingByParam(drawing));
          });
          const eventParams = {
            workbook,
            images
          };
          this.fireEvent(this.Event.OverGridImageChanged, eventParams);
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.OverGridImageSelected,
        () => commandService.onCommandExecuted((commandInfo) => {
          if (commandInfo.id !== SetDrawingSelectedOperation.id) return;
          const drawings = commandInfo.params;
          const workbook = this.getActiveWorkbook();
          if (workbook == null) {
            return;
          }
          const drawingManagerService = injector.get(IDrawingManagerService);
          const selectedDrawings = drawings.map((drawing) => {
            return drawingManagerService.getDrawingByParam(drawing);
          });
          const eventParams = {
            workbook,
            selectedImages: this._createFOverGridImage(selectedDrawings)
          };
          this.fireEvent(this.Event.OverGridImageSelected, eventParams);
        })
      )
    );
  }
  _createFOverGridImage(drawings) {
    return drawings.map((drawing) => {
      return this._injector.createInstance(FOverGridImage, drawing);
    });
  }
};
FUniver.extend(FUniverSheetsDrawingMixin);

// ../packages/sheets-drawing-ui/src/facade/f-worksheet.ts
var FWorksheetDrawingUIMixin = class extends FWorksheet {
  getFloatDomById(id) {
    const floatDomService = this._injector.get(SheetCanvasFloatDomManagerService);
    const info = floatDomService.getFloatDomInfo(id);
    if (!info) return null;
    const { unitId, subUnitId } = info;
    const { rect } = info;
    const state = rect.getState();
    const { left = 0, top = 0, width = 0, height = 0, flipX = false, flipY = false, angle = 0, skewX = 0, skewY = 0 } = state;
    const drawingParm = this._injector.get(ISheetDrawingService).getDrawingByParam({
      drawingId: info.id,
      unitId,
      subUnitId
    });
    if (!drawingParm) return null;
    return {
      position: {
        left,
        top,
        width,
        height,
        flipX,
        flipY,
        angle,
        skewX,
        skewY
      },
      componentKey: drawingParm.componentKey,
      allowTransform: drawingParm.allowTransform,
      data: drawingParm.data,
      id: info.id
    };
  }
  getAllFloatDoms() {
    const floatDomService = this._injector.get(SheetCanvasFloatDomManagerService);
    const unitId = this._workbook.getUnitId();
    const subUnitId = this._worksheet.getSheetId();
    return Array.from(floatDomService.getFloatDomsBySubUnitId(unitId, subUnitId).values()).map((info) => {
      const { rect } = info;
      const drawingParm = this._injector.get(ISheetDrawingService).getDrawingByParam({
        drawingId: info.id,
        unitId,
        subUnitId
      });
      const { left, top, width, height, flipX, flipY, angle, skewX, skewY } = rect.getState();
      return {
        position: {
          left,
          top,
          width,
          height,
          flipX,
          flipY,
          angle,
          skewX,
          skewY
        },
        componentKey: drawingParm.componentKey,
        allowTransform: drawingParm.allowTransform,
        data: drawingParm.data,
        id: info.id
      };
    });
  }
  updateFloatDom(id, config) {
    var _a, _b;
    const sheetSkeletonService = this._injector.get(SheetSkeletonService);
    const floatDomService = this._injector.get(SheetCanvasFloatDomManagerService);
    const info = floatDomService.getFloatDomInfo(id);
    if (!info) return this;
    const { unitId, subUnitId } = info;
    const skeleton = sheetSkeletonService.getSkeleton(unitId, subUnitId);
    if (!skeleton) return this;
    const drawingParm = this._injector.get(ISheetDrawingService).getDrawingByParam({
      unitId,
      subUnitId,
      drawingId: id
    });
    const newParam = {
      ...drawingParm,
      componentKey: config.componentKey || drawingParm.componentKey,
      allowTransform: config.allowTransform !== void 0 ? config.allowTransform : drawingParm.allowTransform,
      data: config.data || drawingParm.data,
      sheetTransform: config.position ? (_a = transformToDrawingPosition(config.position, skeleton)) != null ? _a : drawingParm.sheetTransform : drawingParm.sheetTransform,
      transform: {
        ...drawingParm.transform,
        ...config.position
        // Merge with existing transform
      },
      axisAlignSheetTransform: config.position ? (_b = transformToAxisAlignPosition(config.position, skeleton)) != null ? _b : drawingParm.sheetTransform : drawingParm.sheetTransform
    };
    const res = this._commandService.syncExecuteCommand(SetSheetDrawingCommand.id, { unitId, subUnitId, drawings: [newParam] });
    if (!res) {
      throw new Error("updateFloatDom failed");
    }
    return this;
  }
  batchUpdateFloatDoms(updates) {
    var _a, _b;
    const floatDomService = this._injector.get(SheetCanvasFloatDomManagerService);
    const drawingService = this._injector.get(ISheetDrawingService);
    const sheetSkeletonService = this._injector.get(SheetSkeletonService);
    const drawings = [];
    for (const update of updates) {
      const info = floatDomService.getFloatDomInfo(update.id);
      if (!info) continue;
      const { unitId, subUnitId } = info;
      const skeleton = sheetSkeletonService.getSkeleton(unitId, subUnitId);
      if (!skeleton) continue;
      const drawingParm = drawingService.getDrawingByParam({
        unitId,
        subUnitId,
        drawingId: update.id
      });
      if (!drawingParm) continue;
      const newParam = {
        ...drawingParm,
        componentKey: update.config.componentKey || drawingParm.componentKey,
        allowTransform: update.config.allowTransform !== void 0 ? update.config.allowTransform : drawingParm.allowTransform,
        data: update.config.data || drawingParm.data,
        sheetTransform: update.config.position ? (_a = transformToDrawingPosition(update.config.position, skeleton)) != null ? _a : drawingParm.sheetTransform : drawingParm.sheetTransform,
        transform: {
          ...drawingParm.transform,
          ...update.config.position
          // Merge with existing transform
        },
        axisAlignSheetTransform: update.config.position ? (_b = transformToAxisAlignPosition(update.config.position, skeleton)) != null ? _b : drawingParm.sheetTransform : drawingParm.sheetTransform
      };
      drawings.push(newParam);
    }
    if (drawings.length > 0) {
      const unitId = this._workbook.getUnitId();
      const subUnitId = this._worksheet.getSheetId();
      const res = this._commandService.syncExecuteCommand(SetSheetDrawingCommand.id, { unitId, subUnitId, drawings });
      if (!res) {
        throw new Error("batchUpdateFloatDoms failed");
      }
    }
    return this;
  }
  removeFloatDom(id) {
    const floatDomService = this._injector.get(SheetCanvasFloatDomManagerService);
    floatDomService.removeFloatDom(id);
    return this;
  }
  addFloatDomToPosition(layer, id) {
    const unitId = this._workbook.getUnitId();
    const subUnitId = this._worksheet.getSheetId();
    const { key, disposableCollection } = transformComponentKey(layer, this._injector.get(ComponentManager));
    const floatDomService = this._injector.get(SheetCanvasFloatDomManagerService);
    const res = floatDomService.addFloatDomToPosition({ ...layer, componentKey: key, unitId, subUnitId }, id);
    if (res) {
      disposableCollection.add(res.dispose);
      return {
        id: res.id,
        dispose: () => {
          disposableCollection.dispose();
          res.dispose();
        }
      };
    }
    disposableCollection.dispose();
    return null;
  }
  addFloatDomToRange(fRange, layer, domLayout, id) {
    const unitId = this._workbook.getUnitId();
    const subUnitId = this._worksheet.getSheetId();
    const { key, disposableCollection } = transformComponentKey(layer, this._injector.get(ComponentManager));
    const floatDomService = this._injector.get(SheetCanvasFloatDomManagerService);
    const res = floatDomService.addFloatDomToRange(fRange.getRange(), { ...layer, componentKey: key, unitId, subUnitId }, domLayout, id);
    if (res) {
      disposableCollection.add(res.dispose);
      return {
        id: res.id,
        dispose: () => {
          disposableCollection.dispose();
          res.dispose();
        }
      };
    }
    disposableCollection.dispose();
    return null;
  }
  addFloatDomToColumnHeader(column, layer, domLayout, id) {
    const unitId = this._workbook.getUnitId();
    const subUnitId = this._worksheet.getSheetId();
    const { key, disposableCollection } = transformComponentKey(layer, this._injector.get(ComponentManager));
    const floatDomService = this._injector.get(SheetCanvasFloatDomManagerService);
    const domRangeDispose = floatDomService.addFloatDomToColumnHeader(column, { ...layer, componentKey: key, unitId, subUnitId }, domLayout, id);
    if (domRangeDispose) {
      disposableCollection.add(domRangeDispose.dispose);
      return {
        id: domRangeDispose.id,
        dispose: () => {
          disposableCollection.dispose();
          domRangeDispose.dispose();
        }
      };
    }
    disposableCollection.dispose();
    return null;
  }
  async saveCellImagesAsync(options, ranges) {
    var _a;
    const batchSaveService = this._injector.get(IBatchSaveImagesService);
    const unitId = this._fWorkbook.getId();
    const subUnitId = this.getSheetId();
    const iRanges = ranges ? ranges.map((r) => r.getRange()) : [this._worksheet.getCellMatrix().getDataRange()];
    const images = batchSaveService.getCellImagesFromRanges(unitId, subUnitId, iRanges);
    if (images.length === 0) {
      return false;
    }
    if (images.length === 1) {
      try {
        await batchSaveService.downloadSingleImage(images[0]);
        return true;
      } catch (error) {
        console.error("Failed to download image:", error);
        return false;
      }
    }
    const fileNameParts = [];
    const useCellAddress = (_a = options == null ? void 0 : options.useCellAddress) != null ? _a : true;
    const useColumnIndex = options == null ? void 0 : options.useColumnIndex;
    if (useCellAddress) {
      fileNameParts.push("cellAddress" /* CELL_ADDRESS */);
    }
    if (useColumnIndex !== void 0) {
      fileNameParts.push("columnValue" /* COLUMN_VALUE */);
    }
    if (fileNameParts.length === 0) {
      fileNameParts.push("cellAddress" /* CELL_ADDRESS */);
    }
    try {
      await batchSaveService.saveImagesWithContext(images, {
        fileNameParts,
        columnIndex: useColumnIndex
      }, unitId, subUnitId);
      return true;
    } catch (error) {
      console.error("Failed to save images:", error);
      return false;
    }
  }
};
FWorksheet.extend(FWorksheetDrawingUIMixin);

// ../packages/sheets-drawing-ui/src/facade/f-event.ts
var FSheetsDrawingUIEventNameMixin = class extends FEventName {
  get BeforeFloatDomAdd() {
    return "BeforeFloatDomAdd";
  }
  get FloatDomAdded() {
    return "FloatDomAdded";
  }
  get BeforeFloatDomUpdate() {
    return "BeforeFloatDomUpdate";
  }
  get FloatDomUpdated() {
    return "FloatDomUpdated";
  }
  get BeforeFloatDomDelete() {
    return "BeforeFloatDomDelete";
  }
  get FloatDomDeleted() {
    return "FloatDomDeleted";
  }
};
FEventName.extend(FSheetsDrawingUIEventNameMixin);

// ../packages/sheets-drawing-ui/src/facade/f-univer.ts
var FUniverSheetsDrawingUIMixin = class extends FUniver {
  /**
   * @ignore
   */
  // eslint-disable-next-line max-lines-per-function
  _initialize(injector) {
    const commandService = injector.get(ICommandService);
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeFloatDomAdd,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          if (commandInfo.id !== InsertSheetDrawingCommand.id) return;
          const params = commandInfo.params;
          const workbook = this.getActiveWorkbook();
          if (workbook == null || params == null) {
            return;
          }
          const { drawings } = params;
          const floatDomDrawings = drawings.filter(
            (drawing) => drawing.drawingType === 8 /* DRAWING_DOM */
          );
          if (floatDomDrawings.length === 0) {
            return;
          }
          const eventParams = {
            workbook,
            drawings: floatDomDrawings
          };
          this.fireEvent(this.Event.BeforeFloatDomAdd, eventParams);
          if (eventParams.cancel) {
            throw new CanceledError();
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.FloatDomAdded,
        () => commandService.onCommandExecuted((commandInfo) => {
          if (commandInfo.id !== InsertSheetDrawingCommand.id) return;
          const params = commandInfo.params;
          const workbook = this.getActiveWorkbook();
          if (workbook == null || params == null) {
            return;
          }
          const { drawings } = params;
          const floatDomDrawings = drawings.filter(
            (drawing) => drawing.drawingType === 8 /* DRAWING_DOM */
          );
          if (floatDomDrawings.length === 0) {
            return;
          }
          const eventParams = {
            workbook,
            drawings: floatDomDrawings
          };
          this.fireEvent(this.Event.FloatDomAdded, eventParams);
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeFloatDomUpdate,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          if (commandInfo.id !== SetSheetDrawingCommand.id) return;
          const params = commandInfo.params;
          const workbook = this.getActiveWorkbook();
          if (workbook == null || params == null) {
            return;
          }
          const { drawings } = params;
          const drawingManagerService = injector.get(IDrawingManagerService);
          const floatDomDrawings = [];
          drawings.forEach((drawing) => {
            const dom = drawingManagerService.getDrawingByParam(drawing);
            if ((dom == null ? void 0 : dom.drawingType) === 8 /* DRAWING_DOM */) {
              floatDomDrawings.push(dom);
            }
          });
          if (floatDomDrawings.length === 0) {
            return;
          }
          const eventParams = {
            workbook,
            drawings: floatDomDrawings
          };
          this.fireEvent(this.Event.BeforeFloatDomUpdate, eventParams);
          if (eventParams.cancel) {
            drawingManagerService.updateNotification(drawings);
            throw new CanceledError();
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.FloatDomUpdated,
        () => commandService.onCommandExecuted((commandInfo) => {
          if (commandInfo.id !== SetSheetDrawingCommand.id) return;
          const params = commandInfo.params;
          const workbook = this.getActiveWorkbook();
          if (workbook == null || params == null) {
            return;
          }
          const { drawings } = params;
          const drawingManagerService = injector.get(IDrawingManagerService);
          const floatDomDrawings = [];
          drawings.forEach((drawing) => {
            const dom = drawingManagerService.getDrawingByParam(drawing);
            if ((dom == null ? void 0 : dom.drawingType) === 8 /* DRAWING_DOM */) {
              floatDomDrawings.push(dom);
            }
          });
          if (floatDomDrawings.length === 0) {
            return;
          }
          const eventParams = {
            workbook,
            drawings: floatDomDrawings
          };
          this.fireEvent(this.Event.FloatDomUpdated, eventParams);
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeFloatDomDelete,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          if (commandInfo.id !== RemoveSheetDrawingCommand.id) return;
          const params = commandInfo.params;
          const workbook = this.getActiveWorkbook();
          if (workbook == null || params == null) {
            return;
          }
          const drawingManagerService = injector.get(IDrawingManagerService);
          const { drawings } = params;
          const floatDomDrawings = drawings.map((drawing) => drawingManagerService.getDrawingByParam(drawing)).filter(
            (drawing) => (drawing == null ? void 0 : drawing.drawingType) === 8 /* DRAWING_DOM */
          );
          if (floatDomDrawings.length === 0) {
            return;
          }
          const eventParams = {
            workbook,
            drawings: floatDomDrawings
          };
          this.fireEvent(this.Event.BeforeFloatDomDelete, eventParams);
          if (eventParams.cancel) {
            throw new CanceledError();
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.FloatDomDeleted,
        () => commandService.onCommandExecuted((commandInfo) => {
          if (commandInfo.id !== RemoveSheetDrawingCommand.id) return;
          const params = commandInfo.params;
          const workbook = this.getActiveWorkbook();
          if (workbook == null || params == null) {
            return;
          }
          const { drawings } = params;
          const eventParams = {
            workbook,
            drawings: drawings.filter((i) => i.drawingType === 8 /* DRAWING_DOM */).map((i) => i.drawingId)
          };
          this.fireEvent(this.Event.FloatDomDeleted, eventParams);
        })
      )
    );
  }
  registerURLImageDownloader(downloader) {
    const urlImageService = this._injector.get(IURLImageService);
    return urlImageService.registerURLImageDownloader(downloader);
  }
};
FUniver.extend(FUniverSheetsDrawingUIMixin);

// ../packages/sheets-drawing-ui/src/facade/f-range.ts
var FRangeSheetsDrawingUIMixin = class extends FRange {
  async insertCellImageAsync(file) {
    var _a;
    const renderManagerService = this._injector.get(IRenderManagerService);
    const controller = (_a = getCurrentTypeOfRenderer(2 /* UNIVER_SHEET */, this._injector.get(IUniverInstanceService), renderManagerService)) == null ? void 0 : _a.with(SheetDrawingUpdateController);
    if (!controller) {
      return false;
    }
    const location = {
      unitId: this._workbook.getUnitId(),
      subUnitId: this._worksheet.getSheetId(),
      row: this.getRow(),
      col: this.getColumn()
    };
    if (typeof file === "string") {
      return controller.insertCellImageByUrl(file, location);
    } else {
      return controller.insertCellImageByFile(file, location);
    }
  }
  async saveCellImagesAsync(options) {
    var _a;
    const batchSaveService = this._injector.get(IBatchSaveImagesService);
    const unitId = this._workbook.getUnitId();
    const subUnitId = this._worksheet.getSheetId();
    const range = this.getRange();
    const images = batchSaveService.getCellImagesFromRanges(unitId, subUnitId, [range]);
    if (images.length === 0) {
      return false;
    }
    if (images.length === 1) {
      try {
        await batchSaveService.downloadSingleImage(images[0]);
        return true;
      } catch (error) {
        console.error("Failed to download image:", error);
        return false;
      }
    }
    const fileNameParts = [];
    const useCellAddress = (_a = options == null ? void 0 : options.useCellAddress) != null ? _a : true;
    const useColumnIndex = options == null ? void 0 : options.useColumnIndex;
    if (useCellAddress) {
      fileNameParts.push("cellAddress" /* CELL_ADDRESS */);
    }
    if (useColumnIndex !== void 0) {
      fileNameParts.push("columnValue" /* COLUMN_VALUE */);
    }
    if (fileNameParts.length === 0) {
      fileNameParts.push("cellAddress" /* CELL_ADDRESS */);
    }
    try {
      await batchSaveService.saveImagesWithContext(images, {
        fileNameParts,
        columnIndex: useColumnIndex
      }, unitId, subUnitId);
      return true;
    } catch (error) {
      console.error("Failed to save images:", error);
      return false;
    }
  }
};
FRange.extend(FRangeSheetsDrawingUIMixin);

// ../packages/sheets-crosshair-highlight/src/facade/f-univer.ts
var FUniverSheetsCrosshairHighlightMixin = class extends FUniver {
  /**
   * @ignore
   */
  _initialize(injector) {
    const commandService = injector.get(ICommandService);
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.CrosshairHighlightEnabledChanged,
        () => commandService.onCommandExecuted((commandInfo) => {
          if (commandInfo.id === EnableCrosshairHighlightOperation.id || commandInfo.id === DisableCrosshairHighlightOperation.id || commandInfo.id === ToggleCrosshairHighlightOperation.id) {
            const activeSheet = this.getActiveSheet();
            if (!activeSheet) return;
            const eventParams = {
              enabled: this.getCrosshairHighlightEnabled(),
              ...activeSheet
            };
            this.fireEvent(this.Event.CrosshairHighlightEnabledChanged, eventParams);
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.CrosshairHighlightColorChanged,
        () => commandService.onCommandExecuted((commandInfo) => {
          if (commandInfo.id === SetCrosshairHighlightColorOperation.id) {
            const activeSheet = this.getActiveSheet();
            if (!activeSheet) return;
            const eventParams = {
              color: this.getCrosshairHighlightColor(),
              ...activeSheet
            };
            this.fireEvent(this.Event.CrosshairHighlightColorChanged, eventParams);
          }
        })
      )
    );
  }
  setCrosshairHighlightEnabled(enabled) {
    if (enabled) {
      this._commandService.syncExecuteCommand(EnableCrosshairHighlightOperation.id);
    } else {
      this._commandService.syncExecuteCommand(DisableCrosshairHighlightOperation.id);
    }
    return this;
  }
  setCrosshairHighlightColor(color) {
    this._commandService.syncExecuteCommand(SetCrosshairHighlightColorOperation.id, {
      value: color
    });
    return this;
  }
  getCrosshairHighlightEnabled() {
    const crosshairHighlightService = this._injector.get(SheetsCrosshairHighlightService);
    return crosshairHighlightService.enabled;
  }
  getCrosshairHighlightColor() {
    const crosshairHighlightService = this._injector.get(SheetsCrosshairHighlightService);
    return crosshairHighlightService.color;
  }
  get CROSSHAIR_HIGHLIGHT_COLORS() {
    return CROSSHAIR_HIGHLIGHT_COLORS;
  }
};
FUniver.extend(FUniverSheetsCrosshairHighlightMixin);

// ../packages/sheets-crosshair-highlight/src/facade/f-event.ts
var FSheetsCrosshairHighlightEventNameMixin = class extends FEventName {
  get CrosshairHighlightEnabledChanged() {
    return "CrosshairHighlightEnabledChanged";
  }
  get CrosshairHighlightColorChanged() {
    return "CrosshairHighlightColorChanged";
  }
};
FEventName.extend(FSheetsCrosshairHighlightEventNameMixin);

// ../packages/sheets-sort/src/facade/f-univer.ts
var FUniverSheetsSortMixin = class extends FUniver {
  /**
   * @ignore
   */
  _initialize(injector) {
    const commandService = injector.get(ICommandService);
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.SheetBeforeRangeSort,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          if (commandInfo.id !== SortRangeCommand.id) return;
          this._beforeRangeSort(commandInfo);
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.SheetRangeSorted,
        () => commandService.onCommandExecuted((commandInfo) => {
          if (commandInfo.id !== SortRangeCommand.id) return;
          this._onRangeSorted(commandInfo);
        })
      )
    );
  }
  _beforeRangeSort(commandInfo) {
    const params = commandInfo.params;
    const fWorkbook = this.getWorkbook(params.unitId);
    const fWorksheet = fWorkbook.getSheetBySheetId(params.subUnitId);
    const { startColumn, endColumn, startRow, endRow } = params.range;
    const fRange = fWorksheet.getRange(startRow, startColumn, endRow - startRow + 1, endColumn - startColumn + 1);
    const eventParams = {
      workbook: fWorkbook,
      worksheet: fWorksheet,
      range: fRange,
      sortColumn: params.orderRules.map((rule) => ({
        column: rule.colIndex - startColumn,
        ascending: rule.type === "asc" /* ASC */
      }))
    };
    this.fireEvent(this.Event.SheetBeforeRangeSort, eventParams);
    if (eventParams.cancel) {
      throw new Error("SortRangeCommand canceled.");
    }
  }
  _onRangeSorted(commandInfo) {
    const params = commandInfo.params;
    const fWorkbook = this.getWorkbook(params.unitId);
    const fWorksheet = fWorkbook.getSheetBySheetId(params.subUnitId);
    const { startColumn, endColumn, startRow, endRow } = params.range;
    const fRange = fWorksheet.getRange(startRow, startColumn, endRow - startRow + 1, endColumn - startColumn + 1);
    const eventParams = {
      workbook: fWorkbook,
      worksheet: fWorksheet,
      range: fRange,
      sortColumn: params.orderRules.map((rule) => ({
        column: rule.colIndex - startColumn,
        ascending: rule.type === "asc" /* ASC */
      }))
    };
    this.fireEvent(this.Event.SheetRangeSorted, eventParams);
    if (eventParams.cancel) {
      throw new Error("SortRangeCommand canceled.");
    }
  }
};
FUniver.extend(FUniverSheetsSortMixin);

// ../packages/sheets-sort/src/facade/f-range.ts
var FRangeSheetsSortMixin = class extends FRange {
  sort(column) {
    const columnBase = this._range.startColumn;
    const columns = Array.isArray(column) ? column : [column];
    const orderRules = columns.map((c) => {
      if (typeof c === "number") {
        return { colIndex: c + columnBase, type: "asc" /* ASC */ };
      }
      return {
        colIndex: c.column + columnBase,
        type: c.ascending ? "asc" /* ASC */ : "desc" /* DESC */
      };
    });
    this._commandService.syncExecuteCommand(SortRangeCommand.id, {
      orderRules,
      range: this._range,
      hasTitle: false,
      unitId: this._workbook.getUnitId(),
      subUnitId: this._worksheet.getSheetId()
    });
    return this;
  }
};
FRange.extend(FRangeSheetsSortMixin);

// ../packages/sheets-sort/src/facade/f-worksheet.ts
var FWorksheetSortMixin = class extends FWorksheet {
  sort(colIndex, asc = true) {
    const orderRules = [{
      colIndex,
      type: asc ? "asc" /* ASC */ : "desc" /* DESC */
    }];
    const range = {
      startRow: 0,
      startColumn: 0,
      endRow: this._worksheet.getRowCount() - 1,
      endColumn: this._worksheet.getColumnCount() - 1,
      rangeType: 3 /* ALL */
    };
    this._commandService.syncExecuteCommand(SortRangeCommand.id, {
      orderRules,
      range,
      hasTitle: false,
      unitId: this._workbook.getUnitId(),
      subUnitId: this._worksheet.getSheetId()
    });
    return this;
  }
};
FWorksheet.extend(FWorksheetSortMixin);

// ../packages/sheets-sort/src/facade/f-event.ts
var FSheetsSortEventNameMixin = class {
  get SheetRangeSorted() {
    return "SheetRangeSorted";
  }
  get SheetBeforeRangeSort() {
    return "SheetBeforeRangeSort";
  }
};
FEventName.extend(FSheetsSortEventNameMixin);
