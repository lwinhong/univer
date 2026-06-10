import {
  FRange,
  FWorkbook,
  FWorksheet
} from "./chunk-BKEMWLI4.js";
import {
  SetNumfmtCommand,
  SheetsNumfmtCellContentController
} from "./chunk-S7WCECRJ.js";
import {
  CellAlertManagerService,
  DragManagerService,
  HoverManagerService,
  IEditorBridgeService,
  IMarkSelectionService,
  ISheetCellDropdownManagerService,
  ISheetClipboardService,
  ISheetSelectionRenderService,
  SetCellEditVisibleOperation,
  SetColumnHeaderHeightCommand,
  SetRowHeaderWidthCommand,
  SetWorksheetColAutoWidthCommand,
  SetZoomRatioCommand,
  SheetCanvasPopManagerService,
  SheetPasteShortKeyCommand,
  SheetPermissionRenderManagerService,
  SheetScrollManagerService,
  SheetSkeletonManagerService,
  SheetsScrollRenderController
} from "./chunk-SHS2QU7R.js";
import {
  RichTextEditingMutation
} from "./chunk-NDZAFLT3.js";
import {
  BuiltInUIPart,
  ComponentManager,
  CopyCommand,
  CutCommand,
  HTML_CLIPBOARD_MIME_TYPE,
  IClipboardInterfaceService,
  IDialogService,
  IFontService,
  IMenuManagerService,
  IMessageService,
  IShortcutService,
  ISidebarService,
  IUIPartsService,
  KeyCode,
  MenuManagerPosition,
  PLAIN_TEXT_CLIPBOARD_MIME_TYPE,
  PasteCommand,
  RibbonPosition,
  RibbonStartGroup,
  SheetPasteShortKeyCommandName,
  connectInjector,
  supportClipboardAPI
} from "./chunk-JUXP4HVV.js";
import {
  FBase,
  FEnum,
  FEventName,
  FHooks,
  FUniver
} from "./chunk-66S2PCKH.js";
import {
  CalculationMode,
  FormulaCalculationSessionService,
  IRegisterFunctionService,
  PLUGIN_CONFIG_KEY_BASE,
  RegisterFunctionService
} from "./chunk-T3S76CBR.js";
import {
  COMMAND_LISTENER_SKELETON_CHANGE,
  ENGINE_FORMULA_CYCLE_REFERENCE_COUNT,
  ENGINE_FORMULA_RETURN_DEPENDENCY_TREE,
  FormulaDataModel,
  GlobalComputingStatusService,
  IDefinedNamesService,
  IFunctionService,
  INTERCEPTOR_POINT,
  ISuperTableService,
  LexerTreeBuilder,
  SetCellFormulaDependencyCalculationMutation,
  SetCellFormulaDependencyCalculationResultMutation,
  SetFormulaCalculationNotificationMutation,
  SetFormulaCalculationStartMutation,
  SetFormulaCalculationStopMutation,
  SetFormulaDependencyCalculationMutation,
  SetFormulaDependencyCalculationResultMutation,
  SetFormulaStringBatchCalculationMutation,
  SetFormulaStringBatchCalculationResultMutation,
  SetQueryFormulaDependencyAllMutation,
  SetQueryFormulaDependencyAllResultMutation,
  SetQueryFormulaDependencyMutation,
  SetQueryFormulaDependencyResultMutation,
  SetTriggerFormulaCalculationStartMutation,
  SetWorksheetRowIsAutoHeightCommand,
  SheetInterceptorService,
  SheetsSelectionsService,
  extractFormulaError,
  getSkeletonChangedEffectedRange
} from "./chunk-42RBW3MT.js";
import {
  IRenderManagerService,
  SHEET_VIEWPORT_KEY
} from "./chunk-KC3DSKIX.js";
import {
  CanceledError,
  DOCS_NORMAL_EDITOR_UNIT_ID_KEY,
  DisposableCollection,
  ICommandService,
  IConfigService,
  ILogService,
  IPermissionService,
  IUniverInstanceService,
  Inject,
  Injector,
  LifecycleService,
  RichTextValue,
  awaitTime,
  combineLatest,
  debounce_default,
  filter,
  firstValueFrom,
  generateRandomId,
  map,
  race,
  timer,
  toDisposable
} from "./chunk-33LYH36L.js";
import {
  __decorateClass,
  __decorateParam,
  __publicField
} from "./chunk-DO7PIA5W.js";

// ../packages/ui/src/facade/f-menu-builder.ts
var FMenuBase = class extends FBase {
  /**
   * Append the menu to any menu position on Univer UI.
   * @param {string | string[]} path - Some predefined path to append the menu. The paths can be an array,
   * or an array joined by `|` separator. Since lots of submenus reuse the same name,
   * you may need to specify their parent menus as well.
   *
   * @example
   * ```typescript
   * // This menu item will appear on every `contextMenu.others` section.
   * univerAPI.createMenu({
   *   id: 'custom-menu-id-1',
   *   title: 'Custom Menu 1',
   *   action: () => {
   *     console.log('Custom Menu 1 clicked');
   *   },
   * }).appendTo('contextMenu.others');
   *
   * // This menu item will only appear on the `contextMenu.others` section on the main area.
   * univerAPI.createMenu({
   *   id: 'custom-menu-id-2',
   *   title: 'Custom Menu 2',
   *   action: () => {
   *     console.log('Custom Menu 2 clicked');
   *   },
   * }).appendTo(['contextMenu.mainArea', 'contextMenu.others']);
   * ```
   */
  appendTo(path) {
    const paths = typeof path === "string" ? path.split("|") : path;
    const len = paths.length;
    const menuConfig = {};
    let obj = menuConfig;
    const schema = this.__getSchema();
    paths.forEach((p, index) => {
      if (index === len - 1) {
        obj[p] = schema;
      } else {
        obj[p] = {};
      }
      obj = obj[p];
    });
    this._menuManagerService.mergeMenu(menuConfig);
  }
};
var FMenu = class extends FMenuBase {
  constructor(_item, _injector, _commandService, _menuManagerService) {
    super();
    __publicField(this, "_item", _item);
    __publicField(this, "_injector", _injector);
    __publicField(this, "_commandService", _commandService);
    __publicField(this, "_menuManagerService", _menuManagerService);
    __publicField(this, "_commandToRegister", /* @__PURE__ */ new Map());
    __publicField(this, "_buildingSchema");
    const commandId = typeof _item.action === "string" ? _item.action : generateRandomId(12);
    if (commandId !== _item.action) {
      this._commandToRegister.set(commandId, _item.action);
    }
    this._buildingSchema = {
      // eslint-disable-next-line ts/explicit-function-return-type
      menuItemFactory: () => ({
        id: _item.id,
        type: 0 /* BUTTON */,
        // we only support button for now
        icon: _item.icon,
        title: _item.title,
        tooltip: _item.tooltip,
        commandId
      })
    };
    if (typeof _item.order !== "undefined") {
      this._buildingSchema.order = _item.order;
    }
  }
  /**
   * @ignore
   */
  __getSchema() {
    this._commandToRegister.forEach((command, id) => {
      if (!this._commandService.hasCommand(id)) {
        this._commandService.registerCommand({
          id,
          type: 0 /* COMMAND */,
          handler: command
        });
      }
    });
    return { [this._item.id]: this._buildingSchema };
  }
};
__publicField(FMenu, "RibbonStartGroup", RibbonStartGroup);
__publicField(FMenu, "RibbonPosition", RibbonPosition);
__publicField(FMenu, "MenuManagerPosition", MenuManagerPosition);
FMenu = __decorateClass([
  __decorateParam(1, Inject(Injector)),
  __decorateParam(2, ICommandService),
  __decorateParam(3, IMenuManagerService)
], FMenu);
var FSubmenu = class extends FMenuBase {
  constructor(_item, _injector, _menuManagerService) {
    super();
    __publicField(this, "_item", _item);
    __publicField(this, "_injector", _injector);
    __publicField(this, "_menuManagerService", _menuManagerService);
    __publicField(this, "_menuByGroups", []);
    __publicField(this, "_submenus", []);
    __publicField(this, "_buildingSchema");
    this._buildingSchema = {
      // eslint-disable-next-line ts/explicit-function-return-type
      menuItemFactory: () => ({
        id: _item.id,
        type: 3 /* SUBITEMS */,
        icon: _item.icon,
        title: _item.title,
        tooltip: _item.tooltip
      })
    };
    if (typeof _item.order !== "undefined") {
      this._buildingSchema.order = _item.order;
    }
  }
  /**
   * Add a menu to the submenu. It can be a {@link FMenu} or a {@link FSubmenu}.
   * @param {FMenu | FSubmenu} submenu - Menu to add to the submenu.
   * @returns {FSubmenu} The FSubmenu itself for chaining calls.
   * @example
   * ```typescript
   * // Create two leaf menus.
   * const menu1 = univerAPI.createMenu({
   *   id: 'submenu-nested-1',
   *   title: 'Item 1',
   *   action: () => {
   *     console.log('Item 1 clicked');
   *   }
   * });
   * const menu2 = univerAPI.createMenu({
   *   id: 'submenu-nested-2',
   *   title: 'Item 2',
   *   action: () => {
   *     console.log('Item 2 clicked');
   *   }
   * });
   *
   * // Add the leaf menus to a submenu.
   * const submenu = univerAPI.createSubmenu({ id: 'submenu-nested', title: 'Nested Submenu' })
   *   .addSubmenu(menu1)
   *   .addSeparator()
   *   .addSubmenu(menu2);
   *
   * // Create a root submenu append to the `contextMenu.others` section.
   * univerAPI.createSubmenu({ id: 'custom-submenu', title: 'Custom Submenu' })
   *   .addSubmenu(submenu)
   *   .appendTo('contextMenu.others');
   * ```
   */
  addSubmenu(submenu) {
    this._submenus.push(submenu);
    return this;
  }
  /**
   * Add a separator to the submenu.
   * @returns {FSubmenu} The FSubmenu itself for chaining calls.
   * @example
   * ```typescript
   * // Create two leaf menus.
   * const menu1 = univerAPI.createMenu({
   *   id: 'submenu-nested-1',
   *   title: 'Item 1',
   *   action: () => {
   *     console.log('Item 1 clicked');
   *   }
   * });
   * const menu2 = univerAPI.createMenu({
   *   id: 'submenu-nested-2',
   *   title: 'Item 2',
   *   action: () => {
   *     console.log('Item 2 clicked');
   *   }
   * });
   *
   * // Add the leaf menus to a submenu and add a separator between them.
   * // Append the submenu to the `contextMenu.others` section.
   * univerAPI.createSubmenu({ id: 'submenu-nested', title: 'Nested Submenu' })
   *   .addSubmenu(menu1)
   *   .addSeparator()
   *   .addSubmenu(menu2)
   *   .appendTo('contextMenu.others');
   * ```
   */
  addSeparator() {
    this._menuByGroups.push(this._submenus);
    this._submenus = [];
    return this;
  }
  /**
   * @ignore
   */
  __getSchema() {
    const schema = {};
    this.addSeparator();
    this._menuByGroups.forEach((group, index) => {
      const groupSchema = {};
      group.forEach((menu) => {
        Object.assign(groupSchema, menu.__getSchema());
      });
      schema[`${this._item.id}-group-${index}`] = groupSchema;
    });
    return { [this._item.id]: Object.assign(this._buildingSchema, schema) };
  }
};
FSubmenu = __decorateClass([
  __decorateParam(1, Inject(Injector)),
  __decorateParam(2, IMenuManagerService)
], FSubmenu);

// ../packages/ui/src/facade/f-shortcut.ts
var FShortcut = class extends FBase {
  constructor(_injector, _renderManagerService, _univerInstanceService, _shortcutService) {
    super();
    __publicField(this, "_injector", _injector);
    __publicField(this, "_renderManagerService", _renderManagerService);
    __publicField(this, "_univerInstanceService", _univerInstanceService);
    __publicField(this, "_shortcutService", _shortcutService);
    __publicField(this, "_forceDisableDisposable", null);
  }
  /**
   * Enable shortcuts of Univer.
   * @returns {FShortcut} The Facade API instance itself for chaining.
   *
   * @example
   * ```typescript
   * fShortcut.enableShortcut(); // Use the FShortcut instance used by disableShortcut before, do not create a new instance
   * ```
   */
  enableShortcut() {
    var _a;
    (_a = this._forceDisableDisposable) == null ? void 0 : _a.dispose();
    this._forceDisableDisposable = null;
    return this;
  }
  /**
   * Disable shortcuts of Univer.
   * @returns {FShortcut} The Facade API instance itself for chaining.
   *
   * @example
   * ```typescript
   * const fShortcut = univerAPI.getShortcut();
   * fShortcut.disableShortcut();
   * ```
   */
  disableShortcut() {
    if (!this._forceDisableDisposable) {
      this._forceDisableDisposable = this._shortcutService.forceDisable();
    }
    return this;
  }
  /**
   * Trigger shortcut of Univer by a KeyboardEvent and return the matched shortcut item.
   * @param {KeyboardEvent} e - The KeyboardEvent to trigger.
   * @returns {IShortcutItem<object> | undefined} The matched shortcut item.
   *
   * @example
   * ```typescript
   * // Assum the current sheet is empty sheet.
   * const fWorkbook = univerAPI.getActiveWorkbook();
   * const fWorksheet = fWorkbook.getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * const fRange = fWorksheet.getRange('A1');
   *
   * // Set A1 cell active and set value to 'Hello Univer'.
   * fRange.activate();
   * fRange.setValue('Hello Univer');
   * console.log(fRange.getCellStyle().bold); // false
   *
   * // Set A1 cell bold by shortcut.
   * const fShortcut = univerAPI.getShortcut();
   * const pseudoEvent = new KeyboardEvent('keydown', {
   *   key: 'b',
   *   ctrlKey: true,
   *   keyCode: univerAPI.Enum.KeyCode.B
   * });
   * const ifShortcutItem = fShortcut.triggerShortcut(pseudoEvent);
   * if (ifShortcutItem) {
   *   const commandId = ifShortcutItem.id;
   *   console.log(fRange.getCellStyle().bold); // true
   * }
   * ```
   */
  triggerShortcut(e) {
    const workbook = this._univerInstanceService.getCurrentUnitOfType(2 /* UNIVER_SHEET */);
    if (!workbook) {
      return;
    }
    const renderUnit = this._renderManagerService.getRenderById(workbook.getUnitId());
    if (!renderUnit) {
      return;
    }
    const canvas = renderUnit.engine.getCanvasElement();
    canvas.dispatchEvent(e);
    return this._shortcutService.dispatch(e);
  }
  /**
   * Dispatch a KeyboardEvent to the shortcut service and return the matched shortcut item.
   * @param {KeyboardEvent} e - The KeyboardEvent to dispatch.
   * @returns {IShortcutItem<object> | undefined} The matched shortcut item.
   *
   * @example
   * ```typescript
   * const fShortcut = univerAPI.getShortcut();
   * const pseudoEvent = new KeyboardEvent('keydown', { key: 's', ctrlKey: true });
   * const ifShortcutItem = fShortcut.dispatchShortcutEvent(pseudoEvent);
   * if (ifShortcutItem) {
   *   const commandId = ifShortcutItem.id;
   *   // Do something with the commandId.
   * }
   * ```
   */
  dispatchShortcutEvent(e) {
    return this._shortcutService.dispatch(e);
  }
};
FShortcut = __decorateClass([
  __decorateParam(0, Inject(Injector)),
  __decorateParam(1, Inject(IRenderManagerService)),
  __decorateParam(2, IUniverInstanceService),
  __decorateParam(3, IShortcutService)
], FShortcut);

// ../packages/ui/src/facade/f-univer.ts
var FUniverUIMixin = class extends FUniver {
  getURL() {
    return new URL(window.location.href);
  }
  getShortcut() {
    return this._injector.createInstance(FShortcut);
  }
  copy() {
    return this._commandService.executeCommand(CopyCommand.id);
  }
  paste() {
    return this._commandService.executeCommand(PasteCommand.id);
  }
  createMenu(menuItem) {
    return this._injector.createInstance(FMenu, menuItem);
  }
  createSubmenu(submenuItem) {
    return this._injector.createInstance(FSubmenu, submenuItem);
  }
  openSiderbar(params) {
    const sideBarService = this._injector.get(ISidebarService);
    return sideBarService.open(params);
  }
  openSidebar(params) {
    return this.openSiderbar(params);
  }
  openDialog(dialog) {
    const dialogService = this._injector.get(IDialogService);
    const disposable = dialogService.open({
      ...dialog,
      onClose: () => {
        disposable.dispose();
      }
    });
    return disposable;
  }
  getComponentManager() {
    return this._injector.get(ComponentManager);
  }
  showMessage(options) {
    const messageService = this._injector.get(IMessageService);
    messageService.show(options);
    return this;
  }
  setUIVisible(ui, visible) {
    const uiPartService = this._injector.get(IUIPartsService);
    uiPartService.setUIVisible(ui, visible);
    return this;
  }
  isUIVisible(ui) {
    const uiPartService = this._injector.get(IUIPartsService);
    return uiPartService.isUIVisible(ui);
  }
  registerUIPart(key, component) {
    const uiPartService = this._injector.get(IUIPartsService);
    return uiPartService.registerComponent(key, () => connectInjector(component, this._injector));
  }
  registerComponent(name, component, options) {
    const componentManager = this._injector.get(ComponentManager);
    return this.disposeWithMe(componentManager.register(name, component, options));
  }
  setCurrent(unitId) {
    const rendererManagerService = this._injector.get(IRenderManagerService);
    const renderUnit = rendererManagerService.getRenderById(unitId);
    if (!renderUnit) {
      throw new Error("Unit not found");
    }
    this._univerInstanceService.setCurrentUnitForType(unitId);
  }
  addFonts(fonts) {
    const fontService = this._injector.get(IFontService);
    fonts.forEach((font) => {
      fontService.addFont(font);
    });
  }
};
FUniver.extend(FUniverUIMixin);

// ../packages/ui/src/facade/f-hooks.ts
var FHooksSheetsMixin = class extends FHooks {
  onBeforeCopy(callback) {
    const commandService = this._injector.get(ICommandService);
    return commandService.beforeCommandExecuted((command) => {
      if (command.id === CopyCommand.id) {
        callback();
      }
    });
  }
  onCopy(callback) {
    const commandService = this._injector.get(ICommandService);
    return commandService.onCommandExecuted((command) => {
      if (command.id === CopyCommand.id) {
        callback();
      }
    });
  }
  onBeforePaste(callback) {
    const commandService = this._injector.get(ICommandService);
    return commandService.beforeCommandExecuted((command) => {
      if (command.id === PasteCommand.id) {
        callback();
      }
    });
  }
  onPaste(callback) {
    const commandService = this._injector.get(ICommandService);
    return commandService.onCommandExecuted((command) => {
      if (command.id === PasteCommand.id || command.id === SheetPasteShortKeyCommandName) {
        callback();
      }
    });
  }
};
FHooks.extend(FHooksSheetsMixin);

// ../packages/ui/src/facade/f-enum.ts
var FUIEnumMixin = class extends FEnum {
  get BuiltInUIPart() {
    return BuiltInUIPart;
  }
  get KeyCode() {
    return KeyCode;
  }
};
FEnum.extend(FUIEnumMixin);

// ../packages/sheets-ui/src/facade/f-univer.ts
var FUniverSheetsUIMixin = class extends FUniver {
  // eslint-disable-next-line max-lines-per-function
  _initSheetUIEvent(injector) {
    const commandService = injector.get(ICommandService);
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeSheetEditStart,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          if (commandInfo.id !== SetCellEditVisibleOperation.id) return;
          const target = this.getActiveSheet();
          if (!target) return;
          const { workbook, worksheet } = target;
          const editorBridgeService = injector.get(IEditorBridgeService);
          const params = commandInfo.params;
          const { visible, keycode, eventType } = params;
          const loc = editorBridgeService.getEditLocation();
          if (!visible) return;
          const eventParams = {
            row: loc.row,
            column: loc.column,
            eventType,
            keycode,
            workbook,
            worksheet
          };
          this.fireEvent(this.Event.BeforeSheetEditStart, eventParams);
          if (eventParams.cancel) {
            throw new CanceledError();
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeSheetEditEnd,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          if (commandInfo.id !== SetCellEditVisibleOperation.id) return;
          const target = this.getActiveSheet();
          if (!target) return;
          const { workbook, worksheet } = target;
          const editorBridgeService = injector.get(IEditorBridgeService);
          const univerInstanceService = injector.get(IUniverInstanceService);
          const params = commandInfo.params;
          const { visible, keycode, eventType } = params;
          const loc = editorBridgeService.getEditLocation();
          if (visible) return;
          const eventParams = {
            row: loc.row,
            column: loc.column,
            eventType,
            keycode,
            workbook,
            worksheet,
            value: RichTextValue.create(univerInstanceService.getUnit(DOCS_NORMAL_EDITOR_UNIT_ID_KEY).getSnapshot()),
            isConfirm: keycode !== 27 /* ESC */
          };
          this.fireEvent(this.Event.BeforeSheetEditEnd, eventParams);
          if (eventParams.cancel) {
            throw new CanceledError();
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.SheetEditStarted,
        () => commandService.onCommandExecuted((commandInfo) => {
          if (commandInfo.id !== SetCellEditVisibleOperation.id) return;
          const params = commandInfo.params;
          const target = this.getSheetCommandTarget(params);
          if (!target) return;
          const { workbook, worksheet } = target;
          const { visible, keycode, eventType } = params;
          if (!visible) return;
          const loc = injector.get(IEditorBridgeService).getEditLocation();
          const eventParams = {
            row: loc.row,
            column: loc.column,
            eventType,
            keycode,
            workbook,
            worksheet
          };
          this.fireEvent(this.Event.SheetEditStarted, eventParams);
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.SheetEditEnded,
        () => commandService.onCommandExecuted((commandInfo) => {
          if (commandInfo.id !== SetCellEditVisibleOperation.id) return;
          const params = commandInfo.params;
          const target = this.getSheetCommandTarget(params);
          if (!target) return;
          const { workbook, worksheet } = target;
          const { visible, keycode, eventType } = params;
          if (visible) return;
          const loc = injector.get(IEditorBridgeService).getEditLocation();
          const eventParams = {
            row: loc.row,
            column: loc.column,
            eventType,
            keycode,
            workbook,
            worksheet,
            isConfirm: keycode !== 27 /* ESC */
          };
          this.fireEvent(this.Event.SheetEditEnded, eventParams);
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.SheetEditChanging,
        () => commandService.onCommandExecuted((commandInfo) => {
          if (commandInfo.id !== RichTextEditingMutation.id) return;
          const params = commandInfo.params;
          const target = this.getActiveSheet();
          if (!target) return;
          const { workbook, worksheet } = target;
          const editorBridgeService = injector.get(IEditorBridgeService);
          const univerInstanceService = injector.get(IUniverInstanceService);
          if (!editorBridgeService.isVisible().visible) return;
          const { unitId } = params;
          if (unitId === DOCS_NORMAL_EDITOR_UNIT_ID_KEY) {
            const { row, column } = editorBridgeService.getEditLocation();
            const eventParams = {
              workbook,
              worksheet,
              row,
              column,
              value: RichTextValue.create(univerInstanceService.getUnit(DOCS_NORMAL_EDITOR_UNIT_ID_KEY).getSnapshot())
            };
            this.fireEvent(this.Event.SheetEditChanging, eventParams);
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeSheetZoomChange,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          if (commandInfo.id !== SetZoomRatioCommand.id) return;
          const params = commandInfo.params;
          const target = this.getSheetCommandTarget(params);
          if (!target) return;
          const { workbook, worksheet } = target;
          const { zoomRatio: zoom } = params;
          const eventParams = {
            workbook,
            worksheet,
            zoom
          };
          this.fireEvent(this.Event.BeforeSheetZoomChange, eventParams);
          if (eventParams.cancel) {
            throw new CanceledError();
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.SheetZoomChanged,
        () => commandService.onCommandExecuted((commandInfo) => {
          if (commandInfo.id !== SetZoomRatioCommand.id) return;
          const params = commandInfo.params;
          const target = this.getSheetCommandTarget(params);
          if (!target) return;
          const { workbook, worksheet } = target;
          const { zoomRatio: zoom } = params;
          const eventParams = {
            workbook,
            worksheet,
            zoom
          };
          this.fireEvent(this.Event.SheetZoomChanged, eventParams);
        })
      )
    );
  }
  // eslint-disable-next-line max-lines-per-function
  _initObserverListener(injector) {
    const renderManagerService = injector.get(IRenderManagerService);
    const lifeCycleService = injector.get(LifecycleService);
    const lifecycle$Disposable = new DisposableCollection();
    this.disposeWithMe(lifeCycleService.lifecycle$.subscribe((lifecycle) => {
      if (lifecycle !== 2 /* Rendered */) return;
      const hoverManagerService = injector.get(HoverManagerService);
      const dragManagerService = injector.get(DragManagerService);
      if (!hoverManagerService) return;
      lifecycle$Disposable.dispose();
      lifecycle$Disposable.add(
        this.registerEventHandler(
          this.Event.CellClicked,
          () => {
            var _a;
            return (_a = hoverManagerService.currentClickedCell$) == null ? void 0 : _a.pipe(filter((cell) => !!cell)).subscribe((cell) => {
              const { unitId, subUnitId, row, col: column } = cell.location;
              const target = this.getSheetCommandTarget({ unitId, subUnitId });
              if (!target) return;
              const { workbook, worksheet } = target;
              const eventParams = {
                workbook,
                worksheet,
                row,
                column
              };
              this.fireEvent(this.Event.CellClicked, eventParams);
            });
          }
        )
      );
      lifecycle$Disposable.add(
        this.registerEventHandler(
          this.Event.CellHover,
          () => {
            var _a;
            return (_a = hoverManagerService.currentRichText$) == null ? void 0 : _a.pipe(filter((cell) => !!cell)).subscribe((cell) => {
              const { unitId, subUnitId, row, col: column } = cell;
              const target = this.getSheetCommandTarget({ unitId, subUnitId });
              if (!target) return;
              const { workbook, worksheet } = target;
              const eventParams = {
                workbook,
                worksheet,
                row,
                column
              };
              this.fireEvent(this.Event.CellHover, eventParams);
            });
          }
        )
      );
      lifecycle$Disposable.add(
        this.registerEventHandler(
          this.Event.CellPointerDown,
          () => {
            var _a;
            return (_a = hoverManagerService.currentPointerDownCell$) == null ? void 0 : _a.pipe(filter((cell) => !!cell)).subscribe((cell) => {
              const { unitId, subUnitId, row, col: column } = cell;
              const target = this.getSheetCommandTarget({ unitId, subUnitId });
              if (!target) return;
              const { workbook, worksheet } = target;
              const eventParams = {
                workbook,
                worksheet,
                row,
                column
              };
              this.fireEvent(this.Event.CellPointerDown, eventParams);
            });
          }
        )
      );
      lifecycle$Disposable.add(
        this.registerEventHandler(
          this.Event.CellPointerUp,
          () => {
            var _a;
            return (_a = hoverManagerService.currentPointerUpCell$) == null ? void 0 : _a.pipe(filter((cell) => !!cell)).subscribe((cell) => {
              const { unitId, subUnitId, row, col: column } = cell;
              const target = this.getSheetCommandTarget({ unitId, subUnitId });
              if (!target) return;
              const { workbook, worksheet } = target;
              const eventParams = {
                workbook,
                worksheet,
                row,
                column
              };
              this.fireEvent(this.Event.CellPointerUp, eventParams);
            });
          }
        )
      );
      lifecycle$Disposable.add(
        this.registerEventHandler(
          this.Event.CellPointerMove,
          () => {
            var _a;
            return (_a = hoverManagerService.currentCellPosWithEvent$) == null ? void 0 : _a.pipe(filter((cell) => !!cell)).subscribe((cell) => {
              const { unitId, subUnitId, row, col: column } = cell;
              const target = this.getSheetCommandTarget({ unitId, subUnitId });
              if (!target) return;
              const { workbook, worksheet } = target;
              const eventParams = {
                workbook,
                worksheet,
                row,
                column
              };
              this.fireEvent(this.Event.CellPointerMove, eventParams);
            });
          }
        )
      );
      lifecycle$Disposable.add(
        this.registerEventHandler(
          this.Event.DragOver,
          () => {
            var _a;
            return (_a = dragManagerService.currentCell$) == null ? void 0 : _a.pipe(filter((cell) => !!cell)).subscribe((cell) => {
              const { unitId, subUnitId, row, col: column } = cell.location;
              const target = this.getSheetCommandTarget({ unitId, subUnitId });
              if (!target) return;
              const { workbook, worksheet } = target;
              const eventParams = {
                workbook,
                worksheet,
                ...cell,
                row,
                column
              };
              this.fireEvent(this.Event.DragOver, eventParams);
            });
          }
        )
      );
      lifecycle$Disposable.add(
        this.registerEventHandler(
          this.Event.Drop,
          () => {
            var _a;
            return (_a = dragManagerService.endCell$) == null ? void 0 : _a.pipe(filter((cell) => !!cell)).subscribe((cell) => {
              const { unitId, subUnitId, row, col: column } = cell.location;
              const target = this.getSheetCommandTarget({ unitId, subUnitId });
              if (!target) return;
              const { workbook, worksheet } = target;
              const eventParams = {
                workbook,
                worksheet,
                ...cell,
                row,
                column
              };
              this.fireEvent(this.Event.Drop, eventParams);
            });
          }
        )
      );
      lifecycle$Disposable.add(
        this.registerEventHandler(
          this.Event.RowHeaderClick,
          () => {
            var _a;
            return (_a = hoverManagerService.currentRowHeaderClick$) == null ? void 0 : _a.pipe(filter((header) => !!header)).subscribe((header) => {
              const { unitId, subUnitId, index: row } = header;
              const target = this.getSheetCommandTarget({ unitId, subUnitId });
              if (!target) return;
              const { workbook, worksheet } = target;
              const eventParams = {
                workbook,
                worksheet,
                row
              };
              this.fireEvent(this.Event.RowHeaderClick, eventParams);
            });
          }
        )
      );
      lifecycle$Disposable.add(
        this.registerEventHandler(
          this.Event.RowHeaderPointerDown,
          () => {
            var _a;
            return (_a = hoverManagerService.currentRowHeaderPointerDown$) == null ? void 0 : _a.pipe(filter((header) => !!header)).subscribe((header) => {
              const { unitId, subUnitId, index: row } = header;
              const target = this.getSheetCommandTarget({ unitId, subUnitId });
              if (!target) return;
              const { workbook, worksheet } = target;
              const eventParams = {
                workbook,
                worksheet,
                row
              };
              this.fireEvent(this.Event.RowHeaderPointerDown, eventParams);
            });
          }
        )
      );
      lifecycle$Disposable.add(
        this.registerEventHandler(
          this.Event.RowHeaderPointerUp,
          () => {
            var _a;
            return (_a = hoverManagerService.currentRowHeaderPointerUp$) == null ? void 0 : _a.pipe(filter((header) => !!header)).subscribe((header) => {
              const { unitId, subUnitId, index: row } = header;
              const target = this.getSheetCommandTarget({ unitId, subUnitId });
              if (!target) return;
              const { workbook, worksheet } = target;
              const eventParams = {
                workbook,
                worksheet,
                row
              };
              this.fireEvent(this.Event.RowHeaderPointerUp, eventParams);
            });
          }
        )
      );
      lifecycle$Disposable.add(
        this.registerEventHandler(
          this.Event.RowHeaderHover,
          () => {
            var _a;
            return (_a = hoverManagerService.currentHoveredRowHeader$) == null ? void 0 : _a.pipe(filter((header) => !!header)).subscribe((header) => {
              const { unitId, subUnitId, index: row } = header;
              const target = this.getSheetCommandTarget({ unitId, subUnitId });
              if (!target) return;
              const { workbook, worksheet } = target;
              const eventParams = {
                workbook,
                worksheet,
                row
              };
              this.fireEvent(this.Event.RowHeaderHover, eventParams);
            });
          }
        )
      );
      lifecycle$Disposable.add(
        this.registerEventHandler(
          this.Event.ColumnHeaderClick,
          () => {
            var _a;
            return (_a = hoverManagerService.currentColHeaderClick$) == null ? void 0 : _a.pipe(filter((header) => !!header)).subscribe((header) => {
              const { unitId, subUnitId, index: column } = header;
              const target = this.getSheetCommandTarget({ unitId, subUnitId });
              if (!target) return;
              const { workbook, worksheet } = target;
              const eventParams = {
                workbook,
                worksheet,
                column
              };
              this.fireEvent(this.Event.ColumnHeaderClick, eventParams);
            });
          }
        )
      );
      lifecycle$Disposable.add(
        this.registerEventHandler(
          this.Event.ColumnHeaderPointerDown,
          () => {
            var _a;
            return (_a = hoverManagerService.currentColHeaderPointerDown$) == null ? void 0 : _a.pipe(filter((header) => !!header)).subscribe((header) => {
              const { unitId, subUnitId, index: column } = header;
              const target = this.getSheetCommandTarget({ unitId, subUnitId });
              if (!target) return;
              const { workbook, worksheet } = target;
              const eventParams = {
                workbook,
                worksheet,
                column
              };
              this.fireEvent(this.Event.ColumnHeaderPointerDown, eventParams);
            });
          }
        )
      );
      lifecycle$Disposable.add(
        this.registerEventHandler(
          this.Event.ColumnHeaderPointerUp,
          () => {
            var _a;
            return (_a = hoverManagerService.currentColHeaderPointerUp$) == null ? void 0 : _a.pipe(filter((header) => !!header)).subscribe((header) => {
              const { unitId, subUnitId, index: column } = header;
              const target = this.getSheetCommandTarget({ unitId, subUnitId });
              if (!target) return;
              const { workbook, worksheet } = target;
              const eventParams = {
                workbook,
                worksheet,
                column
              };
              this.fireEvent(this.Event.ColumnHeaderPointerUp, eventParams);
            });
          }
        )
      );
      lifecycle$Disposable.add(
        this.registerEventHandler(
          this.Event.ColumnHeaderHover,
          () => {
            var _a;
            return (_a = hoverManagerService.currentHoveredColHeader$) == null ? void 0 : _a.pipe(filter((header) => !!header)).subscribe((header) => {
              const { unitId, subUnitId, index: column } = header;
              const target = this.getSheetCommandTarget({ unitId, subUnitId });
              if (!target) return;
              const { workbook, worksheet } = target;
              const eventParams = {
                workbook,
                worksheet,
                column
              };
              this.fireEvent(this.Event.ColumnHeaderHover, eventParams);
            });
          }
        )
      );
      this.disposeWithMe(lifecycle$Disposable);
    }));
    let sheetRenderUnit;
    const combined$ = combineLatest([
      renderManagerService.created$,
      lifeCycleService.lifecycle$
    ]);
    const combined$Disposable = new DisposableCollection();
    this.disposeWithMe(combined$.subscribe(([created, lifecycle]) => {
      if (created.type === 2 /* UNIVER_SHEET */) {
        sheetRenderUnit = created;
      }
      if (lifecycle <= 2 /* Rendered */) return;
      if (!sheetRenderUnit) return;
      const workbook = this.getWorkbook(sheetRenderUnit.unitId);
      if (!workbook) return;
      combined$Disposable.dispose();
      const scrollManagerService = sheetRenderUnit.with(SheetScrollManagerService);
      const selectionService = sheetRenderUnit.with(SheetsSelectionsService);
      combined$Disposable.add(
        this.registerEventHandler(
          this.Event.Scroll,
          () => scrollManagerService.validViewportScrollInfo$.subscribe((params) => {
            if (!params) return;
            const eventParams = {
              workbook,
              worksheet: workbook.getActiveSheet(),
              ...params
            };
            this.fireEvent(this.Event.Scroll, eventParams);
          })
        )
      );
      combined$Disposable.add(
        this.registerEventHandler(
          this.Event.SelectionMoveStart,
          () => selectionService.selectionMoveStart$.subscribe((selections) => {
            var _a;
            const eventParams = {
              workbook,
              worksheet: workbook.getActiveSheet(),
              selections: (_a = selections == null ? void 0 : selections.map((s) => s.range)) != null ? _a : []
            };
            this.fireEvent(this.Event.SelectionMoveStart, eventParams);
          })
        )
      );
      combined$Disposable.add(
        this.registerEventHandler(
          this.Event.SelectionMoving,
          () => selectionService.selectionMoving$.subscribe((selections) => {
            var _a;
            const eventParams = {
              workbook,
              worksheet: workbook.getActiveSheet(),
              selections: (_a = selections == null ? void 0 : selections.map((s) => s.range)) != null ? _a : []
            };
            this.fireEvent(this.Event.SelectionMoving, eventParams);
          })
        )
      );
      combined$Disposable.add(
        this.registerEventHandler(
          this.Event.SelectionMoveEnd,
          () => selectionService.selectionMoveEnd$.subscribe((selections) => {
            var _a;
            const eventParams = {
              workbook,
              worksheet: workbook.getActiveSheet(),
              selections: (_a = selections == null ? void 0 : selections.map((s) => s.range)) != null ? _a : []
            };
            this.fireEvent(this.Event.SelectionMoveEnd, eventParams);
          })
        )
      );
      combined$Disposable.add(
        this.registerEventHandler(
          this.Event.SelectionChanged,
          () => selectionService.selectionChanged$.subscribe((selections) => {
            var _a;
            const eventParams = {
              workbook,
              worksheet: workbook.getActiveSheet(),
              selections: (_a = selections == null ? void 0 : selections.map((s) => s.range)) != null ? _a : []
            };
            this.fireEvent(this.Event.SelectionChanged, eventParams);
          })
        )
      );
      sheetRenderUnit = null;
      this.disposeWithMe(combined$Disposable);
    }));
  }
  /**
   * @ignore
   */
  // eslint-disable-next-line max-lines-per-function
  _initialize(injector) {
    this._initSheetUIEvent(injector);
    this._initObserverListener(injector);
    const commandService = injector.get(ICommandService);
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeClipboardChange,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          switch (commandInfo.id) {
            case CopyCommand.id:
            case CutCommand.id:
              this._beforeClipboardChange();
              break;
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.ClipboardChanged,
        () => commandService.onCommandExecuted((commandInfo) => {
          switch (commandInfo.id) {
            case CopyCommand.id:
            case CutCommand.id:
              this._clipboardChanged();
              break;
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.BeforeClipboardPaste,
        () => commandService.beforeCommandExecuted((commandInfo) => {
          switch (commandInfo.id) {
            case SheetPasteShortKeyCommand.id:
              this._beforeClipboardPaste(commandInfo.params);
              break;
            case PasteCommand.id:
              this._beforeClipboardPasteAsync();
              break;
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.ClipboardPasted,
        () => commandService.onCommandExecuted((commandInfo) => {
          switch (commandInfo.id) {
            case SheetPasteShortKeyCommand.id:
              this._clipboardPaste(commandInfo.params);
              break;
            case PasteCommand.id:
              this._clipboardPasteAsync();
              break;
          }
        })
      )
    );
    this.disposeWithMe(
      this.registerEventHandler(
        this.Event.SheetSkeletonChanged,
        () => commandService.onCommandExecuted((commandInfo) => {
          if (COMMAND_LISTENER_SKELETON_CHANGE.indexOf(commandInfo.id) > -1) {
            const sheet = this.getActiveSheet();
            if (!sheet) return;
            const ranges = getSkeletonChangedEffectedRange(commandInfo, sheet.worksheet.getMaxColumns()).map((range) => {
              var _a, _b;
              return (_b = (_a = this.getWorkbook(range.unitId)) == null ? void 0 : _a.getSheetBySheetId(range.subUnitId)) == null ? void 0 : _b.getRange(range.range);
            }).filter(Boolean);
            if (!ranges.length) return;
            const eventParams = {
              workbook: sheet.workbook,
              worksheet: sheet.worksheet,
              payload: commandInfo,
              skeleton: sheet.worksheet.getSkeleton(),
              effectedRanges: ranges
            };
            this.fireEvent(this.Event.SheetSkeletonChanged, eventParams);
          }
        })
      )
    );
  }
  _generateClipboardCopyParam() {
    const workbook = this.getActiveWorkbook();
    const worksheet = workbook == null ? void 0 : workbook.getActiveSheet();
    const range = workbook == null ? void 0 : workbook.getActiveRange();
    if (!workbook || !worksheet || !range) {
      return;
    }
    const clipboardService = this._injector.get(ISheetClipboardService);
    const content = clipboardService.generateCopyContent(workbook.getId(), worksheet.getSheetId(), range.getRange());
    if (!content) {
      return;
    }
    const { html, plain } = content;
    const eventParams = {
      workbook,
      worksheet,
      text: plain,
      html,
      fromSheet: worksheet,
      fromRange: range
    };
    return eventParams;
  }
  _beforeClipboardChange() {
    const eventParams = this._generateClipboardCopyParam();
    if (!eventParams) return;
    this.fireEvent(this.Event.BeforeClipboardChange, eventParams);
    if (eventParams.cancel) {
      throw new CanceledError();
    }
  }
  _clipboardChanged() {
    const eventParams = this._generateClipboardCopyParam();
    if (!eventParams) return;
    this.fireEvent(this.Event.ClipboardChanged, eventParams);
  }
  _generateClipboardPasteParam(params) {
    if (!params) {
      return;
    }
    const { htmlContent, textContent } = params;
    const workbook = this.getActiveWorkbook();
    const worksheet = workbook == null ? void 0 : workbook.getActiveSheet();
    if (!workbook || !worksheet) {
      return;
    }
    const eventParams = {
      workbook,
      worksheet,
      text: textContent,
      html: htmlContent
    };
    return eventParams;
  }
  async _generateClipboardPasteParamAsync() {
    const workbook = this.getActiveWorkbook();
    const worksheet = workbook == null ? void 0 : workbook.getActiveSheet();
    if (!workbook || !worksheet) {
      return;
    }
    const clipboardInterfaceService = this._injector.get(IClipboardInterfaceService);
    const clipboardItems = await clipboardInterfaceService.read();
    const item = clipboardItems[0];
    let eventParams;
    if (item) {
      const types = item.types;
      const text = types.indexOf(PLAIN_TEXT_CLIPBOARD_MIME_TYPE) !== -1 ? await item.getType(PLAIN_TEXT_CLIPBOARD_MIME_TYPE).then((blob) => blob && blob.text()) : "";
      const html = types.indexOf(HTML_CLIPBOARD_MIME_TYPE) !== -1 ? await item.getType(HTML_CLIPBOARD_MIME_TYPE).then((blob) => blob && blob.text()) : "";
      eventParams = {
        workbook,
        worksheet,
        text,
        html
      };
    }
    return eventParams;
  }
  _beforeClipboardPaste(params) {
    const eventParams = this._generateClipboardPasteParam(params);
    if (!eventParams) return;
    this.fireEvent(this.Event.BeforeClipboardPaste, eventParams);
    if (eventParams.cancel) {
      throw new CanceledError();
    }
  }
  _clipboardPaste(params) {
    const eventParams = this._generateClipboardPasteParam(params);
    if (!eventParams) return;
    this.fireEvent(this.Event.ClipboardPasted, eventParams);
    if (eventParams.cancel) {
      throw new CanceledError();
    }
  }
  async _beforeClipboardPasteAsync() {
    if (!supportClipboardAPI()) {
      const logService = this._injector.get(ILogService);
      logService.warn("[Facade]: The navigator object only supports the browser environment");
      return;
    }
    const eventParams = await this._generateClipboardPasteParamAsync();
    if (!eventParams) return;
    this.fireEvent(this.Event.BeforeClipboardPaste, eventParams);
    if (eventParams.cancel) {
      throw new CanceledError();
    }
  }
  async _clipboardPasteAsync() {
    if (!supportClipboardAPI()) {
      const logService = this._injector.get(ILogService);
      logService.warn("[Facade]: The navigator object only supports the browser environment");
      return;
    }
    const eventParams = await this._generateClipboardPasteParamAsync();
    if (!eventParams) return;
    this.fireEvent(this.Event.ClipboardPasted, eventParams);
    if (eventParams.cancel) {
      throw new CanceledError();
    }
  }
  registerSheetRowHeaderExtension(unitId, ...extensions) {
    const sheetComponent = this._getSheetRenderComponent(unitId, "__SpreadsheetRowHeader__" /* ROW */);
    const registerDisposable = sheetComponent.register(...extensions);
    return toDisposable(() => {
      registerDisposable.dispose();
      sheetComponent.makeDirty(true);
    });
  }
  registerSheetColumnHeaderExtension(unitId, ...extensions) {
    const sheetComponent = this._getSheetRenderComponent(unitId, "__SpreadsheetColumnHeader__" /* COLUMN */);
    const registerDisposable = sheetComponent.register(...extensions);
    return toDisposable(() => {
      registerDisposable.dispose();
      sheetComponent.makeDirty(true);
    });
  }
  registerSheetMainExtension(unitId, ...extensions) {
    const sheetComponent = this._getSheetRenderComponent(unitId, "__SpreadsheetRender__" /* MAIN */);
    const registerDisposable = sheetComponent.register(...extensions);
    return toDisposable(() => {
      registerDisposable.dispose();
      sheetComponent.makeDirty(true);
    });
  }
  registerCellCustomRender(customRender, effect = 1 /* Style */, priority) {
    return this._injector.get(SheetInterceptorService).intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
      effect,
      handler: (cell, pos, next) => {
        if (!cell) {
          return next(cell);
        }
        if (!cell.customRender && customRender) {
          cell.customRender = [...customRender];
        }
        return next(cell);
      },
      priority
    });
  }
  /**
   * Get sheet render component from render by unitId and view key.
   * @private
   * @param {string} unitId The unit id of the spreadsheet.
   * @param {SHEET_VIEW_KEY} viewKey The view key of the spreadsheet.
   * @returns {Nullable<RenderComponentType>} The render component.
   */
  _getSheetRenderComponent(unitId, viewKey) {
    const renderManagerService = this._injector.get(IRenderManagerService);
    const render = renderManagerService.getRenderUnitById(unitId);
    if (!render) {
      throw new Error(`Render Unit with unitId ${unitId} not found`);
    }
    const { components } = render;
    const renderComponent = components.get(viewKey);
    if (!renderComponent) {
      throw new Error("Render component not found");
    }
    return renderComponent;
  }
  pasteIntoSheet(htmlContent, textContent, files) {
    return this._commandService.executeCommand(SheetPasteShortKeyCommand.id, { htmlContent, textContent, files });
  }
  setProtectedRangeShadowStrategy(strategy) {
    const service = this._injector.get(SheetPermissionRenderManagerService);
    service.setProtectedRangeShadowStrategy(strategy);
  }
  getProtectedRangeShadowStrategy() {
    const service = this._injector.get(SheetPermissionRenderManagerService);
    return service.getProtectedRangeShadowStrategy();
  }
  getProtectedRangeShadowStrategy$() {
    const service = this._injector.get(SheetPermissionRenderManagerService);
    return service.getProtectedRangeShadowStrategy$();
  }
  setPermissionDialogVisible(visible) {
    const permissionService = this._injector.get(IPermissionService);
    permissionService.setShowComponents(visible);
  }
};
FUniver.extend(FUniverSheetsUIMixin);

// ../packages/sheets-ui/src/facade/f-workbook.ts
var FWorkbookSheetsUIMixin = class extends FWorkbook {
  openSiderbar(params) {
    this._logDeprecation("openSiderbar");
    const sideBarService = this._injector.get(ISidebarService);
    return sideBarService.open(params);
  }
  openDialog(dialog) {
    this._logDeprecation("openDialog");
    const dialogService = this._injector.get(IDialogService);
    const disposable = dialogService.open({
      ...dialog,
      onClose: () => {
        disposable.dispose();
      }
    });
    return disposable;
  }
  customizeColumnHeader(cfg) {
    const unitId = this._workbook.getUnitId();
    const sheetColumn = this._getSheetRenderComponent(unitId, "__SpreadsheetColumnHeader__" /* COLUMN */);
    sheetColumn.setCustomHeader(cfg);
  }
  customizeRowHeader(cfg) {
    const unitId = this._workbook.getUnitId();
    const sheetRow = this._getSheetRenderComponent(unitId, "__SpreadsheetRowHeader__" /* ROW */);
    sheetRow.setCustomHeader(cfg);
  }
  /**
   * Get sheet render component from render by unitId and view key.
   * @private
   * @param {string} unitId The unit id of the spreadsheet.
   * @param {SHEET_VIEW_KEY} viewKey The view key of the spreadsheet.
   * @returns {Nullable<RenderComponentType>} The render component.
   */
  _getSheetRenderComponent(unitId, viewKey) {
    const renderManagerService = this._injector.get(IRenderManagerService);
    const render = renderManagerService.getRenderUnitById(unitId);
    if (!render) {
      throw new Error(`Render Unit with unitId ${unitId} not found`);
    }
    const { components } = render;
    const renderComponent = components.get(viewKey);
    if (!renderComponent) {
      throw new Error("Render component not found");
    }
    return renderComponent;
  }
  _logDeprecation(name) {
    const logService = this._injector.get(ILogService);
    logService.warn("[FWorkbook]", `${name} is deprecated. Please use the function of the same name on "FUniver".`);
  }
  startEditing() {
    const commandService = this._injector.get(ICommandService);
    const editorBridgeService = this._injector.get(IEditorBridgeService);
    if (editorBridgeService.isVisible().visible) {
      return true;
    }
    return commandService.syncExecuteCommand(SetCellEditVisibleOperation.id, {
      eventType: 3 /* Dblclick */,
      unitId: this._workbook.getUnitId(),
      visible: true
    });
  }
  async endEditingAsync(save = true) {
    const commandService = this._injector.get(ICommandService);
    const editorBridgeService = this._injector.get(IEditorBridgeService);
    if (editorBridgeService.isVisible().visible) {
      commandService.syncExecuteCommand(SetCellEditVisibleOperation.id, {
        eventType: 4 /* Keyboard */,
        keycode: save ? 13 /* ENTER */ : 27 /* ESC */,
        visible: false,
        unitId: this._workbook.getUnitId()
      });
    }
    await awaitTime(0);
    return true;
  }
  abortEditingAsync() {
    return this.endEditingAsync(false);
  }
  isCellEditing() {
    const editorBridgeService = this._injector.get(IEditorBridgeService);
    return editorBridgeService.isVisible().visible;
  }
  getScrollStateBySheetId(sheetId) {
    const unitId = this._workbook.getUnitId();
    const renderManagerService = this._injector.get(IRenderManagerService);
    const render = renderManagerService.getRenderUnitById(unitId);
    if (!render) return null;
    const scm = render.with(SheetScrollManagerService);
    return scm.getScrollStateByParam({ unitId, sheetId });
  }
  disableSelection() {
    const unitId = this._workbook.getUnitId();
    const renderManagerService = this._injector.get(IRenderManagerService);
    const render = renderManagerService.getRenderUnitById(unitId);
    if (render) {
      render.with(ISheetSelectionRenderService).disableSelection();
    }
    return this;
  }
  enableSelection() {
    const unitId = this._workbook.getUnitId();
    const renderManagerService = this._injector.get(IRenderManagerService);
    const render = renderManagerService.getRenderUnitById(unitId);
    if (render) {
      render.with(ISheetSelectionRenderService).enableSelection();
    }
    return this;
  }
  transparentSelection() {
    const unitId = this._workbook.getUnitId();
    const renderManagerService = this._injector.get(IRenderManagerService);
    const render = renderManagerService.getRenderUnitById(unitId);
    if (render) {
      render.with(ISheetSelectionRenderService).transparentSelection();
    }
    return this;
  }
  showSelection() {
    const unitId = this._workbook.getUnitId();
    const renderManagerService = this._injector.get(IRenderManagerService);
    const render = renderManagerService.getRenderUnitById(unitId);
    if (render) {
      render.with(ISheetSelectionRenderService).showSelection();
    }
    return this;
  }
};
FWorkbook.extend(FWorkbookSheetsUIMixin);

// ../packages/sheets-ui/src/facade/f-worksheet.ts
var FWorksheetUIMixin = class extends FWorksheet {
  refreshCanvas() {
    const renderManagerService = this._injector.get(IRenderManagerService);
    const unitId = this._fWorkbook.id;
    const render = renderManagerService.getRenderUnitById(unitId);
    if (!render) {
      throw new Error(`Render Unit with unitId ${unitId} not found`);
    }
    render.with(SheetSkeletonManagerService).reCalculate();
    render.components.forEach((component) => {
      var _a;
      (_a = component.makeDirty) == null ? void 0 : _a.call(component);
    });
    return this;
  }
  highlightRanges(ranges, style, primary) {
    const markSelectionService = this._injector.get(IMarkSelectionService);
    const ids = [];
    for (const range of ranges) {
      const iRange = range.getRange();
      const id = markSelectionService.addShapeWithNoFresh({ range: iRange, style, primary });
      if (id) {
        ids.push(id);
      }
    }
    markSelectionService.refreshShapes();
    if (ids.length === 0) {
      throw new Error("Failed to highlight current range");
    }
    return toDisposable(() => {
      ids.forEach((id) => {
        markSelectionService.removeShape(id);
      });
    });
  }
  zoom(zoomRatio) {
    const commandService = this._injector.get(ICommandService);
    const _zoomRatio = Math.min(Math.max(zoomRatio, 0.1), 4);
    commandService.executeCommand(SetZoomRatioCommand.id, {
      unitId: this._workbook.getUnitId(),
      subUnitId: this._worksheet.getSheetId(),
      zoomRatio: _zoomRatio
    });
    return this;
  }
  getZoom() {
    return this._worksheet.getZoomRatio();
  }
  getVisibleRange() {
    const unitId = this._workbook.getUnitId();
    const renderManagerService = this._injector.get(IRenderManagerService);
    const render = renderManagerService.getRenderUnitById(unitId);
    if (!render) return null;
    const skm = render.with(SheetSkeletonManagerService);
    const sk = skm.getCurrentSkeleton();
    if (!sk) return null;
    return sk.getVisibleRangeByViewport("viewMain" /* VIEW_MAIN */);
  }
  getVisibleRangesOfAllViewports() {
    const unitId = this._workbook.getUnitId();
    const renderManagerService = this._injector.get(IRenderManagerService);
    const render = renderManagerService.getRenderUnitById(unitId);
    if (!render) return null;
    const skm = render.with(SheetSkeletonManagerService);
    const sk = skm.getCurrentSkeleton();
    if (!sk) return null;
    return sk.getVisibleRanges();
  }
  scrollToCell(row, column, duration) {
    const unitId = this._workbook.getUnitId();
    const renderManagerService = this._injector.get(IRenderManagerService);
    const render = renderManagerService.getRenderUnitById(unitId);
    if (render) {
      const scrollRenderController = render == null ? void 0 : render.with(SheetsScrollRenderController);
      scrollRenderController.scrollToCell(row, column, duration);
    }
    return this;
  }
  getScrollState() {
    const emptyScrollState = {
      offsetX: 0,
      offsetY: 0,
      sheetViewStartColumn: 0,
      sheetViewStartRow: 0
    };
    const unitId = this._workbook.getUnitId();
    const sheetId = this._worksheet.getSheetId();
    const renderManagerService = this._injector.get(IRenderManagerService);
    const render = renderManagerService.getRenderUnitById(unitId);
    if (!render) return emptyScrollState;
    const sheetScrollManagerService = render.with(SheetScrollManagerService);
    const scrollState = sheetScrollManagerService.getScrollStateByParam({ unitId, sheetId });
    return scrollState || emptyScrollState;
  }
  getSkeleton() {
    var _a;
    const service = (_a = this._injector.get(IRenderManagerService).getRenderUnitById(this._workbook.getUnitId())) == null ? void 0 : _a.with(SheetSkeletonManagerService);
    return service == null ? void 0 : service.getSkeleton(this._worksheet.getSheetId());
  }
  autoResizeColumn(columnPosition) {
    return this.autoResizeColumns(columnPosition, 1);
  }
  autoResizeColumns(startColumn, numColumns) {
    const unitId = this._workbook.getUnitId();
    const subUnitId = this._worksheet.getSheetId();
    const ranges = [
      {
        startColumn,
        endColumn: startColumn + numColumns - 1,
        startRow: 0,
        endRow: this._worksheet.getRowCount() - 1
      }
    ];
    this._commandService.syncExecuteCommand(SetWorksheetColAutoWidthCommand.id, {
      unitId,
      subUnitId,
      ranges
    });
    return this;
  }
  setColumnAutoWidth(columnPosition, numColumn) {
    return this.autoResizeColumns(columnPosition, numColumn);
  }
  autoResizeRows(startRow, numRows) {
    const unitId = this._workbook.getUnitId();
    const subUnitId = this._worksheet.getSheetId();
    const ranges = [
      {
        startRow,
        endRow: startRow + numRows - 1,
        startColumn: 0,
        endColumn: this._worksheet.getColumnCount() - 1
      }
    ];
    this._commandService.syncExecuteCommand(SetWorksheetRowIsAutoHeightCommand.id, {
      unitId,
      subUnitId,
      ranges
    });
    return this;
  }
  customizeColumnHeader(cfg) {
    var _a, _b;
    const unitId = this._workbook.getUnitId();
    const subUnitId = this._worksheet.getSheetId();
    const renderManagerService = this._injector.get(IRenderManagerService);
    const render = renderManagerService.getRenderUnitById(unitId);
    if (render && ((_a = cfg.headerStyle) == null ? void 0 : _a.size)) {
      const skm = render.with(SheetSkeletonManagerService);
      skm.setColumnHeaderSize(render, subUnitId, (_b = cfg.headerStyle) == null ? void 0 : _b.size);
    }
    const sheetColumn = this._getSheetRenderComponent(unitId, "__SpreadsheetColumnHeader__" /* COLUMN */);
    sheetColumn.setCustomHeader(cfg, subUnitId);
  }
  customizeRowHeader(cfg) {
    var _a, _b;
    const unitId = this._workbook.getUnitId();
    const subUnitId = this._worksheet.getSheetId();
    const renderManagerService = this._injector.get(IRenderManagerService);
    const render = renderManagerService.getRenderUnitById(unitId);
    if (render && ((_a = cfg.headerStyle) == null ? void 0 : _a.size)) {
      const skm = render.with(SheetSkeletonManagerService);
      skm.setRowHeaderSize(render, subUnitId, (_b = cfg.headerStyle) == null ? void 0 : _b.size);
    }
    const sheetRow = this._getSheetRenderComponent(unitId, "__SpreadsheetRowHeader__" /* ROW */);
    sheetRow.setCustomHeader(cfg, subUnitId);
  }
  setColumnHeaderHeight(height) {
    const unitId = this._workbook.getUnitId();
    const subUnitId = this._worksheet.getSheetId();
    this._commandService.executeCommand(SetColumnHeaderHeightCommand.id, {
      unitId,
      subUnitId,
      size: height
    });
    return this;
  }
  setRowHeaderWidth(width) {
    const unitId = this._workbook.getUnitId();
    const subUnitId = this._worksheet.getSheetId();
    this._commandService.executeCommand(SetRowHeaderWidthCommand.id, {
      unitId,
      subUnitId,
      size: width
    });
    return this;
  }
  /**
   * Get sheet render component from render by unitId and view key.
   * @private
   * @param {string} unitId The unit id of the spreadsheet.
   * @param {SHEET_VIEW_KEY} viewKey The view key of the spreadsheet.
   * @returns {Nullable<RenderComponentType>} The render component.
   */
  _getSheetRenderComponent(unitId, viewKey) {
    const renderManagerService = this._injector.get(IRenderManagerService);
    const render = renderManagerService.getRenderUnitById(unitId);
    if (!render) {
      throw new Error(`Render Unit with unitId ${unitId} not found`);
    }
    const { components } = render;
    const renderComponent = components.get(viewKey);
    if (!renderComponent) {
      throw new Error("Render component not found");
    }
    return renderComponent;
  }
};
FWorksheet.extend(FWorksheetUIMixin);

// ../packages/sheets-ui/src/facade/f-event.ts
var FSheetsUIEventNameMixin = class extends FEventName {
  get BeforeClipboardChange() {
    return "BeforeClipboardChange";
  }
  get ClipboardChanged() {
    return "ClipboardChanged";
  }
  get BeforeClipboardPaste() {
    return "BeforeClipboardPaste";
  }
  get ClipboardPasted() {
    return "ClipboardPasted";
  }
  get BeforeSheetEditStart() {
    return "BeforeSheetEditStart";
  }
  get SheetEditStarted() {
    return "SheetEditStarted";
  }
  get SheetEditChanging() {
    return "SheetEditChanging";
  }
  get BeforeSheetEditEnd() {
    return "BeforeSheetEditEnd";
  }
  get SheetEditEnded() {
    return "SheetEditEnded";
  }
  get CellClicked() {
    return "CellClicked";
  }
  get CellHover() {
    return "CellHover";
  }
  get CellPointerDown() {
    return "CellPointerDown";
  }
  get CellPointerUp() {
    return "CellPointerUp";
  }
  get CellPointerMove() {
    return "CellPointerMove";
  }
  get DragOver() {
    return "DragOver";
  }
  get Drop() {
    return "Drop";
  }
  get Scroll() {
    return "Scroll";
  }
  get SelectionMoveStart() {
    return "SelectionMoveStart";
  }
  get SelectionChanged() {
    return "SelectionChanged";
  }
  get SelectionMoving() {
    return "SelectionMoving";
  }
  get SelectionMoveEnd() {
    return "SelectionMoveEnd";
  }
  get RowHeaderClick() {
    return "RowHeaderClick";
  }
  get RowHeaderPointerDown() {
    return "RowHeaderPointerDown";
  }
  get RowHeaderPointerUp() {
    return "RowHeaderPointerUp";
  }
  get RowHeaderHover() {
    return "RowHeaderHover";
  }
  get ColumnHeaderClick() {
    return "ColumnHeaderClick";
  }
  get ColumnHeaderPointerDown() {
    return "ColumnHeaderPointerDown";
  }
  get ColumnHeaderPointerUp() {
    return "ColumnHeaderPointerUp";
  }
  get ColumnHeaderHover() {
    return "ColumnHeaderHover";
  }
  get SheetSkeletonChanged() {
    return "SheetSkeletonChanged";
  }
  get BeforeSheetZoomChange() {
    return "BeforeSheetZoomChange";
  }
  get SheetZoomChanged() {
    return "SheetZoomChanged";
  }
};
FEventName.extend(FSheetsUIEventNameMixin);

// ../packages/sheets-ui/src/facade/f-enum.ts
var FSheetsUIEnumMixin = class extends FEnum {
  get SHEET_VIEWPORT_KEY() {
    return SHEET_VIEWPORT_KEY;
  }
};
FEnum.extend(FSheetsUIEnumMixin);

// ../packages/sheets-ui/src/facade/f-range.ts
var FRangeSheetsUIMixin = class extends FRange {
  getCell() {
    var _a;
    const renderManagerService = this._injector.get(IRenderManagerService);
    const logService = this._injector.get(ILogService);
    const unitId = this._workbook.getUnitId();
    const subUnitId = this._worksheet.getSheetId();
    const render = renderManagerService.getRenderUnitById(unitId);
    const skeleton = (_a = render == null ? void 0 : render.with(SheetSkeletonManagerService).getSkeletonParam(subUnitId)) == null ? void 0 : _a.skeleton;
    if (!skeleton) {
      logService.error("[Facade]: `FRange.getCell` can only be called in current worksheet");
      throw new Error("`FRange.getCell` can only be called in current worksheet");
    }
    return skeleton.getCellWithCoordByIndex(this._range.startRow, this._range.startColumn);
  }
  getCellRect() {
    const { startX: x, startY: y, endX: x2, endY: y2 } = this.getCell();
    const data = { x, y, width: x2 - x, height: y2 - y, top: y, left: x, bottom: y2, right: x2 };
    return { ...data, toJSON: () => JSON.stringify(data) };
  }
  generateHTML() {
    var _a;
    const clipboardService = this._injector.get(ISheetClipboardService);
    const copyContent = clipboardService.generateCopyContent(
      this._workbook.getUnitId(),
      this._worksheet.getSheetId(),
      this._range
    );
    return (_a = copyContent == null ? void 0 : copyContent.html) != null ? _a : "";
  }
  attachPopup(popup) {
    var _a, _b, _c;
    popup.direction = (_a = popup.direction) != null ? _a : "horizontal";
    popup.extraProps = (_b = popup.extraProps) != null ? _b : {};
    popup.offset = (_c = popup.offset) != null ? _c : [0, 0];
    const { key, disposableCollection } = transformComponentKey(popup, this._injector.get(ComponentManager));
    const sheetsPopupService = this._injector.get(SheetCanvasPopManagerService);
    const disposePopup = sheetsPopupService.attachPopupToCell(
      this._range.startRow,
      this._range.startColumn,
      { ...popup, componentKey: key },
      this.getUnitId(),
      this._worksheet.getSheetId()
    );
    if (disposePopup) {
      disposableCollection.add(disposePopup);
      return disposableCollection;
    }
    disposableCollection.dispose();
    return null;
  }
  attachAlertPopup(alert) {
    const cellAlertService = this._injector.get(CellAlertManagerService);
    const location = {
      workbook: this._workbook,
      worksheet: this._worksheet,
      row: this._range.startRow,
      col: this._range.startColumn,
      unitId: this.getUnitId(),
      subUnitId: this._worksheet.getSheetId()
    };
    cellAlertService.showAlert({
      ...alert,
      location
    });
    return {
      dispose: () => {
        cellAlertService.removeAlert(alert.key);
      }
    };
  }
  /**
   * attachRangePopup
   * @param popup
   * @returns {IDisposable} disposable
   * @example
   * ```typescript
   * let fWorksheet = univerAPI.getActiveWorkbook().getSheetByName('Sheet1');
   * if (!fWorksheet) return;
   * let range = fWorksheet.getRange(2, 2, 3, 3);
   * univerAPI.getActiveWorkbook().setActiveRange(range);
   * let disposable = range.attachRangePopup({
   *   componentKey: 'univer.sheet.single-dom-popup',
   *   extraProps: { alert: { type: 0, title: 'This is an Info', message: 'This is an info message' } },
   * });
   * ```
   */
  attachRangePopup(popup) {
    var _a, _b, _c;
    popup.direction = (_a = popup.direction) != null ? _a : "top-center";
    popup.extraProps = (_b = popup.extraProps) != null ? _b : {};
    popup.offset = (_c = popup.offset) != null ? _c : [0, 0];
    const { key, disposableCollection } = transformComponentKey(popup, this._injector.get(ComponentManager));
    const sheetsPopupService = this._injector.get(SheetCanvasPopManagerService);
    const disposePopup = sheetsPopupService.attachRangePopup(
      this._range,
      { ...popup, componentKey: key },
      this.getUnitId(),
      this._worksheet.getSheetId()
    );
    if (disposePopup) {
      disposableCollection.add(disposePopup);
      return disposableCollection;
    }
    disposableCollection.dispose();
    return null;
  }
  highlight(style, primary) {
    const markSelectionService = this._injector.get(IMarkSelectionService);
    const id = markSelectionService.addShape({ range: this._range, style, primary });
    if (!id) {
      throw new Error("Failed to highlight current range");
    }
    return toDisposable(() => {
      markSelectionService.removeShape(id);
    });
  }
  showDropdown(param) {
    const cellDropdownManagerService = this._injector.get(ISheetCellDropdownManagerService);
    return cellDropdownManagerService.showDropdown(param);
  }
};
FRange.extend(FRangeSheetsUIMixin);
function transformComponentKey(component, componentManager) {
  const { componentKey, isVue3 } = component;
  let key;
  const disposableCollection = new DisposableCollection();
  if (typeof componentKey === "string") {
    key = componentKey;
  } else {
    key = `External_${generateRandomId(6)}`;
    disposableCollection.add(componentManager.register(key, componentKey, { framework: isVue3 ? "vue3" : "react" }));
  }
  return {
    key,
    disposableCollection
  };
}

// ../packages/engine-formula/src/facade/f-formula.ts
var FFormula = class extends FBase {
  constructor(_commandService, _injector, _lexerTreeBuilder, _configService, _functionService, _definedNamesService, _superTableService) {
    super();
    __publicField(this, "_commandService", _commandService);
    __publicField(this, "_injector", _injector);
    __publicField(this, "_lexerTreeBuilder", _lexerTreeBuilder);
    __publicField(this, "_configService", _configService);
    __publicField(this, "_functionService", _functionService);
    __publicField(this, "_definedNamesService", _definedNamesService);
    __publicField(this, "_superTableService", _superTableService);
    this._initialize();
  }
  /**
   * @ignore
   */
  _initialize() {
  }
  /**
   * The tree builder for formula string.
   * @type {LexerTreeBuilder}
   */
  get lexerTreeBuilder() {
    return this._lexerTreeBuilder;
  }
  /**
   * Offsets the formula
   * @param {string} formulaString - The formula string to offset
   * @param {number} refOffsetX - The offset column
   * @param {number} refOffsetY - The offset row
   * @param {boolean} [ignoreAbsolute] - Whether to ignore the absolute reference
   * @returns {string} The offset formula string
   *
   * @example
   * ```ts
   * const formulaEngine = univerAPI.getFormula();
   * const result = formulaEngine.moveFormulaRefOffset('=SUM(A1,B2)', 1, 1);
   * console.log(result);
   * ```
   */
  moveFormulaRefOffset(formulaString, refOffsetX, refOffsetY, ignoreAbsolute) {
    return this._lexerTreeBuilder.moveFormulaRefOffset(formulaString, refOffsetX, refOffsetY, ignoreAbsolute);
  }
  /**
   * Resolves the formula string to a 'node' node
   * @param {string} formulaString - The formula string to resolve
   * @returns {Array<ISequenceNode | string>} The nodes of the formula string
   *
   * @example
   * ```ts
   * const formulaEngine = univerAPI.getFormula();
   * const nodes = formulaEngine.sequenceNodesBuilder('=SUM(A1,B2)');
   * console.log(nodes);
   * ```
   */
  sequenceNodesBuilder(formulaString) {
    return this._lexerTreeBuilder.sequenceNodesBuilder(formulaString) || [];
  }
  /**
   * Start the calculation of the formula.
   *
   * @example
   * ```ts
   * const formulaEngine = univerAPI.getFormula();
   * formulaEngine.executeCalculation();
   * ```
   */
  executeCalculation() {
    this._commandService.executeCommand(SetTriggerFormulaCalculationStartMutation.id, { commands: [], forceCalculation: true }, { onlyLocal: true });
  }
  /**
   * Stop the calculation of the formula.
   *
   * @example
   * ```ts
   * const formulaEngine = univerAPI.getFormula();
   * formulaEngine.stopCalculation();
   * ```
   */
  stopCalculation() {
    this._commandService.executeCommand(SetFormulaCalculationStopMutation.id, {});
  }
  /**
   * Listening calculation starts.
   * @param {Function} callback - The callback function to be called when the formula calculation starts.
   * @returns {IDisposable} The disposable instance.
   *
   * @example
   * ```ts
   * const formulaEngine = univerAPI.getFormula();
   * formulaEngine.calculationStart((forceCalculation) => {
   *   console.log('Calculation start', forceCalculation);
   * });
   * ```
   */
  calculationStart(callback) {
    return this._commandService.onCommandExecuted((command) => {
      if (command.id === SetFormulaCalculationStartMutation.id) {
        const params = command.params;
        callback(params.forceCalculation);
      }
    });
  }
  /**
   * Listening calculation ends.
   * @param {Function} callback - The callback function to be called when the formula calculation ends.
   * @returns {IDisposable} The disposable instance.
   *
   * @example
   * ```ts
   * const formulaEngine = univerAPI.getFormula();
   * formulaEngine.calculationEnd((functionsExecutedState) => {
   *   console.log('Calculation end', functionsExecutedState);
   * });
   * ```
   */
  calculationEnd(callback) {
    return this._commandService.onCommandExecuted((command) => {
      if (command.id !== SetFormulaCalculationNotificationMutation.id) {
        return;
      }
      const params = command.params;
      if (params.functionsExecutedState !== void 0) {
        callback(params.functionsExecutedState);
      }
    });
  }
  /**
   * @deprecated Use `onCalculationResultApplied` instead.
   */
  whenComputingCompleteAsync(timeout) {
    const gcss = this._injector.get(GlobalComputingStatusService);
    if (gcss.computingStatus) return Promise.resolve(true);
    return firstValueFrom(race(
      gcss.computingStatus$.pipe(filter((computing) => computing)),
      timer(timeout != null ? timeout : 3e4).pipe(map(() => false))
    ));
  }
  /**
   * @deprecated Use `onCalculationResultApplied` instead.
   */
  onCalculationEnd() {
    return new Promise((resolve, reject) => {
      const timer2 = setTimeout(() => {
        reject(new Error("Calculation end timeout"));
      }, 3e4);
      const disposable = this.calculationEnd(() => {
        clearTimeout(timer2);
        disposable.dispose();
        resolve();
      });
    });
  }
  /**
   * Listening calculation processing.
   * @param {Function} callback - The callback function to be called when the formula calculation is in progress.
   * @returns {IDisposable} The disposable instance.
   *
   * @example
   * ```ts
   * const formulaEngine = univerAPI.getFormula();
   * formulaEngine.calculationProcessing((stageInfo) => {
   *   console.log('Calculation processing', stageInfo);
   * });
   * ```
   */
  calculationProcessing(callback) {
    return this._commandService.onCommandExecuted((command) => {
      if (command.id !== SetFormulaCalculationNotificationMutation.id) {
        return;
      }
      const params = command.params;
      if (params.stageInfo !== void 0) {
        callback(params.stageInfo);
      }
    });
  }
  /**
   * When a formula contains a circular reference, set the maximum number of iterations for the formula calculation.
   * @param {number} maxIteration The maximum number of iterations. The default value is 1.
   *
   * @example
   * ```ts
   * // Set the maximum number of iterations for the formula calculation to 5.
   * // The default value is 1.
   * const formulaEngine = univerAPI.getFormula();
   * formulaEngine.setMaxIteration(5);
   * ```
   */
  setMaxIteration(maxIteration) {
    this._configService.setConfig(ENGINE_FORMULA_CYCLE_REFERENCE_COUNT, maxIteration);
  }
  /**
   * Execute a batch of formulas asynchronously and receive computed results.
   *
   * Each formula cell is represented as a string array:
   *   [fullFormula, ...subFormulas]
   *
   * Where:
   *   - fullFormula (index 0) is the complete formula expression written in the cell.
   *     Example: "=SUM(A1:A10) + SQRT(D7)".
   *
   *   - subFormulas (index 1+) are **optional decomposed expressions** extracted from
   *     the full formula. Each of them can be independently computed by the formula engine.
   *
   *     These sub-expressions can include:
   *       - Single-cell references:  "A2", "B2", "C5"
   *       - Range references:        "A1:A10"
   *       - Function calls:          "SQRT(D7)", "ABS(A2-B2)"
   *       - Any sub-formula that was parsed out of the original formula and can be
   *         evaluated on its own.
   *
   *     The batch execution engine may use these sub-formulas for dependency resolution,
   *     incremental computation, or performance optimizations.
   *
   * @param {IFormulaStringMap} formulas
   *        Nested structure (unit → sheet → row → column) describing formulas and
   *        their decomposed sub-expressions.
   *
   * @param {number} [timeout]
   *        Optional timeout in milliseconds. If no result is received within this
   *        period, the promise will be rejected.
   *
   * @returns {Promise<IFormulaExecuteResultMap>}
   *          A promise that resolves with the computed value map mirroring
   *          the input structure.
   *
   * @example
   * ```ts
   * const formulaEngine = univerAPI.getFormula();
   * const formulas = {
   *   Book1: {
   *     Sheet1: {
   *       2: {
   *         3: [
   *           // Full formula:
   *           "=SUM(A1:A10) + SQRT(D7)",
   *
   *           // Decomposed sub-formulas (each one can be evaluated independently):
   *           "SUM(A1:A10)",   // sub-formula 1
   *           "SQRT(D7)",      // sub-formula 2
   *           "A1:A10",        // range reference
   *           "D7",            // single-cell reference
   *         ],
   *       },
   *       4: {
   *         5: [
   *           "=A2 + B2 + SQRT(C5)",
   *           "A2",
   *           "B2",
   *           "SQRT(C5)",
   *         ],
   *       }
   *     },
   *   },
   * };
   *
   * const result = await formulaEngine.executeFormulas(formulas);
   * console.log(result);
   * ```
   */
  executeFormulas(formulas, timeout = 3e4) {
    return new Promise((resolve, reject) => {
      const disposable = this._commandService.onCommandExecuted((command) => {
        if (command.id !== SetFormulaStringBatchCalculationResultMutation.id) {
          return;
        }
        const params = command.params;
        clearTimeout(timer2);
        disposable.dispose();
        if (params.result != null) {
          resolve(params.result);
        } else {
          reject(new Error("Formula batch calculation returned no result"));
        }
      });
      const timer2 = setTimeout(() => {
        disposable.dispose();
        reject(new Error("Formula batch calculation timeout"));
      }, timeout);
      this._commandService.executeCommand(
        SetFormulaStringBatchCalculationMutation.id,
        { formulas },
        { onlyLocal: true }
      );
    });
  }
  /**
   * Retrieve all formula dependency trees that were produced during the latest
   * dependency-analysis run. This triggers a local dependency-calculation command
   * and returns the complete set of dependency trees once the calculation finishes.
   *
   * @param {number} [timeout]
   *        Optional timeout in milliseconds. If no result is received within this
   *        period, the promise will be rejected.
   *
   * @returns {Promise<IFormulaDependencyTreeJson[]>}
   *          A promise that resolves with the array of dependency trees.
   *
   * @example
   * ```ts
   * const formulaEngine = univerAPI.getFormula();
   *
   * // Fetch all dependency trees generated for the current workbook.
   * const trees = await formulaEngine.getAllDependencyTrees();
   * console.log('All dependency trees:', trees);
   * ```
   */
  getAllDependencyTrees(timeout = 3e4) {
    return new Promise((resolve, reject) => {
      const disposable = this._commandService.onCommandExecuted((command) => {
        if (command.id !== SetFormulaDependencyCalculationResultMutation.id) {
          return;
        }
        const params = command.params;
        clearTimeout(timer2);
        disposable.dispose();
        if (params.result != null) {
          resolve(params.result);
        } else {
          resolve([]);
        }
      });
      const timer2 = setTimeout(() => {
        disposable.dispose();
        reject(new Error("Formula dependency calculation timeout"));
      }, timeout);
      this._commandService.executeCommand(
        SetFormulaDependencyCalculationMutation.id,
        void 0,
        { onlyLocal: true }
      );
    });
  }
  /**
   * Retrieve the dependency tree of a specific cell. This triggers a local
   * dependency-calculation command for the given unit, sheet, and cell location,
   * and returns the computed dependency tree when the calculation is completed.
   *
   * @param param The target cell location:
   *   - `unitId`  The workbook ID.
   *   - `sheetId` The sheet ID.
   *   - `row`     The zero-based row index.
   *   - `column`  The zero-based column index.
   *
   * @param {number} [timeout]
   *        Optional timeout in milliseconds. If no result is received within this
   *        period, the promise will be rejected.
   *
   * @returns {Promise<IFormulaDependencyTreeFullJson | undefined>}
   *          A promise that resolves with the dependency tree or `undefined`
   *          if no tree exists for that cell.
   *
   * @example
   * ```ts
   * const formulaEngine = univerAPI.getFormula();
   *
   * // Query the dependency tree for cell B2 in a specific sheet.
   * const tree = await formulaEngine.getCellDependencyTree({
   *   unitId: 'workbook1',
   *   sheetId: 'sheet1',
   *   row: 1,
   *   column: 1,
   * });
   *
   * console.log('Cell dependency tree:', tree);
   * ```
   */
  getCellDependencyTree(param, timeout = 3e4) {
    return new Promise((resolve, reject) => {
      const disposable = this._commandService.onCommandExecuted((command) => {
        if (command.id !== SetCellFormulaDependencyCalculationResultMutation.id) {
          return;
        }
        const params = command.params;
        clearTimeout(timer2);
        disposable.dispose();
        resolve(params.result);
      });
      const timer2 = setTimeout(() => {
        disposable.dispose();
        reject(new Error("Cell dependency calculation timeout"));
      }, timeout);
      this._commandService.executeCommand(
        SetCellFormulaDependencyCalculationMutation.id,
        param,
        { onlyLocal: true }
      );
    });
  }
  /**
   * Retrieve the full dependency trees for all formulas that *depend on* the
   * specified ranges. This triggers a local dependency-calculation command and
   * resolves once the calculation completes.
   *
   * @param unitRanges An array of workbook/sheet ranges to query. Each range
   *   includes:
   *   - `unitId`  The workbook ID.
   *   - `sheetId` The sheet ID.
   *   - `range`   The row/column boundaries.
   *
   * @param {number} [timeout]
   *        Optional timeout in milliseconds. If no result is received within this
   *        period, the promise will be rejected.
   *
   * @returns {Promise<IFormulaDependencyTreeJson[]>}
   *          A promise that resolves with an array of `IFormulaDependencyTreeJson`
   *          representing formulas and their relationships within the dependency graph.
   *
   * @example
   * ```ts
   * const formulaEngine = univerAPI.getFormula();
   *
   * // Query all formulas that depend on A1:B10 in Sheet1.
   * const dependents = await formulaEngine.getRangeDependents([
   *   { unitId: 'workbook1', sheetId: 'sheet1', range: { startRow: 0, endRow: 9, startColumn: 0, endColumn: 1 } }
   * ]);
   *
   * console.log('Dependent formulas:', dependents);
   * ```
   */
  getRangeDependents(unitRanges, timeout = 3e4) {
    return new Promise((resolve, reject) => {
      const disposable = this._commandService.onCommandExecuted((command) => {
        if (command.id !== SetQueryFormulaDependencyResultMutation.id) {
          return;
        }
        const params = command.params;
        clearTimeout(timer2);
        disposable.dispose();
        if (params.result != null) {
          resolve(params.result);
        } else {
          resolve([]);
        }
      });
      const timer2 = setTimeout(() => {
        disposable.dispose();
        reject(new Error("Range dependents calculation timeout"));
      }, timeout);
      this._commandService.executeCommand(
        SetQueryFormulaDependencyMutation.id,
        { unitRanges },
        { onlyLocal: true }
      );
    });
  }
  /**
   * Retrieve the dependency trees of all formulas *inside* the specified ranges.
   * Unlike `getRangeDependents`, this API only returns formulas whose definitions
   * physically reside within the queried ranges.
   *
   * Internally this triggers the same dependency-calculation command but with
   * `isInRange = true`, and the promise resolves when the results are ready.
   *
   * @param unitRanges An array of workbook/sheet ranges defining the lookup
   *   boundaries:
   *   - `unitId`  The workbook ID.
   *   - `sheetId` The sheet ID.
   *   - `range`   The zero-based grid range.
   *
   * @param {number} [timeout]
   *        Optional timeout in milliseconds. If no result is received within this
   *        period, the promise will be rejected.
   *
   * @returns {Promise<IFormulaDependencyTreeJson[]>}
   *          A promise that resolves with an array of `IFormulaDependencyTreeJson`
   *          describing every formula found in the provided ranges along with
   *          their parent/child relationships.
   *
   * @example
   * ```ts
   * const formulaEngine = univerAPI.getFormula();
   *
   * // Query all formulas that lie within A1:D20 in Sheet1.
   * const formulasInRange = await formulaEngine.getInRangeFormulas([
   *   { unitId: 'workbook1', sheetId: 'sheet1', range: { startRow: 0, endRow: 19, startColumn: 0, endColumn: 3 } }
   * ]);
   *
   * console.log('Formulas inside range:', formulasInRange);
   * ```
   */
  getInRangeFormulas(unitRanges, timeout = 3e4) {
    return new Promise((resolve, reject) => {
      const disposable = this._commandService.onCommandExecuted((command) => {
        if (command.id !== SetQueryFormulaDependencyResultMutation.id) {
          return;
        }
        const params = command.params;
        clearTimeout(timer2);
        disposable.dispose();
        if (params.result != null) {
          resolve(params.result);
        } else {
          resolve([]);
        }
      });
      const timer2 = setTimeout(() => {
        disposable.dispose();
        reject(new Error("In-range formulas calculation timeout"));
      }, timeout);
      this._commandService.executeCommand(
        SetQueryFormulaDependencyMutation.id,
        { unitRanges, isInRange: true },
        { onlyLocal: true }
      );
    });
  }
  /**
   * Enable or disable emitting formula dependency trees after each formula calculation.
   *
   * When enabled, the formula engine will emit the dependency trees produced by
   * each completed formula calculation through the internal command system.
   * Consumers can obtain the result by listening for the corresponding
   * calculation-result command.
   *
   * When disabled, dependency trees will not be emitted.
   *
   * This option only controls whether dependency trees are exposed.
   * It does not affect formula calculation behavior.
   *
   * @param {boolean} value
   *        Whether to emit formula dependency trees after calculation.
   *        - `true`: Emit dependency trees after each calculation.
   *        - `false`: Do not emit dependency trees (default behavior).
   *
   * @example
   * ```ts
   * const formulaEngine = univerAPI.getFormula();
   *
   * // Enable dependency tree emission
   * formulaEngine.setFormulaReturnDependencyTree(true);
   *
   * // Listen for dependency trees produced by formula calculation
   * const trees = await new Promise<IFormulaDependencyTreeJson[]>((resolve, reject) => {
   *   const timer = setTimeout(() => {
   *     disposable.dispose();
   *     reject(new Error('Timeout waiting for formula dependency trees'));
   *   }, 30_000);
   *
   *   const disposable = commandService.onCommandExecuted((command) => {
   *     if (command.id !== SetFormulaDependencyCalculationResultMutation.id) {
   *       return;
   *     }
   *
   *     clearTimeout(timer);
   *     disposable.dispose();
   *
   *     const params = command.params as ISetFormulaDependencyCalculationResultMutation;
   *     resolve(params.result ?? []);
   *   });
   * });
   *
   * console.log('Dependency trees:', trees);
   * ```
   */
  setFormulaReturnDependencyTree(value) {
    this._configService.setConfig(ENGINE_FORMULA_RETURN_DEPENDENCY_TREE, value);
  }
  /**
   * Parse a formula string and return its **formula expression tree**.
   *
   * This API analyzes the syntactic structure of a formula and builds an
   * expression tree that reflects how the formula is composed (functions,
   * operators, ranges, and nested expressions), without performing calculation
   * or dependency evaluation.
   *
   * The returned tree is suitable for:
   * - Formula structure visualization
   * - Explaining complex formulas (e.g. LET / LAMBDA)
   * - Debugging or inspecting formula composition
   * - Building advanced formula tooling
   *
   * ---
   *
   * @example
   * ```ts
   * const formulaEngine = univerAPI.getFormula();
   *
   * const formula = '=LET(x,SUM(A1,B1,A1:B10),y,OFFSET(A1:B10,0,1),SUM(x,y)+x)+1';
   *
   * const exprTree = formulaEngine.getFormulaExpressTree(formula);
   *
   * console.log(exprTree);
   * ```
   *
   * Example output (simplified):
   *
   * ```json
   * {
   *   "value": "let(x,sum(A1,B1,A1:B10),y,offset(A1:B10,0,1),sum(x,y)+x)+1",
   *   "children": [
   *     {
   *       "value": "let(x,sum(A1,B1,A1:B10),y,offset(A1:B10,0,1),sum(x,y)+x)",
   *       "children": [
   *         {
   *           "value": "sum(A1,B1,A1:B10)",
   *           "children": [
   *             {
   *               "value": "A1:B10",
   *               "children": []
   *             }
   *           ]
   *         },
   *         {
   *           "value": "offset(A1:B10,0,1)",
   *           "children": [
   *             {
   *               "value": "A1:B10",
   *               "children": []
   *             }
   *           ]
   *         }
   *       ]
   *     }
   *   ]
   * }
   * ```
   *
   * @param formulaString The formula string to parse (with or without leading `=`)
   * @returns A formula expression tree describing the hierarchical structure of the formula
   */
  getFormulaExpressTree(formulaString, unitId) {
    return this._lexerTreeBuilder.getFormulaExprTree(formulaString, unitId, this._functionService.hasExecutor.bind(this._functionService), this._definedNamesService.getValueByName.bind(this._definedNamesService), this._superTableService.getTable.bind(this._superTableService));
  }
  /**
   * Retrieve **both**:
   * 1) the full dependency trees of all formulas that **depend on** the specified ranges, and
   * 2) the dependency trees of all formulas that **physically reside inside** the specified ranges.
   *
   * This is a convenience API that combines the behaviors of
   * `getRangeDependents` and `getInRangeFormulas` into a single call.
   *
   * Internally, it triggers a local dependency-calculation command once and
   * resolves when both result sets are available, avoiding duplicate
   * calculations and event listeners.
   *
   * @param unitRanges An array of workbook/sheet ranges to query. Each range
   *   includes:
   *   - `unitId`  The workbook ID.
   *   - `sheetId` The sheet ID.
   *   - `range`   The zero-based row/column boundaries.
   *
   * @param {number} [timeout]
   *        Optional timeout in milliseconds. If the dependency calculation does
   *        not complete within this period, the promise will be rejected.
   *
   * @returns {Promise<IFormulaDependentsAndInRangeResults>}
   *          A promise that resolves with an object containing:
   *          - `dependents`: Dependency trees of all formulas that depend on the
   *            specified ranges (upstream consumers).
   *          - `inRanges`: Dependency trees of all formulas whose definitions
   *            are located inside the specified ranges.
   *
   * @example
   * ```ts
   * const formulaEngine = univerAPI.getFormula();
   *
   * const result = await formulaEngine.getRangeDependentsAndInRangeFormulas([
   *   {
   *     unitId: 'workbook1',
   *     sheetId: 'sheet1',
   *     range: { startRow: 0, endRow: 9, startColumn: 0, endColumn: 1 },
   *   },
   * ]);
   *
   * console.log('Dependent formulas:', result.dependents);
   * console.log('Formulas inside range:', result.inRanges);
   * ```
   */
  getRangeDependentsAndInRangeFormulas(unitRanges, timeout = 3e4) {
    return new Promise((resolve, reject) => {
      const disposable = this._commandService.onCommandExecuted((command) => {
        if (command.id !== SetQueryFormulaDependencyAllResultMutation.id) {
          return;
        }
        const params = command.params;
        clearTimeout(timer2);
        disposable.dispose();
        if (params.result != null) {
          resolve(params.result);
        } else {
          resolve({ dependents: [], inRanges: [] });
        }
      });
      const timer2 = setTimeout(() => {
        disposable.dispose();
        reject(new Error("Range dependents calculation timeout"));
      }, timeout);
      this._commandService.executeCommand(
        SetQueryFormulaDependencyAllMutation.id,
        { unitRanges },
        { onlyLocal: true }
      );
    });
  }
};
FFormula = __decorateClass([
  __decorateParam(0, Inject(ICommandService)),
  __decorateParam(1, Inject(Injector)),
  __decorateParam(2, Inject(LexerTreeBuilder)),
  __decorateParam(3, IConfigService),
  __decorateParam(4, IFunctionService),
  __decorateParam(5, IDefinedNamesService),
  __decorateParam(6, ISuperTableService)
], FFormula);

// ../packages/engine-formula/src/facade/f-univer.ts
var FUniverEngineFormulaMixin = class extends FUniver {
  getFormula() {
    return this._injector.createInstance(FFormula);
  }
};
FUniver.extend(FUniverEngineFormulaMixin);

// ../packages/sheets-formula/src/facade/f-univer.ts
var FUniverSheetsFormulaMixin = class extends FUniver {
  /**
   * Initialize the FUniver instance.
   * @ignore
   */
  _initialize() {
    this._debouncedFormulaCalculation = debounce_default(() => {
      this._commandService.executeCommand(
        SetTriggerFormulaCalculationStartMutation.id,
        {
          commands: [],
          forceCalculation: true
        },
        {
          onlyLocal: true
        }
      );
    }, 10);
  }
  registerFunction(config) {
    let registerFunctionService = this._injector.get(IRegisterFunctionService);
    if (!registerFunctionService) {
      this._injector.add([IRegisterFunctionService, { useClass: RegisterFunctionService }]);
      registerFunctionService = this._injector.get(IRegisterFunctionService);
    }
    const functionsDisposable = registerFunctionService.registerFunctions(config);
    this._debouncedFormulaCalculation();
    return functionsDisposable;
  }
};
FUniver.extend(FUniverSheetsFormulaMixin);

// ../packages/sheets-formula/src/facade/f-formula.ts
var FFormulaSheetsMixin = class extends FFormula {
  /**
   * Initialize the FUniver instance.
   * @ignore
   */
  _initialize() {
    this._debouncedFormulaCalculation = debounce_default(() => {
      this._commandService.executeCommand(
        SetTriggerFormulaCalculationStartMutation.id,
        {
          commands: [],
          forceCalculation: true
        },
        {
          onlyLocal: true
        }
      );
    }, 10);
  }
  setInitialFormulaComputing(calculationMode) {
    const lifecycleService = this._injector.get(LifecycleService);
    const lifecycleStage = lifecycleService.stage;
    const logService = this._injector.get(ILogService);
    const configService = this._injector.get(IConfigService);
    if (lifecycleStage > 0 /* Starting */) {
      logService.warn("[FFormula]", "CalculationMode is called after the Starting lifecycle and will take effect the next time the Univer Sheet is constructed. If you want it to take effect when the Univer Sheet is initialized this time, consider calling it before the Ready lifecycle or using configuration.");
    }
    const config = configService.getConfig(PLUGIN_CONFIG_KEY_BASE);
    if (!config) {
      configService.setConfig(PLUGIN_CONFIG_KEY_BASE, { initialFormulaComputing: calculationMode });
      return;
    }
    config.initialFormulaComputing = calculationMode;
  }
  registerFunction(name, func, options) {
    var _a;
    let registerFunctionService = this._injector.get(IRegisterFunctionService);
    if (!registerFunctionService) {
      this._injector.add([IRegisterFunctionService, { useClass: RegisterFunctionService }]);
      registerFunctionService = this._injector.get(IRegisterFunctionService);
    }
    const params = {
      name,
      func,
      description: typeof options === "string" ? options : (_a = options == null ? void 0 : options.description) != null ? _a : "",
      locales: typeof options === "object" ? options.locales : void 0
    };
    const functionsDisposable = registerFunctionService.registerFunction(params);
    this._debouncedFormulaCalculation();
    return functionsDisposable;
  }
  registerAsyncFunction(name, func, options) {
    var _a;
    let registerFunctionService = this._injector.get(IRegisterFunctionService);
    if (!registerFunctionService) {
      this._injector.add([IRegisterFunctionService, { useClass: RegisterFunctionService }]);
      registerFunctionService = this._injector.get(IRegisterFunctionService);
    }
    const params = {
      name,
      func,
      description: typeof options === "string" ? options : (_a = options == null ? void 0 : options.description) != null ? _a : "",
      locales: typeof options === "object" ? options.locales : void 0
    };
    const functionsDisposable = registerFunctionService.registerAsyncFunction(params);
    this._debouncedFormulaCalculation();
    return functionsDisposable;
  }
  calculationResultApplied(callback) {
    const subscription = this._injector.get(FormulaCalculationSessionService).resultApplied$.subscribe((result) => {
      requestIdleCallback(() => {
        callback(result);
      });
    });
    return {
      dispose: () => subscription.unsubscribe()
    };
  }
  onCalculationResultApplied(timeout) {
    return this._injector.get(FormulaCalculationSessionService).waitForLatestApplied(timeout);
  }
};
FFormula.extend(FFormulaSheetsMixin);

// ../packages/sheets-formula/src/facade/f-enum.ts
var FSheetsFormulaEnumMixin = class extends FEnum {
  get CalculationMode() {
    return CalculationMode;
  }
};
FEnum.extend(FSheetsFormulaEnumMixin);

// ../packages/sheets-formula/src/facade/f-workbook.ts
var FWorkbookEngineFormulaMixin = class extends FWorkbook {
  getAllFormulaError() {
    const errors = [];
    const workbook = this._workbook;
    const unitId = workbook.getUnitId();
    const worksheets = workbook.getSheets();
    const arrayFormula = this._injector.get(FormulaDataModel).getArrayFormulaCellData();
    worksheets.forEach((worksheet) => {
      var _a;
      const sheetName = worksheet.getName();
      const sheetId = worksheet.getSheetId();
      const cellMatrix = worksheet.getCellMatrix();
      const arrayFormulaSheet = ((_a = arrayFormula == null ? void 0 : arrayFormula[unitId]) == null ? void 0 : _a[sheetId]) || {};
      cellMatrix.forValue((row, column, cell) => {
        var _a2;
        if (!cell) return;
        const arrayFormulaCellData = (_a2 = arrayFormulaSheet == null ? void 0 : arrayFormulaSheet[row]) == null ? void 0 : _a2[column];
        const errorType = extractFormulaError(cell, !!arrayFormulaCellData);
        if (errorType) {
          errors.push({
            sheetName,
            row,
            column,
            formula: cell.f || "",
            errorType
          });
        }
      });
    });
    return errors;
  }
};
FWorkbook.extend(FWorkbookEngineFormulaMixin);

// ../packages/sheets-formula/src/facade/f-range.ts
var FRangeEngineFormulaMixin = class extends FRange {
  getFormulaError() {
    var _a, _b;
    const errors = [];
    const unitId = this._workbook.getUnitId();
    const sheetId = this._worksheet.getSheetId();
    const sheetName = this._worksheet.getName();
    const worksheet = this._workbook.getSheetBySheetId(sheetId);
    if (!worksheet) return errors;
    const arrayFormula = this._injector.get(FormulaDataModel).getArrayFormulaCellData();
    const arrayFormulaSheet = ((_a = arrayFormula == null ? void 0 : arrayFormula[unitId]) == null ? void 0 : _a[sheetId]) || {};
    const cellMatrix = worksheet.getCellMatrix();
    const { startRow, endRow, startColumn, endColumn } = this._range;
    for (let row = startRow; row <= endRow; row++) {
      for (let column = startColumn; column <= endColumn; column++) {
        const cell = cellMatrix.getValue(row, column);
        if (!cell) continue;
        const arrayFormulaCellData = (_b = arrayFormulaSheet == null ? void 0 : arrayFormulaSheet[row]) == null ? void 0 : _b[column];
        const errorType = extractFormulaError(cell, !!arrayFormulaCellData);
        if (errorType) {
          errors.push({
            sheetName,
            row,
            column,
            formula: cell.f || "",
            errorType
          });
        }
      }
    }
    return errors;
  }
};
FRange.extend(FRangeEngineFormulaMixin);

// ../packages/sheets-numfmt/src/facade/f-range.ts
var FRangeSheetsNumfmtMixin = class extends FRange {
  setNumberFormat(pattern) {
    const values = [];
    const { startColumn, startRow, endColumn, endRow } = this._range;
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startColumn; col <= endColumn; col++) {
        values.push({ row, col, pattern });
      }
    }
    this._commandService.syncExecuteCommand(SetNumfmtCommand.id, {
      unitId: this._workbook.getUnitId(),
      subUnitId: this._worksheet.getSheetId(),
      values
    });
    return this;
  }
  setNumberFormats(patterns) {
    var _a;
    const values = [];
    const { startColumn, startRow, endColumn, endRow } = this._range;
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startColumn; col <= endColumn; col++) {
        const pattern = (_a = patterns[row - startRow]) == null ? void 0 : _a[col - startColumn];
        values.push({ row, col, pattern });
      }
    }
    this._commandService.syncExecuteCommand(SetNumfmtCommand.id, {
      unitId: this._workbook.getUnitId(),
      subUnitId: this._worksheet.getSheetId(),
      values
    });
    return this;
  }
  getNumberFormat() {
    var _a, _b;
    const style = this.getCellStyle();
    return (_b = (_a = style == null ? void 0 : style.numberFormat) == null ? void 0 : _a.pattern) != null ? _b : "";
  }
  getNumberFormats() {
    const styles = this.getCellStyles();
    return styles.map((row) => row.map((cellStyle) => {
      var _a, _b;
      return (_b = (_a = cellStyle == null ? void 0 : cellStyle.numberFormat) == null ? void 0 : _a.pattern) != null ? _b : "";
    }));
  }
};
FRange.extend(FRangeSheetsNumfmtMixin);

// ../packages/sheets-numfmt/src/facade/f-workbook.ts
var FWorkbookSheetsNumfmtMixin = class extends FWorkbook {
  setNumfmtLocal(locale) {
    const sheetsNumfmtCellContentController = this._injector.get(SheetsNumfmtCellContentController);
    sheetsNumfmtCellContentController.setNumfmtLocal(locale);
    return this;
  }
};
FWorkbook.extend(FWorkbookSheetsNumfmtMixin);

export {
  transformComponentKey
};
