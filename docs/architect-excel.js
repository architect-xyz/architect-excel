/*
 * ATTENTION: The "eval" devtool has been used (maybe by default in mode: "development").
 * This devtool is neither made for production nor for readable output files.
 * It uses "eval()" calls to create a separate source file in the browser devtools.
 * If you are trying to read the output file, select a different devtool (https://webpack.js.org/configuration/devtool/)
 * or disable the default devtool with "devtool: false".
 * If you are looking for production-ready output files, see mode: "production" (https://webpack.js.org/configuration/mode/).
 */
/******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ "./src/architect-excel.ts":
/*!********************************!*\
  !*** ./src/architect-excel.ts ***!
  \********************************/
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {

eval("\n// Excel Plugin Example: Query Backend API Using TypeScript\nvar __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {\n    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }\n    return new (P || (P = Promise))(function (resolve, reject) {\n        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }\n        function rejected(value) { try { step(generator[\"throw\"](value)); } catch (e) { reject(e); } }\n        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }\n        step((generator = generator.apply(thisArg, _arguments || [])).next());\n    });\n};\nvar __generator = (this && this.__generator) || function (thisArg, body) {\n    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g = Object.create((typeof Iterator === \"function\" ? Iterator : Object).prototype);\n    return g.next = verb(0), g[\"throw\"] = verb(1), g[\"return\"] = verb(2), typeof Symbol === \"function\" && (g[Symbol.iterator] = function() { return this; }), g;\n    function verb(n) { return function (v) { return step([n, v]); }; }\n    function step(op) {\n        if (f) throw new TypeError(\"Generator is already executing.\");\n        while (g && (g = 0, op[0] && (_ = 0)), _) try {\n            if (f = 1, y && (t = op[0] & 2 ? y[\"return\"] : op[0] ? y[\"throw\"] || ((t = y[\"return\"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;\n            if (y = 0, t) op = [op[0] & 2, t.value];\n            switch (op[0]) {\n                case 0: case 1: t = op; break;\n                case 4: _.label++; return { value: op[1], done: false };\n                case 5: _.label++; y = op[1]; op = [0]; continue;\n                case 7: op = _.ops.pop(); _.trys.pop(); continue;\n                default:\n                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }\n                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }\n                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }\n                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }\n                    if (t[2]) _.ops.pop();\n                    _.trys.pop(); continue;\n            }\n            op = body.call(thisArg, _);\n        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }\n        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };\n    }\n};\nObject.defineProperty(exports, \"__esModule\", ({ value: true }));\nexports.fetchData = fetchData;\nexports.fetchMarketSnapshot = fetchMarketSnapshot;\n/*\nimportant functions:\n\nlast price\nmid price\nbid price\nask price\nclose\nopen\nday high\nday low\n*/\n/// <reference types=\"office-js\" />\nvar sdk_1 = __webpack_require__(/*! @afintech/sdk */ \"./node_modules/@afintech/sdk/dist/cjs/index.js\");\nvar config = {\n    host: 'https://app.architect.co',\n    apiKey: '',\n    apiSecret: '',\n    tradingMode: 'live',\n};\nvar client;\n/**\n * Uses the saved API key/secret for a custom function.\n * @customfunction\n */\nfunction fetchData() {\n    return __awaiter(this, void 0, void 0, function () {\n        var apiKey, apiSecret;\n        return __generator(this, function (_a) {\n            apiKey = Office.context.document.settings.get('apiKey');\n            apiSecret = Office.context.document.settings.get('apiSecret');\n            if (!apiKey || !apiSecret) {\n                throw new Error(\"API key or secret not set. Use the ribbon to configure them.\");\n            }\n            // Example usage with the API\n            return [2 /*return*/, \"Using API Key: \".concat(apiKey, \", Secret: \").concat(apiSecret)];\n        });\n    });\n}\n/**\n * Initialize the client with user-provided API key and secret\n * @customfunction\n * @param sheetName Name of the worksheet containing API credentials\n */\nfunction initializeClient(sheetName) {\n    return __awaiter(this, void 0, void 0, function () {\n        var _this = this;\n        return __generator(this, function (_a) {\n            switch (_a.label) {\n                case 0: return [4 /*yield*/, Excel.run(function (context) { return __awaiter(_this, void 0, void 0, function () {\n                        var workbook, worksheet, apiKey, apiSecret;\n                        return __generator(this, function (_a) {\n                            switch (_a.label) {\n                                case 0:\n                                    workbook = context.workbook;\n                                    worksheet = workbook.worksheets.getItemOrNullObject(sheetName);\n                                    return [4 /*yield*/, context.sync()];\n                                case 1:\n                                    _a.sent();\n                                    if (!worksheet.isNullObject) return [3 /*break*/, 3];\n                                    worksheet = workbook.worksheets.add(sheetName);\n                                    worksheet.getRange('A1').values = [['API_KEY:']];\n                                    worksheet.getRange('B1').values = [['']];\n                                    worksheet.getRange('A2').values = [['API_SECRET:']];\n                                    worksheet.getRange('B2').values = [['']];\n                                    console.log(\"Created \".concat(sheetName, \" worksheet. Please fill in your API key and secret.\"));\n                                    return [4 /*yield*/, context.sync()];\n                                case 2:\n                                    _a.sent();\n                                    return [2 /*return*/];\n                                case 3:\n                                    apiKey = worksheet.getRange('B1').values[0][0];\n                                    apiSecret = worksheet.getRange('B2').values[0][0];\n                                    if (!apiKey || !apiSecret) {\n                                        throw new Error('API Key and Secret must be provided in cells B1 and B2.');\n                                    }\n                                    config.apiKey = apiKey;\n                                    config.apiSecret = apiSecret;\n                                    client = (0, sdk_1.create)(config);\n                                    return [2 /*return*/];\n                            }\n                        });\n                    }); })];\n                case 1:\n                    _a.sent();\n                    return [2 /*return*/];\n            }\n        });\n    });\n}\n/**\n * Fetch market snapshot and populate Excel worksheet\n * @param market Market identifier\n */\nfunction fetchMarketSnapshot(market) {\n    return __awaiter(this, void 0, void 0, function () {\n        var snapshot, error_1;\n        return __generator(this, function (_a) {\n            switch (_a.label) {\n                case 0:\n                    _a.trys.push([0, 2, , 3]);\n                    return [4 /*yield*/, client.marketSnapshot([], market)];\n                case 1:\n                    snapshot = _a.sent();\n                    if (!snapshot) {\n                        console.error('No snapshot data received');\n                        return [2 /*return*/];\n                    }\n                    return [3 /*break*/, 3];\n                case 2:\n                    error_1 = _a.sent();\n                    console.error('Error fetching market snapshot:', error_1);\n                    return [3 /*break*/, 3];\n                case 3: return [2 /*return*/];\n            }\n        });\n    });\n}\n/**\n * Main function to initialize and execute the plugin\n */\nfunction main() {\n    return __awaiter(this, void 0, void 0, function () {\n        var sheetName, market, error_2;\n        return __generator(this, function (_a) {\n            switch (_a.label) {\n                case 0:\n                    _a.trys.push([0, 3, , 4]);\n                    sheetName = 'ARCHITECT_CONFIG';\n                    market = 'MES 20250321 CME Future/USD*CME/CQG';\n                    return [4 /*yield*/, initializeClient(sheetName)];\n                case 1:\n                    _a.sent();\n                    return [4 /*yield*/, fetchMarketSnapshot(market)];\n                case 2:\n                    _a.sent();\n                    return [3 /*break*/, 4];\n                case 3:\n                    error_2 = _a.sent();\n                    console.error('Error in main:', error_2);\n                    return [3 /*break*/, 4];\n                case 4: return [2 /*return*/];\n            }\n        });\n    });\n}\n// Entry point for the Office Add-in\nOffice.onReady(function (info) {\n    if (info.host === Office.HostType.Excel) {\n        console.log('Excel Add-in ready.');\n        main();\n    }\n});\n\n\n//# sourceURL=webpack://architect-excel/./src/architect-excel.ts?");

/***/ }),

/***/ "./node_modules/@afintech/sdk/dist/cjs/index.js":
/*!******************************************************!*\
  !*** ./node_modules/@afintech/sdk/dist/cjs/index.js ***!
  \******************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

eval("__webpack_require__.r(__webpack_exports__);\n\nvar __createBinding =\n  (undefined && undefined.__createBinding) ||\n  (Object.create\n    ? function (o, m, k, k2) {\n        if (k2 === undefined) k2 = k;\n        var desc = Object.getOwnPropertyDescriptor(m, k);\n        if (\n          !desc ||\n          ('get' in desc ? !m.__esModule : desc.writable || desc.configurable)\n        ) {\n          desc = {\n            enumerable: true,\n            get: function () {\n              return m[k];\n            },\n          };\n        }\n        Object.defineProperty(o, k2, desc);\n      }\n    : function (o, m, k, k2) {\n        if (k2 === undefined) k2 = k;\n        o[k2] = m[k];\n      });\nvar __setModuleDefault =\n  (undefined && undefined.__setModuleDefault) ||\n  (Object.create\n    ? function (o, v) {\n        Object.defineProperty(o, 'default', { enumerable: true, value: v });\n      }\n    : function (o, v) {\n        o['default'] = v;\n      });\nvar __importStar =\n  (undefined && undefined.__importStar) ||\n  function (mod) {\n    if (mod && mod.__esModule) return mod;\n    var result = {};\n    if (mod != null)\n      for (var k in mod)\n        if (k !== 'default' && Object.prototype.hasOwnProperty.call(mod, k))\n          __createBinding(result, mod, k);\n    __setModuleDefault(result, mod);\n    return result;\n  };\nObject.defineProperty(exports, '__esModule', { value: true });\nexports.enums = exports.L1BookSnapshot = void 0;\nexports.create = create;\nconst graphql_http_1 = require('graphql-http');\nexports.L1BookSnapshot = __importStar(require('./grpc/l1booksnapshot.js'));\nconst sdk_js_1 = require('./sdk.js');\nconst graphql_js_1 = require('./graphql/graphql.js');\n/**\n * Creates an instance of the Architect SDK Client\n */\nfunction create(config) {\n  return new sdk_js_1.Client(config, graphql_http_1.createClient);\n}\n// TODO: codegen this enum generator\nexports.enums = {\n  AlgoControlCommand: graphql_js_1.AlgoControlCommand,\n  AlgoKind: graphql_js_1.AlgoKind,\n  AlgoRunningStatus: graphql_js_1.AlgoRunningStatus,\n  CandleWidth: graphql_js_1.CandleWidth,\n  CmeSecurityType: graphql_js_1.CmeSecurityType,\n  CreateOrderType: graphql_js_1.CreateOrderType,\n  CreateTimeInForceInstruction: graphql_js_1.CreateTimeInForceInstruction,\n  EnvironmentKind: graphql_js_1.EnvironmentKind,\n  FillKind: graphql_js_1.FillKind,\n  LicenseTier: graphql_js_1.LicenseTier,\n  MMAlgoKind: graphql_js_1.MMAlgoKind,\n  MinOrderQuantityUnit: graphql_js_1.MinOrderQuantityUnit,\n  OrderSource: graphql_js_1.OrderSource,\n  OrderStateFlags: graphql_js_1.OrderStateFlags,\n  Reason: graphql_js_1.Reason,\n  ReferencePrice: graphql_js_1.ReferencePrice,\n};\n\n\n//# sourceURL=webpack://architect-excel/./node_modules/@afintech/sdk/dist/cjs/index.js?");

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	(() => {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = (exports) => {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	})();
/******/ 	
/************************************************************************/
/******/ 	
/******/ 	// startup
/******/ 	// Load entry module and return exports
/******/ 	// This entry module is referenced by other modules so it can't be inlined
/******/ 	var __webpack_exports__ = __webpack_require__("./src/architect-excel.ts");
/******/ 	
/******/ })()
;