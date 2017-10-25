"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = y[op[0] & 2 ? "return" : op[0] ? "throw" : "next"]) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [0, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var React = require("react");
var ReactDom = require("react-dom");
var decorators_1 = require("@microsoft/decorators");
var sp_core_library_1 = require("@microsoft/sp-core-library");
var sp_application_base_1 = require("@microsoft/sp-application-base");
var TenantGlobalNavBar_1 = require("./components/TenantGlobalNavBar");
var TenantGlobalFooterBar_1 = require("./components/TenantGlobalFooterBar");
var SPTermStore = require("../../components/SPTermStoreService");
var strings = require("TenantGlobalNavBarApplicationCustomizerStrings");
var LOG_SOURCE = 'TenantGlobalNavBarApplicationCustomizer';
/** A Custom Action which can be run during execution of a Client Side Application */
var TenantGlobalNavBarApplicationCustomizer = (function (_super) {
    __extends(TenantGlobalNavBarApplicationCustomizer, _super);
    function TenantGlobalNavBarApplicationCustomizer() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    TenantGlobalNavBarApplicationCustomizer.prototype.onInit = function () {
        return __awaiter(this, void 0, void 0, function () {
            var termStoreService, _a, _b;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        sp_core_library_1.Log.info(LOG_SOURCE, "Initialized " + strings.Title);
                        termStoreService = new SPTermStore.SPTermStoreService({
                            spHttpClient: this.context.spHttpClient,
                            siteAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
                        });
                        if (!(this.properties.TopMenuTermSet != null)) return [3 /*break*/, 2];
                        _a = this;
                        return [4 /*yield*/, termStoreService.getTermsFromTermSetAsync(this.properties.TopMenuTermSet)];
                    case 1:
                        _a._topMenuItems = _c.sent();
                        _c.label = 2;
                    case 2:
                        if (!(this.properties.BottomMenuTermSet != null)) return [3 /*break*/, 4];
                        _b = this;
                        return [4 /*yield*/, termStoreService.getTermsFromTermSetAsync(this.properties.BottomMenuTermSet)];
                    case 3:
                        _b._bottomMenuItems = _c.sent();
                        _c.label = 4;
                    case 4:
                        // Call render method for generating the needed html elements
                        this._renderPlaceHolders();
                        return [2 /*return*/, Promise.resolve()];
                }
            });
        });
    };
    TenantGlobalNavBarApplicationCustomizer.prototype._renderPlaceHolders = function () {
        console.log('Available placeholders: ', this.context.placeholderProvider.placeholderNames.map(function (name) { return sp_application_base_1.PlaceholderName[name]; }).join(', '));
        // Handling the top placeholder
        if (!this._topPlaceholder) {
            this._topPlaceholder =
                this.context.placeholderProvider.tryCreateContent(sp_application_base_1.PlaceholderName.Top, { onDispose: this._onDispose });
            // The extension should not assume that the expected placeholder is available.
            if (!this._topPlaceholder) {
                console.error('The expected placeholder (Top) was not found.');
                return;
            }
            if (this._topMenuItems != null && this._topMenuItems.length > 0) {
                var element = React.createElement(TenantGlobalNavBar_1.default, {
                    menuItems: this._topMenuItems,
                });
                ReactDom.render(element, this._topPlaceholder.domElement);
            }
        }
        // Handling the bottom placeholder
        if (!this._bottomPlaceholder) {
            this._bottomPlaceholder =
                this.context.placeholderProvider.tryCreateContent(sp_application_base_1.PlaceholderName.Bottom, { onDispose: this._onDispose });
            // The extension should not assume that the expected placeholder is available.
            if (!this._bottomPlaceholder) {
                console.error('The expected placeholder (Bottom) was not found.');
                return;
            }
            if (this._bottomMenuItems != null && this._bottomMenuItems.length > 0) {
                var element = React.createElement(TenantGlobalFooterBar_1.default, {
                    menuItems: this._bottomMenuItems,
                });
                ReactDom.render(element, this._bottomPlaceholder.domElement);
            }
        }
    };
    TenantGlobalNavBarApplicationCustomizer.prototype._onDispose = function () {
        console.log('[TenantGlobalNavBarApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
    };
    __decorate([
        decorators_1.override
    ], TenantGlobalNavBarApplicationCustomizer.prototype, "onInit", null);
    return TenantGlobalNavBarApplicationCustomizer;
}(sp_application_base_1.BaseApplicationCustomizer));
exports.default = TenantGlobalNavBarApplicationCustomizer;

//# sourceMappingURL=TenantGlobalNavBarApplicationCustomizer.js.map
