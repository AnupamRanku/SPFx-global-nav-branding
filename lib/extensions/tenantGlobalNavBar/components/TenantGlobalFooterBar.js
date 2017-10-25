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
Object.defineProperty(exports, "__esModule", { value: true });
var React = require("react");
var AppCustomizer_module_scss_1 = require("../AppCustomizer.module.scss");
var CommandBar_1 = require("office-ui-fabric-react/lib/CommandBar");
var ContextualMenu_1 = require("office-ui-fabric-react/lib/ContextualMenu");
var TenantGlobalFooterBar = (function (_super) {
    __extends(TenantGlobalFooterBar, _super);
    /**
    * Main constructor for the component
    */
    function TenantGlobalFooterBar() {
        var _this = _super.call(this) || this;
        _this.state = {};
        return _this;
    }
    TenantGlobalFooterBar.prototype.projectMenuItem = function (menuItem, itemType) {
        var _this = this;
        return ({
            key: menuItem.identity,
            name: menuItem.name,
            itemType: itemType,
            href: menuItem.terms.length == 0 ?
                (menuItem.localCustomProperties["_Sys_Nav_SimpleLinkUrl"] != undefined ?
                    menuItem.localCustomProperties["_Sys_Nav_SimpleLinkUrl"]
                    : null)
                : null,
            subMenuProps: menuItem.terms.length > 0 ?
                { items: menuItem.terms.map(function (i) { return (_this.projectMenuItem(i, ContextualMenu_1.ContextualMenuItemType.Normal)); }) }
                : null,
            isSubMenu: itemType != ContextualMenu_1.ContextualMenuItemType.Header,
        });
    };
    TenantGlobalFooterBar.prototype.render = function () {
        var _this = this;
        var commandBarItems = this.props.menuItems.map(function (i) {
            return (_this.projectMenuItem(i, ContextualMenu_1.ContextualMenuItemType.Header));
        });
        return (React.createElement("div", { className: "ms-bgColor-neutralLighter ms-fontColor-white " + AppCustomizer_module_scss_1.default.app },
            React.createElement("div", { className: "ms-bgColor-neutralLighter ms-fontColor-white " + AppCustomizer_module_scss_1.default.bottom },
                React.createElement(CommandBar_1.CommandBar, { className: AppCustomizer_module_scss_1.default.commandBarFooter, isSearchBoxVisible: false, elipisisAriaLabel: 'More options', items: commandBarItems }))));
    };
    return TenantGlobalFooterBar;
}(React.Component));
exports.default = TenantGlobalFooterBar;

//# sourceMappingURL=TenantGlobalFooterBar.js.map
