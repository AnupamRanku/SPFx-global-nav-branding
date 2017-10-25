"use strict";
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
var sp_core_library_1 = require("@microsoft/sp-core-library");
var sp_http_1 = require("@microsoft/sp-http");
/**
 * @class
 * Service implementation to manage term stores in SharePoint
 * Basic implementation taken from: https://oliviercc.github.io/sp-client-custom-fields/
 */
var SPTermStoreService = (function () {
    /**
     * @function
     * Service constructor
     */
    function SPTermStoreService(config) {
        this.spHttpClient = config.spHttpClient;
        this.siteAbsoluteUrl = config.siteAbsoluteUrl;
    }
    /**
     * @function
     * Gets the collection of term stores in the current SharePoint env
     */
    SPTermStoreService.prototype.getTermsFromTermSetAsync = function (termSetName) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            var contextInfoUrl, httpPostOptions, response, jsonResponse, clientServiceUrl, data, serviceResponse, serviceJSONResponse, result, termSetsCollections, termSetCollection, childTermSets, termSets, termSet, termsCollection, childItems;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!(sp_core_library_1.Environment.type === sp_core_library_1.EnvironmentType.SharePoint ||
                            sp_core_library_1.Environment.type === sp_core_library_1.EnvironmentType.ClassicSharePoint)) return [3 /*break*/, 6];
                        contextInfoUrl = this.siteAbsoluteUrl + "/_api/contextinfo";
                        httpPostOptions = {
                            headers: {
                                "accept": "application/json",
                                "content-type": "application/json"
                            }
                        };
                        return [4 /*yield*/, this.spHttpClient.post(contextInfoUrl, sp_http_1.SPHttpClient.configurations.v1, httpPostOptions)];
                    case 1:
                        response = _a.sent();
                        return [4 /*yield*/, response.json()];
                    case 2:
                        jsonResponse = _a.sent();
                        this.formDigest = jsonResponse.FormDigestValue;
                        clientServiceUrl = this.siteAbsoluteUrl + '/_vti_bin/client.svc/ProcessQuery';
                        data = '<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="JavaScript Client" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectIdentityQuery Id="3" ObjectPathId="1" /><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><ObjectPath Id="8" ObjectPathId="7" /><Query Id="9" ObjectPathId="7"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="Terms" SelectAll="true"><Query SelectAllProperties="false"><Properties /></Query></Property></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="1" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="4" ParentId="1" Name="GetDefaultSiteCollectionTermStore" /><Method Id="7" ParentId="4" Name="GetTermSetsByName"><Parameters><Parameter Type="String">' + termSetName + '</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>';
                        httpPostOptions = {
                            headers: {
                                'accept': 'application/json',
                                'content-type': 'application/json',
                                "X-RequestDigest": this.formDigest
                            },
                            body: data
                        };
                        return [4 /*yield*/, this.spHttpClient.post(clientServiceUrl, sp_http_1.SPHttpClient.configurations.v1, httpPostOptions)];
                    case 3:
                        serviceResponse = _a.sent();
                        return [4 /*yield*/, serviceResponse.json()];
                    case 4:
                        serviceJSONResponse = _a.sent();
                        result = new Array();
                        termSetsCollections = serviceJSONResponse.filter(function (child) { return (child != null && child['_ObjectType_'] !== undefined && child['_ObjectType_'] === "SP.Taxonomy.TermSetCollection"); });
                        if (!(termSetsCollections != null && termSetsCollections.length > 0)) return [3 /*break*/, 6];
                        termSetCollection = termSetsCollections[0];
                        childTermSets = termSetCollection['_Child_Items_'];
                        termSets = childTermSets.filter(function (child) { return (child != null && child['_ObjectType_'] !== undefined && child['_ObjectType_'] === "SP.Taxonomy.TermSet"); });
                        if (!(termSets != null && termSets.length > 0)) return [3 /*break*/, 6];
                        termSet = termSets[0];
                        termsCollection = termSet['Terms'];
                        childItems = termsCollection['_Child_Items_'];
                        return [4 /*yield*/, Promise.all(childItems.map(function (t) { return __awaiter(_this, void 0, void 0, function () {
                                return __generator(this, function (_a) {
                                    switch (_a.label) {
                                        case 0: return [4 /*yield*/, this.projectTermAsync(t)];
                                        case 1: return [2 /*return*/, _a.sent()];
                                    }
                                });
                            }); }))];
                    case 5: return [2 /*return*/, (_a.sent())];
                    case 6: 
                    // Default empty array in case of any missing data
                    return [2 /*return*/, (new Promise(function (resolve, reject) {
                            resolve(new Array());
                        }))];
                }
            });
        });
    };
    /**
     * @function
     * Gets the child terms of another term of the Term Store in the current SharePoint env
     */
    SPTermStoreService.prototype.getChildTermsAsync = function (term) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            var clientServiceUrl, data, httpPostOptions, serviceResponse, serviceJSONResponse, termsCollections, termsCollection, childItems;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!(Number(term['TermsCount']) > 0)) return [3 /*break*/, 4];
                        clientServiceUrl = this.siteAbsoluteUrl + '/_vti_bin/client.svc/ProcessQuery';
                        data = '<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="20" ObjectPathId="19" /><Query Id="21" ObjectPathId="19"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="CustomSortOrder" ScalarProperty="true" /><Property Name="LocalCustomProperties" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><Property Id="19" ParentId="16" Name="Terms" /><Identity Id="16" Name="' + term['_ObjectIdentity_'] + '" /></ObjectPaths></Request>';
                        httpPostOptions = {
                            headers: {
                                'accept': 'application/json',
                                'content-type': 'application/json',
                                "X-RequestDigest": this.formDigest
                            },
                            body: data
                        };
                        return [4 /*yield*/, this.spHttpClient.post(clientServiceUrl, sp_http_1.SPHttpClient.configurations.v1, httpPostOptions)];
                    case 1:
                        serviceResponse = _a.sent();
                        return [4 /*yield*/, serviceResponse.json()];
                    case 2:
                        serviceJSONResponse = _a.sent();
                        termsCollections = serviceJSONResponse.filter(function (child) { return (child != null && child['_ObjectType_'] !== undefined && child['_ObjectType_'] === "SP.Taxonomy.TermCollection"); });
                        if (!(termsCollections != null && termsCollections.length > 0)) return [3 /*break*/, 4];
                        termsCollection = termsCollections[0];
                        childItems = termsCollection['_Child_Items_'];
                        return [4 /*yield*/, Promise.all(childItems.map(function (t) { return __awaiter(_this, void 0, void 0, function () {
                                return __generator(this, function (_a) {
                                    switch (_a.label) {
                                        case 0: return [4 /*yield*/, this.projectTermAsync(t)];
                                        case 1: return [2 /*return*/, _a.sent()];
                                    }
                                });
                            }); }))];
                    case 3: return [2 /*return*/, (_a.sent())];
                    case 4: 
                    // Default empty array in case of any missing data
                    return [2 /*return*/, (new Promise(function (resolve, reject) {
                            resolve(new Array());
                        }))];
                }
            });
        });
    };
    /**
     * @function
     * Projects a Term object into an object of type ISPTermObject, including child terms
     * @param guid
     */
    SPTermStoreService.prototype.projectTermAsync = function (term) {
        return __awaiter(this, void 0, void 0, function () {
            var _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _a = {
                            identity: term['_ObjectIdentity_'] !== undefined ? term['_ObjectIdentity_'] : "",
                            isAvailableForTagging: term['IsAvailableForTagging'] !== undefined ? term['IsAvailableForTagging'] : false,
                            guid: term['Id'] !== undefined ? this.cleanGuid(term['Id']) : "",
                            name: term['Name'] !== undefined ? term['Name'] : "",
                            customSortOrder: term['CustomSortOrder'] !== undefined ? term['CustomSortOrder'] : ""
                        };
                        return [4 /*yield*/, this.getChildTermsAsync(term)];
                    case 1: return [2 /*return*/, (_a.terms = _b.sent(),
                            _a.localCustomProperties = term['LocalCustomProperties'] !== undefined ? term['LocalCustomProperties'] : null,
                            _a)];
                }
            });
        });
    };
    /**
     * @function
     * Clean the Guid from the Web Service response
     * @param guid
     */
    SPTermStoreService.prototype.cleanGuid = function (guid) {
        if (guid !== undefined)
            return guid.replace('/Guid(', '').replace('/', '').replace(')', '');
        else
            return '';
    };
    return SPTermStoreService;
}());
exports.SPTermStoreService = SPTermStoreService;

//# sourceMappingURL=SPTermStoreService.js.map
