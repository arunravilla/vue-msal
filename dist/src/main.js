"use strict";
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
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
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
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
var __spreadArray = (this && this.__spreadArray) || function (to, from) {
    for (var i = 0, il = from.length, j = to.length; i < il; i++, j++)
        to[j] = from[i];
    return to;
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.MSAL = void 0;
var lodash_1 = __importDefault(require("lodash"));
var axios_1 = __importDefault(require("axios"));
var UserAgentApplicationExtended_1 = require("./UserAgentApplicationExtended");
var MSAL = /** @class */ (function () {
    function MSAL(options) {
        var _this = this;
        this.options = options;
        this.tokenExpirationTimers = {};
        this.data = {
            isAuthenticated: false,
            accessToken: '',
            idToken: '',
            user: {},
            graph: {},
            custom: {}
        };
        this.callbackQueue = [];
        this.auth = {
            clientId: '',
            authority: '',
            tenantId: 'common',
            tenantName: 'login.microsoftonline.com',
            validateAuthority: true,
            redirectUri: window.location.href,
            postLogoutRedirectUri: window.location.href,
            navigateToLoginRequestUrl: true,
            requireAuthOnInitialize: false,
            autoRefreshToken: true,
            onAuthentication: function (error, response) { },
            onToken: function (error, response) { },
            beforeSignOut: function () { }
        };
        this.cache = {
            cacheLocation: 'localStorage',
            storeAuthStateInCookie: true
        };
        this.request = {
            scopes: ["user.read"]
        };
        this.graph = {
            callAfterInit: false,
            endpoints: { profile: '/me' },
            baseUrl: 'https://graph.microsoft.com/v1.0',
            onResponse: function (response) { }
        };
        if (!options.auth.clientId) {
            throw new Error('auth.clientId is required');
        }
        this.auth = Object.assign(this.auth, options.auth);
        this.cache = Object.assign(this.cache, options.cache);
        this.request = Object.assign(this.request, options.request);
        this.graph = Object.assign(this.graph, options.graph);
        this.lib = new UserAgentApplicationExtended_1.UserAgentApplicationExtended({
            auth: {
                clientId: this.auth.clientId,
                authority: this.auth.authority || "https://" + this.auth.tenantName + "/" + this.auth.tenantId,
                validateAuthority: this.auth.validateAuthority,
                redirectUri: this.auth.redirectUri,
                postLogoutRedirectUri: this.auth.postLogoutRedirectUri,
                navigateToLoginRequestUrl: this.auth.navigateToLoginRequestUrl
            },
            cache: this.cache,
            system: options.system
        });
        this.getSavedCallbacks();
        this.executeCallbacks();
        // Register Callbacks for redirect flow
        this.lib.handleRedirectCallback(function (error, response) {
            if (!_this.isAuthenticated()) {
                _this.saveCallback('auth.onAuthentication', error, response);
            }
            else {
                _this.acquireToken();
            }
        });
        if (this.auth.requireAuthOnInitialize) {
            this.signIn();
        }
        this.data.isAuthenticated = this.isAuthenticated();
        if (this.data.isAuthenticated) {
            this.data.user = this.lib.getAccount();
            this.acquireToken().then(function () {
                if (_this.graph.callAfterInit) {
                    _this.initialMSGraphCall();
                }
            });
        }
        this.getStoredCustomData();
    }
    MSAL.prototype.signIn = function () {
        if (!this.lib.isCallback(window.location.hash) && !this.lib.getAccount()) {
            // request can be used for login or token request, however in more complex situations this can have diverging options
            this.lib.loginRedirect(this.request);
        }
    };
    MSAL.prototype.signOut = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this.options.auth.beforeSignOut) return [3 /*break*/, 2];
                        return [4 /*yield*/, this.options.auth.beforeSignOut(this)];
                    case 1:
                        _a.sent();
                        _a.label = 2;
                    case 2:
                        this.lib.logout();
                        return [2 /*return*/];
                }
            });
        });
    };
    MSAL.prototype.isAuthenticated = function () {
        return !this.lib.isCallback(window.location.hash) && !!this.lib.getAccount();
    };
    MSAL.prototype.acquireToken = function (request, retries) {
        if (request === void 0) { request = this.request; }
        if (retries === void 0) { retries = 0; }
        return __awaiter(this, void 0, void 0, function () {
            var response, error_1;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 6]);
                        return [4 /*yield*/, this.lib.acquireTokenSilent(request)];
                    case 1:
                        response = _a.sent();
                        this.handleTokenResponse(null, response);
                        return [2 /*return*/, response];
                    case 2:
                        error_1 = _a.sent();
                        if (!this.requiresInteraction(error_1.errorCode)) return [3 /*break*/, 3];
                        this.lib.acquireTokenRedirect(request);
                        return [3 /*break*/, 5];
                    case 3:
                        if (!(retries > 0)) return [3 /*break*/, 5];
                        return [4 /*yield*/, new Promise(function (resolve) {
                                setTimeout(function () { return __awaiter(_this, void 0, void 0, function () {
                                    var res;
                                    return __generator(this, function (_a) {
                                        switch (_a.label) {
                                            case 0: return [4 /*yield*/, this.acquireToken(request, retries - 1)];
                                            case 1:
                                                res = _a.sent();
                                                resolve(res);
                                                return [2 /*return*/];
                                        }
                                    });
                                }); }, 60 * 1000);
                            })];
                    case 4: return [2 /*return*/, _a.sent()];
                    case 5: return [2 /*return*/, false];
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    MSAL.prototype.handleTokenResponse = function (error, response) {
        if (error) {
            this.saveCallback('auth.onToken', error, null);
            return;
        }
        var setCallback = false;
        if (response.tokenType === 'access_token' && this.data.accessToken !== response.accessToken) {
            this.setToken('accessToken', response.accessToken, response.expiresOn, response.scopes);
            setCallback = true;
        }
        if (this.data.idToken !== response.idToken.rawIdToken) {
            this.setToken('idToken', response.idToken.rawIdToken, new Date(response.idToken.expiration * 1000), [this.auth.clientId]);
            setCallback = true;
        }
        if (setCallback) {
            this.saveCallback('auth.onToken', null, response);
        }
    };
    MSAL.prototype.setToken = function (tokenType, token, expiresOn, scopes) {
        var _this = this;
        var expirationOffset = this.lib.config.system.tokenRenewalOffsetSeconds * 1000;
        var expiration = expiresOn.getTime() - (new Date()).getTime() - expirationOffset;
        if (expiration >= 0) {
            this.data[tokenType] = token;
        }
        if (this.tokenExpirationTimers[tokenType])
            clearTimeout(this.tokenExpirationTimers[tokenType]);
        this.tokenExpirationTimers[tokenType] = window.setTimeout(function () { return __awaiter(_this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this.auth.autoRefreshToken) return [3 /*break*/, 2];
                        return [4 /*yield*/, this.acquireToken({ scopes: scopes }, 3)];
                    case 1:
                        _a.sent();
                        return [3 /*break*/, 3];
                    case 2:
                        this.data[tokenType] = '';
                        _a.label = 3;
                    case 3: return [2 /*return*/];
                }
            });
        }); }, expiration);
    };
    MSAL.prototype.requiresInteraction = function (errorCode) {
        if (!errorCode || !errorCode.length) {
            return false;
        }
        return errorCode === "consent_required" ||
            errorCode === "interaction_required" ||
            errorCode === "login_required";
    };
    // MS GRAPH
    MSAL.prototype.initialMSGraphCall = function () {
        return __awaiter(this, void 0, void 0, function () {
            var callback, initEndpoints, resultsObj_1, forcedIds, endpoints, id, storedIds, storedData, _a, singleRequests, batchRequests_1, singlePromises, batchPromises, mixedResults, resultsToSave_1, error_2;
            var _this = this;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        callback = this.graph.onResponse;
                        initEndpoints = this.graph.endpoints;
                        if (!(typeof initEndpoints === 'object' && !lodash_1.default.isEmpty(initEndpoints))) return [3 /*break*/, 5];
                        resultsObj_1 = {};
                        forcedIds = [];
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        endpoints = {};
                        for (id in initEndpoints) {
                            endpoints[id] = this.getEndpointObject(initEndpoints[id]);
                            if (endpoints[id].force) {
                                forcedIds.push(id);
                            }
                        }
                        storedIds = [];
                        storedData = this.lib.store.getItem("msal.msgraph-" + this.data.accessToken);
                        if (storedData) {
                            storedData = JSON.parse(storedData);
                            storedIds = Object.keys(storedData);
                            Object.assign(resultsObj_1, storedData);
                        }
                        _a = this.categorizeRequests(endpoints, lodash_1.default.difference(storedIds, forcedIds)), singleRequests = _a.singleRequests, batchRequests_1 = _a.batchRequests;
                        singlePromises = singleRequests.map(function (endpoint) { return __awaiter(_this, void 0, void 0, function () {
                            var res, _a, _b;
                            return __generator(this, function (_c) {
                                switch (_c.label) {
                                    case 0:
                                        res = {};
                                        _a = res;
                                        _b = endpoint.id;
                                        return [4 /*yield*/, this.msGraph(endpoint)];
                                    case 1:
                                        _a[_b] = _c.sent();
                                        return [2 /*return*/, res];
                                }
                            });
                        }); });
                        batchPromises = Object.keys(batchRequests_1).map(function (key) {
                            var batchUrl = (key === 'default') ? undefined : key;
                            return _this.msGraph(batchRequests_1[key], batchUrl);
                        });
                        return [4 /*yield*/, Promise.all(__spreadArray(__spreadArray([], singlePromises), batchPromises))];
                    case 2:
                        mixedResults = _b.sent();
                        mixedResults.map(function (res) {
                            for (var key in res) {
                                res[key] = res[key].body;
                            }
                            Object.assign(resultsObj_1, res);
                        });
                        resultsToSave_1 = __assign({}, resultsObj_1);
                        forcedIds.map(function (id) { return delete resultsToSave_1[id]; });
                        this.lib.store.setItem("msal.msgraph-" + this.data.accessToken, JSON.stringify(resultsToSave_1));
                        this.data.graph = resultsObj_1;
                        return [3 /*break*/, 4];
                    case 3:
                        error_2 = _b.sent();
                        console.error(error_2);
                        return [3 /*break*/, 4];
                    case 4:
                        if (callback)
                            this.saveCallback('graph.onResponse', this.data.graph);
                        _b.label = 5;
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    MSAL.prototype.msGraph = function (endpoints, batchUrl) {
        if (batchUrl === void 0) { batchUrl = undefined; }
        return __awaiter(this, void 0, void 0, function () {
            var error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 5, , 6]);
                        if (!Array.isArray(endpoints)) return [3 /*break*/, 2];
                        return [4 /*yield*/, this.executeBatchRequest(endpoints, batchUrl)];
                    case 1: return [2 /*return*/, _a.sent()];
                    case 2: return [4 /*yield*/, this.executeSingleRequest(endpoints)];
                    case 3: return [2 /*return*/, _a.sent()];
                    case 4: return [3 /*break*/, 6];
                    case 5:
                        error_3 = _a.sent();
                        throw error_3;
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    MSAL.prototype.executeBatchRequest = function (endpoints, batchUrl) {
        if (batchUrl === void 0) { batchUrl = this.graph.baseUrl; }
        return __awaiter(this, void 0, void 0, function () {
            var requests, data, result, keys, numKeys;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        requests = endpoints.map(function (endpoint, index) { return _this.createRequest(endpoint, index); });
                        return [4 /*yield*/, axios_1.default.request({
                                url: batchUrl + "/$batch",
                                method: 'POST',
                                data: { requests: requests },
                                headers: { Authorization: "Bearer " + this.data.accessToken },
                                responseType: 'json'
                            })];
                    case 1:
                        data = (_a.sent()).data;
                        result = {};
                        data.responses.map(function (response) {
                            var key = response.id;
                            delete response.id;
                            return result[key] = response;
                        });
                        keys = Object.keys(result);
                        numKeys = keys.sort().filter(function (key, index) {
                            if (key.search('defaultID-') === 0) {
                                key = key.replace('defaultID-', '');
                            }
                            return parseInt(key) === index;
                        });
                        if (numKeys.length === keys.length) {
                            result = lodash_1.default.values(result);
                        }
                        return [2 /*return*/, result];
                }
            });
        });
    };
    MSAL.prototype.executeSingleRequest = function (endpoint) {
        return __awaiter(this, void 0, void 0, function () {
            var request, res;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        request = this.createRequest(endpoint);
                        if (request.url.search('http') !== 0) {
                            request.url = this.graph.baseUrl + request.url;
                        }
                        return [4 /*yield*/, axios_1.default.request(lodash_1.default.defaultsDeep(request, {
                                url: request.url,
                                method: request.method,
                                responseType: 'json',
                                headers: { Authorization: "Bearer " + this.data.accessToken }
                            }))];
                    case 1:
                        res = _a.sent();
                        return [2 /*return*/, {
                                status: res.status,
                                headers: res.headers,
                                body: res.data
                            }];
                }
            });
        });
    };
    MSAL.prototype.createRequest = function (endpoint, index) {
        if (index === void 0) { index = 0; }
        var request = {
            url: '',
            method: 'GET',
            id: "defaultID-" + index
        };
        endpoint = this.getEndpointObject(endpoint);
        if (endpoint.url) {
            Object.assign(request, endpoint);
        }
        else {
            throw ({ error: 'invalid endpoint', endpoint: endpoint });
        }
        return request;
    };
    MSAL.prototype.categorizeRequests = function (endpoints, excludeIds) {
        var res = {
            singleRequests: [],
            batchRequests: {}
        };
        for (var key in endpoints) {
            var endpoint = __assign({ id: key }, endpoints[key]);
            if (!lodash_1.default.includes(excludeIds, key)) {
                if (endpoint.batchUrl) {
                    var batchUrl = endpoint.batchUrl;
                    delete endpoint.batchUrl;
                    if (!res.batchRequests.hasOwnProperty(batchUrl)) {
                        res.batchRequests[batchUrl] = [];
                    }
                    res.batchRequests[batchUrl].push(endpoint);
                }
                else {
                    res.singleRequests.push(endpoint);
                }
            }
        }
        return res;
    };
    MSAL.prototype.getEndpointObject = function (endpoint) {
        if (typeof endpoint === "string") {
            endpoint = { url: endpoint };
        }
        if (typeof endpoint === "object" && !endpoint.url) {
            throw new Error('invalid endpoint url');
        }
        return endpoint;
    };
    // CUSTOM DATA
    MSAL.prototype.saveCustomData = function (key, data) {
        if (!this.data.custom.hasOwnProperty(key)) {
            this.data.custom[key] = null;
        }
        this.data.custom[key] = data;
        this.storeCustomData();
    };
    MSAL.prototype.storeCustomData = function () {
        if (!lodash_1.default.isEmpty(this.data.custom)) {
            this.lib.store.setItem('msal.custom', JSON.stringify(this.data.custom));
        }
        else {
            this.lib.store.removeItem('msal.custom');
        }
    };
    MSAL.prototype.getStoredCustomData = function () {
        var customData = {};
        var customDataStr = this.lib.store.getItem('msal.custom');
        if (customDataStr) {
            customData = JSON.parse(customDataStr);
        }
        this.data.custom = customData;
    };
    // CALLBACKS
    MSAL.prototype.saveCallback = function (callbackPath) {
        var args = [];
        for (var _i = 1; _i < arguments.length; _i++) {
            args[_i - 1] = arguments[_i];
        }
        if (lodash_1.default.get(this.options, callbackPath)) {
            var callbackQueueObject_1 = {
                id: lodash_1.default.uniqueId("cb-" + callbackPath),
                callback: callbackPath,
                arguments: args
            };
            lodash_1.default.remove(this.callbackQueue, function (obj) { return obj.id === callbackQueueObject_1.id; });
            this.callbackQueue.push(callbackQueueObject_1);
            this.storeCallbackQueue();
            this.executeCallbacks([callbackQueueObject_1]);
        }
    };
    MSAL.prototype.getSavedCallbacks = function () {
        var callbackQueueStr = this.lib.store.getItem('msal.callbackqueue');
        if (callbackQueueStr) {
            this.callbackQueue = __spreadArray(__spreadArray([], this.callbackQueue), JSON.parse(callbackQueueStr));
        }
    };
    MSAL.prototype.executeCallbacks = function (callbacksToExec) {
        if (callbacksToExec === void 0) { callbacksToExec = this.callbackQueue; }
        return __awaiter(this, void 0, void 0, function () {
            var _loop_1, this_1, _a, _b, _i, i;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        if (!callbacksToExec.length) return [3 /*break*/, 4];
                        _loop_1 = function (i) {
                            var cb, callback, e_1;
                            return __generator(this, function (_d) {
                                switch (_d.label) {
                                    case 0:
                                        cb = callbacksToExec[i];
                                        callback = lodash_1.default.get(this_1.options, cb.callback);
                                        _d.label = 1;
                                    case 1:
                                        _d.trys.push([1, 3, , 4]);
                                        return [4 /*yield*/, callback.apply(void 0, __spreadArray([this_1], cb.arguments))];
                                    case 2:
                                        _d.sent();
                                        lodash_1.default.remove(this_1.callbackQueue, function (currentCb) {
                                            return cb.id === currentCb.id;
                                        });
                                        this_1.storeCallbackQueue();
                                        return [3 /*break*/, 4];
                                    case 3:
                                        e_1 = _d.sent();
                                        console.warn("Callback '" + cb.id + "' failed with error: ", e_1.message);
                                        return [3 /*break*/, 4];
                                    case 4: return [2 /*return*/];
                                }
                            });
                        };
                        this_1 = this;
                        _a = [];
                        for (_b in callbacksToExec)
                            _a.push(_b);
                        _i = 0;
                        _c.label = 1;
                    case 1:
                        if (!(_i < _a.length)) return [3 /*break*/, 4];
                        i = _a[_i];
                        return [5 /*yield**/, _loop_1(i)];
                    case 2:
                        _c.sent();
                        _c.label = 3;
                    case 3:
                        _i++;
                        return [3 /*break*/, 1];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    MSAL.prototype.storeCallbackQueue = function () {
        if (this.callbackQueue.length) {
            this.lib.store.setItem('msal.callbackqueue', JSON.stringify(this.callbackQueue));
        }
        else {
            this.lib.store.removeItem('msal.callbackqueue');
        }
    };
    return MSAL;
}());
exports.MSAL = MSAL;
//# sourceMappingURL=main.js.map