"use strict";

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.MSAL = void 0;

var _lodash = _interopRequireDefault(require("lodash"));

var _axios = _interopRequireDefault(require("axios"));

var _UserAgentApplicationExtended = require("./UserAgentApplicationExtended");

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

class MSAL {
  tokenExpirationTimers = {};
  data = {
    isAuthenticated: false,
    accessToken: '',
    idToken: '',
    user: {},
    graph: {},
    custom: {}
  };
  callbackQueue = [];
  auth = {
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
    onAuthentication: (error, response) => {},
    onToken: (error, response) => {},
    beforeSignOut: () => {}
  };
  cache = {
    cacheLocation: 'localStorage',
    storeAuthStateInCookie: true
  };
  request = {
    scopes: ["user.read"]
  };
  graph = {
    callAfterInit: false,
    endpoints: {
      profile: '/me'
    },
    baseUrl: 'https://graph.microsoft.com/v1.0',
    onResponse: response => {}
  };

  constructor(options) {
    this.options = options;

    if (!options.auth.clientId) {
      throw new Error('auth.clientId is required');
    }

    this.auth = Object.assign(this.auth, options.auth);
    this.cache = Object.assign(this.cache, options.cache);
    this.request = Object.assign(this.request, options.request);
    this.graph = Object.assign(this.graph, options.graph);
    this.lib = new _UserAgentApplicationExtended.UserAgentApplicationExtended({
      auth: {
        clientId: this.auth.clientId,
        authority: this.auth.authority || `https://${this.auth.tenantName}/${this.auth.tenantId}`,
        validateAuthority: this.auth.validateAuthority,
        redirectUri: this.auth.redirectUri,
        postLogoutRedirectUri: this.auth.postLogoutRedirectUri,
        navigateToLoginRequestUrl: this.auth.navigateToLoginRequestUrl
      },
      cache: this.cache,
      system: options.system
    });
    this.getSavedCallbacks();
    this.executeCallbacks(); // Register Callbacks for redirect flow

    this.lib.handleRedirectCallback((error, response) => {
      if (!this.isAuthenticated()) {
        this.saveCallback('auth.onAuthentication', error, response);
      } else {
        this.acquireToken();
      }
    });

    if (this.auth.requireAuthOnInitialize) {
      this.signIn();
    }

    this.data.isAuthenticated = this.isAuthenticated();

    if (this.data.isAuthenticated) {
      this.data.user = this.lib.getAccount();
      this.acquireToken().then(() => {
        if (this.graph.callAfterInit) {
          this.initialMSGraphCall();
        }
      });
    }

    this.getStoredCustomData();
  }

  signIn() {
    if (!this.lib.isCallback(window.location.hash) && !this.lib.getAccount()) {
      // request can be used for login or token request, however in more complex situations this can have diverging options
      this.lib.loginRedirect(this.request);
    }
  }

  async signOut() {
    if (this.options.auth.beforeSignOut) {
      await this.options.auth.beforeSignOut(this);
    }

    this.lib.logout();
  }

  isAuthenticated() {
    return !this.lib.isCallback(window.location.hash) && !!this.lib.getAccount();
  }

  async acquireToken(request = this.request, retries = 0) {
    try {
      //Always start with acquireTokenSilent to obtain a token in the signed in user from cache
      const response = await this.lib.acquireTokenSilent(request);
      this.handleTokenResponse(null, response);
      return response;
    } catch (error) {
      // Upon acquireTokenSilent failure (due to consent or interaction or login required ONLY)
      // Call acquireTokenRedirect
      if (this.requiresInteraction(error.errorCode)) {
        this.lib.acquireTokenRedirect(request);
      } else if (retries > 0) {
        return await new Promise(resolve => {
          setTimeout(async () => {
            const res = await this.acquireToken(request, retries - 1);
            resolve(res);
          }, 60 * 1000);
        });
      }

      return false;
    }
  }

  handleTokenResponse(error, response) {
    if (error) {
      this.saveCallback('auth.onToken', error, null);
      return;
    }

    let setCallback = false;

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
  }

  setToken(tokenType, token, expiresOn, scopes) {
    const expirationOffset = this.lib.config.system.tokenRenewalOffsetSeconds * 1000;
    const expiration = expiresOn.getTime() - new Date().getTime() - expirationOffset;

    if (expiration >= 0) {
      this.data[tokenType] = token;
    }

    if (this.tokenExpirationTimers[tokenType]) clearTimeout(this.tokenExpirationTimers[tokenType]);
    this.tokenExpirationTimers[tokenType] = window.setTimeout(async () => {
      if (this.auth.autoRefreshToken) {
        await this.acquireToken({
          scopes
        }, 3);
      } else {
        this.data[tokenType] = '';
      }
    }, expiration);
  }

  requiresInteraction(errorCode) {
    if (!errorCode || !errorCode.length) {
      return false;
    }

    return errorCode === "consent_required" || errorCode === "interaction_required" || errorCode === "login_required";
  } // MS GRAPH


  async initialMSGraphCall() {
    const {
      onResponse: callback
    } = this.graph;
    let initEndpoints = this.graph.endpoints;

    if (typeof initEndpoints === 'object' && !_lodash.default.isEmpty(initEndpoints)) {
      const resultsObj = {};
      const forcedIds = [];

      try {
        const endpoints = {};

        for (const id in initEndpoints) {
          endpoints[id] = this.getEndpointObject(initEndpoints[id]);

          if (endpoints[id].force) {
            forcedIds.push(id);
          }
        }

        let storedIds = [];
        let storedData = this.lib.store.getItem(`msal.msgraph-${this.data.accessToken}`);

        if (storedData) {
          storedData = JSON.parse(storedData);
          storedIds = Object.keys(storedData);
          Object.assign(resultsObj, storedData);
        }

        const {
          singleRequests,
          batchRequests
        } = this.categorizeRequests(endpoints, _lodash.default.difference(storedIds, forcedIds));
        const singlePromises = singleRequests.map(async endpoint => {
          const res = {};
          res[endpoint.id] = await this.msGraph(endpoint);
          return res;
        });
        const batchPromises = Object.keys(batchRequests).map(key => {
          const batchUrl = key === 'default' ? undefined : key;
          return this.msGraph(batchRequests[key], batchUrl);
        });
        const mixedResults = await Promise.all([...singlePromises, ...batchPromises]);
        mixedResults.map(res => {
          for (const key in res) {
            res[key] = res[key].body;
          }

          Object.assign(resultsObj, res);
        });
        const resultsToSave = { ...resultsObj
        };
        forcedIds.map(id => delete resultsToSave[id]);
        this.lib.store.setItem(`msal.msgraph-${this.data.accessToken}`, JSON.stringify(resultsToSave));
        this.data.graph = resultsObj;
      } catch (error) {
        console.error(error);
      }

      if (callback) this.saveCallback('graph.onResponse', this.data.graph);
    }
  }

  async msGraph(endpoints, batchUrl = undefined) {
    try {
      if (Array.isArray(endpoints)) {
        return await this.executeBatchRequest(endpoints, batchUrl);
      } else {
        return await this.executeSingleRequest(endpoints);
      }
    } catch (error) {
      throw error;
    }
  }

  async executeBatchRequest(endpoints, batchUrl = this.graph.baseUrl) {
    const requests = endpoints.map((endpoint, index) => this.createRequest(endpoint, index));
    const {
      data
    } = await _axios.default.request({
      url: `${batchUrl}/$batch`,
      method: 'POST',
      data: {
        requests: requests
      },
      headers: {
        Authorization: `Bearer ${this.data.accessToken}`
      },
      responseType: 'json'
    });
    let result = {};
    data.responses.map(response => {
      let key = response.id;
      delete response.id;
      return result[key] = response;
    }); // Format result

    const keys = Object.keys(result);
    const numKeys = keys.sort().filter((key, index) => {
      if (key.search('defaultID-') === 0) {
        key = key.replace('defaultID-', '');
      }

      return parseInt(key) === index;
    });

    if (numKeys.length === keys.length) {
      result = _lodash.default.values(result);
    }

    return result;
  }

  async executeSingleRequest(endpoint) {
    const request = this.createRequest(endpoint);

    if (request.url.search('http') !== 0) {
      request.url = this.graph.baseUrl + request.url;
    }

    const res = await _axios.default.request(_lodash.default.defaultsDeep(request, {
      url: request.url,
      method: request.method,
      responseType: 'json',
      headers: {
        Authorization: `Bearer ${this.data.accessToken}`
      }
    }));
    return {
      status: res.status,
      headers: res.headers,
      body: res.data
    };
  }

  createRequest(endpoint, index = 0) {
    const request = {
      url: '',
      method: 'GET',
      id: `defaultID-${index}`
    };
    endpoint = this.getEndpointObject(endpoint);

    if (endpoint.url) {
      Object.assign(request, endpoint);
    } else {
      throw {
        error: 'invalid endpoint',
        endpoint: endpoint
      };
    }

    return request;
  }

  categorizeRequests(endpoints, excludeIds) {
    let res = {
      singleRequests: [],
      batchRequests: {}
    };

    for (const key in endpoints) {
      const endpoint = {
        id: key,
        ...endpoints[key]
      };

      if (!_lodash.default.includes(excludeIds, key)) {
        if (endpoint.batchUrl) {
          const {
            batchUrl
          } = endpoint;
          delete endpoint.batchUrl;

          if (!res.batchRequests.hasOwnProperty(batchUrl)) {
            res.batchRequests[batchUrl] = [];
          }

          res.batchRequests[batchUrl].push(endpoint);
        } else {
          res.singleRequests.push(endpoint);
        }
      }
    }

    return res;
  }

  getEndpointObject(endpoint) {
    if (typeof endpoint === "string") {
      endpoint = {
        url: endpoint
      };
    }

    if (typeof endpoint === "object" && !endpoint.url) {
      throw new Error('invalid endpoint url');
    }

    return endpoint;
  } // CUSTOM DATA


  saveCustomData(key, data) {
    if (!this.data.custom.hasOwnProperty(key)) {
      this.data.custom[key] = null;
    }

    this.data.custom[key] = data;
    this.storeCustomData();
  }

  storeCustomData() {
    if (!_lodash.default.isEmpty(this.data.custom)) {
      this.lib.store.setItem('msal.custom', JSON.stringify(this.data.custom));
    } else {
      this.lib.store.removeItem('msal.custom');
    }
  }

  getStoredCustomData() {
    let customData = {};
    const customDataStr = this.lib.store.getItem('msal.custom');

    if (customDataStr) {
      customData = JSON.parse(customDataStr);
    }

    this.data.custom = customData;
  } // CALLBACKS


  saveCallback(callbackPath, ...args) {
    if (_lodash.default.get(this.options, callbackPath)) {
      const callbackQueueObject = {
        id: _lodash.default.uniqueId(`cb-${callbackPath}`),
        callback: callbackPath,
        arguments: args
      };

      _lodash.default.remove(this.callbackQueue, obj => obj.id === callbackQueueObject.id);

      this.callbackQueue.push(callbackQueueObject);
      this.storeCallbackQueue();
      this.executeCallbacks([callbackQueueObject]);
    }
  }

  getSavedCallbacks() {
    const callbackQueueStr = this.lib.store.getItem('msal.callbackqueue');

    if (callbackQueueStr) {
      this.callbackQueue = [...this.callbackQueue, ...JSON.parse(callbackQueueStr)];
    }
  }

  async executeCallbacks(callbacksToExec = this.callbackQueue) {
    if (callbacksToExec.length) {
      for (let i in callbacksToExec) {
        const cb = callbacksToExec[i];

        const callback = _lodash.default.get(this.options, cb.callback);

        try {
          await callback(this, ...cb.arguments);

          _lodash.default.remove(this.callbackQueue, function (currentCb) {
            return cb.id === currentCb.id;
          });

          this.storeCallbackQueue();
        } catch (e) {
          console.warn(`Callback '${cb.id}' failed with error: `, e.message);
        }
      }
    }
  }

  storeCallbackQueue() {
    if (this.callbackQueue.length) {
      this.lib.store.setItem('msal.callbackqueue', JSON.stringify(this.callbackQueue));
    } else {
      this.lib.store.removeItem('msal.callbackqueue');
    }
  }

}

exports.MSAL = MSAL;