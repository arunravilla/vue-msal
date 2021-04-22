'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = exports.msalMixin = void 0;

var _main = require("./src/main");

var _mixin = require("./mixin");

const msalMixin = _mixin.mixin;
exports.msalMixin = msalMixin;

class msalPlugin {
  static install(Vue, options) {
    Vue.prototype.$msal = new msalPlugin(options, Vue);
  }

  constructor(options, Vue = undefined) {
    const msal = new _main.MSAL(options);

    if (Vue && options.framework && options.framework.globalMixin) {
      Vue.mixin(_mixin.mixin);
    }

    const exposed = {
      data: msal.data,

      signIn() {
        msal.signIn();
      },

      async signOut() {
        await msal.signOut();
      },

      isAuthenticated() {
        return msal.isAuthenticated();
      },

      async acquireToken(request, retries = 0) {
        return await msal.acquireToken(request, retries);
      },

      async msGraph(endpoints, batchUrl) {
        return await msal.msGraph(endpoints, batchUrl);
      },

      saveCustomData(key, data) {
        msal.saveCustomData(key, data);
      }

    };
    return exposed;
  }

}

exports.default = msalPlugin;