"use strict";

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.UserAgentApplicationExtended = void 0;

var _msal = require("msal");

class UserAgentApplicationExtended extends _msal.UserAgentApplication {
  store = {};

  constructor(configuration) {
    super(configuration);
    this.store = this.cacheStorage;
  }

}

exports.UserAgentApplicationExtended = UserAgentApplicationExtended;