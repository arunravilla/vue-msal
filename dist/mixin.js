"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.mixin = void 0;
exports.mixin = {
    data: function () {
        return {
            msal: (this.$msal) ? this.$msal.data : {}
        };
    },
    created: function () {
        var _this = this;
        this.$watch('$msal.data', function (value) { _this.msal = value; }, { deep: true });
    }
};
//# sourceMappingURL=mixin.js.map