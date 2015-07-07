
"use strict";

function SoapRequest(method, soapReqObj) {
    this._method = method;
    this._soapReqObj = soapReqObj;
}

SoapRequest.prototype.getMethod = function () {
    return this._method;
};

SoapRequest.prototype.getRequest = function () {
    return this._soapReqObj;
};

module.exports = SoapRequest;