"use strict";

function ServiceConsumer(service) {
    this._service = service;
    this._serviceMethod = '';
}

ServiceConsumer.prototype._execute = function _execute(soapObj, callback) {
    this._service[this._serviceMethod](soapObj, callback);
};

module.exports = ServiceConsumer;