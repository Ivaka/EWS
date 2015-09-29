"use strict";

var soap = require('soap'),
    path = require('path'),
    Message = require('./Message'),
    Folder = require('./Folder'),
    SoapRequest = require('./SoapRequest'),
	ExtendedProperty = require('./ExtendedProperty'),
	MapiPropertyType = require('./MapiPropertyType'),
	DistinguishedPropertySet = require('./DistinguishedPropertySet');

var noop = function () {};
var noopThrows = function (err) { if (err) {throw new Error(err)} };
var _toString = Object.prototype.toString;
var NO_ERROR = 'NoError';

function EWS(config) {
    this._username = config.domain + '\\' + config.username;
    this._password = config.password;
    this._endpoint = 'https://' + path.join(config.url, 'EWS/Exchange.asmx');
    this._exchangeVersoin = config.version || '2010';

    /**
 *  *      * TODO: Figure out a way to request the Services.wsdl, messages.xsd & types.xsd directly from the endpoint
 *   *           * */

    this._services = path.join(__dirname, '/Services/Services.wsdl');

    this._soapClientOptions = {
        ignoredNamespaces: {
            namespaces: [],
            override: true
        }
    };
}

EWS.prototype.connect = function connect (callback) {
    if (!callback || typeof callback !== 'function'){
        callback = noop;
    }
    soap.createClient(this._services, this._soapClientOptions, function (err, client) {
        if (err) {
            return callback(err, false);
        }

        this._client = client;
        this._client.setSecurity(new soap.BasicAuthSecurity(this._username, this._password));
        /**
 *  *          * TODO: Unhardcode exchage version in header
 *   *                   * */
        this._client.addSoapHeader('<t:RequestServerVersion Version="Exchange2010_SP2" />');

        callback(null, true);
    }.bind(this), this._endpoint);

    return this;
};

EWS.prototype.CreateItem = function (soapRequest, callback) {
    if (!callback || typeof callback !== 'function'){
        callback = noop;
    }
    this._client.CreateItem(soapRequest, function (err, results) {
        if (err) {
            return callback(err, null);
        }

        if (results.ResponseMessages.CreateItemResponseMessage.ResponseCode == NO_ERROR) {
            return callback(null, results.ResponseMessages.CreateItemResponseMessage);
        }
    
	    return callback(new Error(results.ResponseMessages.CreateItemResponseMessage.ResponseCode), results);
    });
};

EWS.prototype.CreateAttachment = function (soapRequest, callback) {
    if (!callback || typeof callback !== 'function'){
        callback = noop;
    }
    this._client.CreateAttachment(soapRequest, function (err, results) {
        if (err) {
            return callback(err, null);
        }

        if (results.ResponseMessages.CreateAttachmentResponseMessage.ResponseCode == NO_ERROR) {
            return callback(null, results.ResponseMessages.CreateAttachmentResponseMessage);
        }

        return callback(new Error(results.ResponseMessages.CreateAttachmentResponseMessage.ResponseCode), results);
    });
};

EWS.prototype.SendItem = function (soapRequest, callback) {
    if (!callback || typeof callback !== 'function'){
        callback = noop;
    }
    this._client.SendItem(soapRequest, function (err, results) {
        if (err) {
            return callback(err, null);
        }

        if (results.ResponseMessages.SendItemResponseMessage.ResponseCode == NO_ERROR) {
            return callback(null, results.ResponseMessages.SendItemResponseMessage);
        }

        return callback(new Error(results.ResponseMessages.SendItemResponseMessage.ResponseCode), results);
    });
};

EWS.prototype.FindItem = function (soapRequest, callback) {
    if (!callback || typeof callback !== 'function'){
        callback = noop;
    }
    this._client.FindItem(soapRequest, function (err, results) {
        if (err) {
            return callback(err, null);
        }

        if (results.ResponseMessages.FindItemResponseMessage.ResponseCode == NO_ERROR) {
            return callback(null, results.ResponseMessages.FindItemResponseMessage);
        }

        return callback(new Error(results.ResponseMessages.FindItemResponseMessage.ResponseCode), results);
    });
};

EWS.prototype.SyncFolderItems = function SyncFolderItems(soapRequest, callback) {
    if (!callback || typeof callback !== 'function'){
        callback = noop;
    }

    this._client.SyncFolderItems(soapRequest, function (err, results) {
        if (err) {
            return callback(err, null);
        }

        if (results.ResponseMessages.SyncFolderItemsResponseMessage.ResponseCode == NO_ERROR) {
            return callback(null, results.ResponseMessages.SyncFolderItemsResponseMessage);
        }

        return callback(new Error(results.ResponseMessages.SyncFolderItemsResponseMessage.ResponseCode), results);
    });
};

EWS.prototype.GetItem = function (soapRequest, callback) {
    if (!callback || typeof callback !== 'function'){
        callback = noop;
    }

    this._client.GetItem(soapRequest, function (err, results) {
        if (err) {
            return callback(err, null);
        }

        var errRes, noErrRes;

        if (_toString.call(results.ResponseMessages.GetItemResponseMessage) === '[object Array]' ) {
            errRes = results.ResponseMessages.GetItemResponseMessage.filter(function (message) {
                return message.ResponseCode !== NO_ERROR;
            });

            noErrRes = results.ResponseMessages.GetItemResponseMessage.filter(function (message) {
                return message.ResponseCode == NO_ERROR;
            });
        } else {
            errRes = results.ResponseMessages.GetItemResponseMessage.ResponseCode == NO_ERROR ? [] : [results.ResponseMessages.GetItemResponseMessage];
            noErrRes = results.ResponseMessages.GetItemResponseMessage.ResponseCode !== NO_ERROR ? [] : [results.ResponseMessages.GetItemResponseMessage];
        }


        if (!errRes.length) {
            return callback(err, noErrRes);
        } else {
            return callback(err || new Error('Failed to fetch some items'), noErrRes, errRes);
        }
    });
}

EWS.prototype.SyncFolderHierarchy = function (soapRequest, callback) {
	if (!callback || typeof callback !== 'function') {
		callback = noop;
	}

	this._client.SyncFolderHierarchy(soapRequest, function (err, results) {
		if (err) {
			return callback(err, null);
		}

		if (results.ResponseMessages.SyncFolderHierarchyResponseMessage.ResponseCode === NO_ERROR) {
			return callback(null, results);
		}

		return callback(new Error(results.ResponseMessages.SyncFolderHierarchyResponseMessage.ResponseCode), results);
	});
};

EWS.prototype.GetAttachment = function (soapRequest, callback) {
	if (!callback || typeof callback !== 'function') {
		callback = noop;
	}
	
	this._client.GetAttachment(soapRequest, function (err, results) {
		if (err) {
			return callback(err, null);
		}
		
		if (results.ResponseMessages.GetAttachmentResponseMessage.ResponseCode === NO_ERROR) {
			return callback(null, results);
		}

		return callback(new Error(results.ResponseMessages.GetAttachmentResponseMessage.ResponseCode), results);
	})
}

EWS.prototype.execute = function execute(soapRequest, callback) {
    this[soapRequest.getMethod()](soapRequest.getRequest(), callback);
};

EWS.prototype.Message = function () {
    return new Message(this);
};

EWS.prototype.MessageCollection = function () {
	return new Message.prototype.MessageCollection(this);
};

EWS.prototype.AttachmentCollection = function () {
	return new Message.prototype.AttachmentCollection(this);
};

EWS.prototype.Folder = function () {
    return new Folder(this);
};

EWS.prototype.ExtendedProperty = ExtendedProperty;
EWS.prototype.MapiPropertyType = MapiPropertyType;
EWS.prototype.DistinguishedPropertySet = DistinguishedPropertySet;

module.exports = EWS;

