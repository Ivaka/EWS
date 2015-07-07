"use strict";

var SoapRequest = require('./SoapRequest');

function _additionalFieldsMap(field) {
    return {
            attributes: {
                FieldURI: field
            }
        };
}

/**
 * TODO: Add method to append additional fields
 * TODO: Add method to set override = true/false for additional fields
 * */

function Folder(service) {
    this._service = service;
    this._defaultFields = [
        'item:ItemId',
        'item:DateTimeCreated',
        'item:DateTimeSent',
        'item:HasAttachments',
        'item:Size',
        'message:From',
        'message:IsRead',
        'item:Importance',
        'item:Subject',
        'item:DateTimeReceived'
    ];

    this._additionalFields = [

    ];
    this._overrideFields = false;
    this._traversal = 'Shallow';
    this._baseShape = 'IdOnly';
    this._basePoint = 'Beginning';
    this._offset = 0;
    this._maxEntriesReturned = 10;
    this._id = 'inbox';
    this._syncState = null;
}

Folder.prototype.setTraversal = function setTraversal(traversal) {
    this._traversal = traversal;
    return this;
};

Folder.prototype.getTraversal = function getTraversal() {
    return this._traversal;
};

Folder.prototype.getBaseshape = function getBaseshape() {
    return this._baseShape;
};

Folder.prototype.setBaseshape = function getBaseshape(baseShape) {
    this._baseShape = baseShape;
    return this;
};

Folder.prototype.Bind = function Bind(id) {
    this._id = id;
    return this;
};

Folder.prototype.setMaxEntriesReturned = function setMaxEntriesReturned(max) {
    this._maxEntriesReturned = max;
    return this;
};

Folder.prototype.getMaxEntriesReturned = function getMaxEntriesReturned() {
    return this._maxEntriesReturned;
};

Folder.prototype.setSyncState = function setSyncState(syncState) {
    this._syncState = syncState;
    return this;
};

Folder.prototype.getSyncState = function getSyncState() {
    return this._syncState;
};

Folder.prototype.FindItems = function (callback) {
    var soapRequest = new SoapRequest('FindItem',
        {
            attributes: {
                Traversal: this._traversal
            },
            ItemShape: {
                BaseShape: this._baseShape,
                AdditionalProperties: {
                    FieldURI: this._overrideFields ? this._additionalFields.map(_additionalFieldsMap) : this._additionalFields.concat(this._defaultFields).map(_additionalFieldsMap)
                }
            },
            IndexedPageItemView: {
                attributes: {
                    BasePoint: this._basePoint,
                    Offset: this._offset,
                    MaxEntriesReturned: this._maxEntriesReturned
                }
            },
            ParentFolderIds: {
                DistinguishedFolderId: {
                    attributes: {
                        Id: this._id
                    }
                }
            }
        });
    this._service.execute(soapRequest, callback);
};

Folder.prototype.SyncFolderItems = function (callback) {
    var pureSoapObject = {
            ItemShape: {
                BaseShape: this._baseShape
            },
            SyncFolderId: {
                DistinguishedFolderId: {
                    attributes: {
                        Id: this._id
                    }
                }
            }
        };
    if (this._syncState) {
        pureSoapObject.SyncState = this._syncState;
    }

    pureSoapObject.MaxChangesReturned = this._maxEntriesReturned;
    pureSoapObject.SyncScope = 'NormalItems';

    var soapRequest = new SoapRequest('SyncFolderItems', pureSoapObject);

    this._service.execute(soapRequest, function (err, results) {
        if (err) {
            return callback(err);
        }
        if (results.ResponseCode === 'NoError') {
            this._syncState = results.SyncState;
        }

        return callback(err, results);
    }.bind(this));
};

module.exports = Folder;