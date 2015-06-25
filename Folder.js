"use strict";

var ServiceConsumer = require('./ServiceConsumer');

function _additionalFieldsMap(field) {
    return {
            attributes: {
                FieldURI: field
            }
        };
}

Folder.prototype = Object.create(ServiceConsumer.prototype);

function Folder(service) {
    ServiceConsumer.apply(this, arguments);
    this._serviceMethod = 'FindItem';
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

Folder.prototype._buildSoapRequest = function () {
    return {
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
    };
};

Folder.prototype.FindItems = function (callback) {
    var soapObj = this._buildSoapRequest();
    this._execute(soapObj, callback);
};

module.exports = Folder;