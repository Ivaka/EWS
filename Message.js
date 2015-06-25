
"use strict";

function Message(service) {
    this._service = service;
    this._serviceMethod = 'CreateItem';
    this._recipients = [];
}

Message.prototype.SendAndSaveCopy = function SendAndSaveCopy (callback) {
    this._messageDisposition = 'SendAndSaveCopy';
    this._sendSaveAction(callback);
};

Message.prototype.Send = function Send (callback) {
    this._messageDisposition = 'SendOnly';
    this._sendSaveAction(callback);
};

Message.prototype.Save = function Save (callback) {
    this._messageDisposition = 'SaveOnly';
    this._sendSaveAction(callback);
};

Message.prototype.addRecipient = function addRecipient(recipient) {
    this._recipients.push(recipient);
};

Message.prototype._sendSaveAction = function _sendSaveAction (callback) {
    var soapObj = this._buildSoapRequest();
    this._execute(soapObj, callback);
};

Message.prototype._buildSoapRequest = function _buildSoapRequest() {
    return {
            attributes: {
                MessageDisposition: this._messageDisposition
            },
            Items: {
                Message: {
                    ItemClass: 'IPM.Note',
                    Subject: this.Subject,
                    Body: {
                        attributes: {
                            BodyType: 'HTML'
                        },
                        $value: this.Body
                    },
                    ToRecipients: this._recipients.map(function (recipient) {
                        return {
                            Mailbox: {
                                EmailAddress: recipient
                            }
                        }
                    }),
                    From: {
                        Mailbox: {
                            EmailAddress: this.From
                        }
                    }
                }
            }
        };
};

Message.prototype._execute = function (soapObj, callback) {
    this._service.CreateItem(soapObj, callback);
};

module.exports = Message;