
"use strict";

var SoapRequest = require('./SoapRequest');

function Message(service) {
    this._service = service;
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
    var soapRequest = new SoapRequest('CreateItem', {
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
    });
    this._service.execute(soapRequest, callback);
};

module.exports = Message;