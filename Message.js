"use strict";

var SoapRequest = require('./SoapRequest');

function Message(service) {
    this._service = service;
    this._recipients = [];
	this._id = null;
	this._syncState = null;
	this._defaultFields = [
		'item:Attachments'		
	];
	this._additionalFields = [];
	this._attachments = [];	
}

Message.prototype.Bind = function (id) {
	this._id = id;
};

Message.prototype.setSyncState = function (syncState) {
	this._syncState = syncState;
};

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

Message.prototype.sendItemAction = function _sendSaveAction (callback) {
    var soapRequest = new SoapRequest('SendItem', {
    	attributes: {
    		SaveItemToFolder: true
    	},
        ItemIds: {
            ItemId: {
		        attributes: {
		            Id: this.ItemId,
		            ChangeKey: this.ChangeKey		     
		        }
            }
        }
    });
    this._service.execute(soapRequest, callback);
};

Message.prototype._sendSaveAction = function _sendSaveAction (callback) {
    var soapRequest = new SoapRequest('CreateItem', {
        attributes: {
            MessageDisposition: this._messageDisposition
        },
        Items: {
            Message: {
            	Attachments: this._attachments.map(function (attachment) {
            		return {
            			FileAttachment: {
            				Name: attachment.Name,
            				IsInline: attachment.IsInline
            			}
            		};
            	}),
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

Message.prototype.BindAttachment = function (attachment) {
	this._attachments.push(attachment);
	return this;	
};

Message.prototype.GetAttachmentInline = function (attachment) {

	return [
		{
			Name: this._attachments[0].name,
			ContentId: this._attachments[0].name,
			IsInline: true,
			IsContactPhoto: false,
			Content: this._attachments[0].content
		},
		{
			Name: this._attachments[1].name,
			ContentId: this._attachments[1].name,
			IsInline: true,
			IsContactPhoto: false,
			Content: this._attachments[1].content
		}
	];
};

/*this._attachments.map(function (attachment) {
        	if (attachment.isInline) {
	        	return {
	        		FileAttachment: {
	        			Name: attachment.name,
	        			ContentId: attachment.ContentId,
	        			IsInline: true,
	        			IsContactPhoto: false,
	        			Content: attachment.content
	        		}
	        	};
        	} else {
	        	return {
	        		FileAttachment: {
	        			Name: attachment.name,
	        			Content: attachment.content
	        		}
	        	};
        	}
        })*/

Message.prototype.CreateAttachment = function (callback) {
	var soapRequest = new SoapRequest('CreateAttachment', {
        ParentItemId: {
	        attributes: {
	            Id: this.ItemId
	        }
        },
        Attachments: {
        	FileAttachment: this.GetAttachmentInline()
        } 
    });
    this._service.execute(soapRequest, callback);
};

Message.prototype.GetAttachments = function (callback) {
	var attachmentsCollection = new AttachmentsCollection(this._service);
	attachmentsCollection.BindToItems(this._attachments);
	
	attachmentsCollection.Load(function (err, results) {
		if (!err) {
			this.Attachments = results;
		}

		callback(err, results);
	}.bind(this));
};

function MessageCollection(service) {
	this._service = service;
	this._defaultFields = [
		'item:Attachments'
	];
}

MessageCollection.prototype.BindToItems = function (items) {
	this._items = items;
}

MessageCollection.prototype.Load = function (callback) {
	var soapRequest = new SoapRequest('GetItem', {
		ItemShape: {	
			BaseShape: 'Default',
			IncludeMimeContent: true,
			AdditionalProperties: {
				FieldURI: this._defaultFields.map(function (field) {
					return { attributes: { FieldURI: field } };
				})
			}
		},
		ItemIds: {
			ItemId: this._items.map(function (item) { 
					if (typeof item == 'string') {
						return {
							attributes: {
								Id: item,
								ChangeKey: ''
							}
						};
					}
					return {
						attributes: {
							Id: item.Id,
							ChangeKey: item.ChangeKey || ''
						}
					}
				})
		}
	});
	this._service.execute(soapRequest, callback);
};

function AttachmentCollection(service) {
	this._service = service;
	this._items = [];
}

AttachmentCollection.prototype.BindToItems = function (items) {
	this._items = items;
};

AttachmentCollection.prototype.Load = function (callback) {
	var soapRequest = new SoapRequest('GetAttachment', {
		AttachmentIds: {
			AttachmentId: this._items.map(function (attachment) {
				return { attributes: attachment };
			})
		}
	});
	this._service.execute(soapRequest, callback);
}


Message.prototype.MessageCollection = MessageCollection;
Message.prototype.AttachmentCollection  = AttachmentCollection;

module.exports = Message;

