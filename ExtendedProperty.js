'use strict';

var uuid = require('node-uuid'),
	MapiPropertyType = require('./MapiPropertyType');

function ExtendedProperty(propertyName, value, propertyType, distinguishedPropertySetId, guid, tag) {
	if(tag){
		this._tag = tag;
	}else{
		if (distinguishedPropertySetId) {
			this._distinguishedPropertySetId = distinguishedPropertySetId;
		}if (guid) {
			this._uuid = guid;
		} else {
			this._uuid = uuid.v4();
		}

		this._name = propertyName || '';
		this._value = value || '';
	}
	this._type = propertyType || MapiPropertyType.STRING;
}

ExtendedProperty.prototype.setName = function (name) {
	this._name = name;
	return this;
};

ExtendedProperty.prototype.getName = function (name) {
	return this._name;
};

ExtendedProperty.prototype.setValue = function (value) {
	this._value = value;
	return this;
};

ExtendedProperty.prototype.getValue = function () {
	return this._value;
};

ExtendedProperty.prototype.setType = function (type) {
	this._type = type;
	return this;
};

ExtendedProperty.prototype.getType = function () {
	return this._type;
};

ExtendedProperty.prototype.getUUID = function () {
	return this._uuid;
};

ExtendedProperty.prototype.toJSON = function () {
	var serialized = { ExtendedFieldURI: {} };
	
	serialized.ExtendedFieldURI.attributes = this.getAttributes();

	if (this._value) {
		serialized.Value = this._value;
	}
	return serialized;
};

ExtendedProperty.prototype.getAttributes = function () {
	var attributes = {};

	if(this._tag){
		attributes.PropertyTag = this._tag
	}else{
		if (this._uuid) {
			attributes.PropertySetId = this._uuid;
		} else if (this._distinguishedPropertySetId) {
			attributes.DistinguishedPropertySetId = this._distinguishedPropertySetId;
		}
		attributes.PropertyName = this._name;
	}
	attributes.PropertyType = this._type;
	
	return attributes;
};

module.exports = ExtendedProperty;