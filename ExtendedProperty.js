'use strict';

var uuid = require('node-uuid'),
	MapiPropertyType = require('./MapiPropertyType');

function ExtendedProperty(propertyName, value, propertyType) {
	this._uuid = uuid.v4();
	this._name = propertyName || '';
	this._type = propertyType || MapiPropertyType.STRING;
	this._value = value || '';
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
	return thus._uuid;
};

ExtendedProperty.prototype.toJSON = function () {
	return {
			ExtendedFieldURI: {
				attributes: {
					PropertySetId: this._uuid,
					PropertyName: this._name,
					PropertyType: this._type
				}
			},
			Value: this._value
	};
};

module.exports = ExtendedProperty;
