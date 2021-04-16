package com.leidos.tamp.beans;

import org.bson.codecs.pojo.annotations.BsonProperty;
import org.bson.types.ObjectId;

public class EquipmentModel {
	ObjectId id;
	@BsonProperty(value="EquipmentType")
	String equipmentType;
	@BsonProperty(value="Manufacturer")
	String manufacturer;
	@BsonProperty(value="Model")
	String model;
	@BsonProperty(value="OEM")
	String oem;
	@BsonProperty(value="Type")
	String type;
	public ObjectId getId() {
		return id;
	}
	public void setId(ObjectId id) {
		this.id = id;
	}
	public String getEquipmentType() {
		return equipmentType;
	}
	public void setEquipmentType(String equipmentType) {
		this.equipmentType = equipmentType;
	}
	public String getManufacturer() {
		return manufacturer;
	}
	public void setManufacturer(String manufacturer) {
		this.manufacturer = manufacturer;
	}
	public String getModel() {
		return model;
	}
	public void setModel(String model) {
		this.model = model;
	}
	public String getOem() {
		return oem;
	}
	public void setOem(String oem) {
		this.oem = oem;
	}
	public String getType() {
		return type;
	}
	public void setType(String type) {
		this.type = type;
	}

}
