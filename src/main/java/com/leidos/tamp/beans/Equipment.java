package com.leidos.tamp.beans;

import org.bson.codecs.pojo.annotations.BsonProperty;
import org.bson.types.ObjectId;

public class Equipment {
	ObjectId id;
	@BsonProperty(value="EquipmentID")
	String equipmentId;
	@BsonProperty(value="MakeModel")
	String makeModel;
	@BsonProperty(value="Airport")
	String airport;
	public ObjectId getId() {
		return id;
	}
	public void setId(ObjectId id) {
		this.id = id;
	}
	public String getEquipmentId() {
		return equipmentId;
	}
	public void setEquipmentId(String equipmentId) {
		this.equipmentId = equipmentId;
	}
	public String getMakeModel() {
		return makeModel;
	}
	public void setMakeModel(String makeModel) {
		this.makeModel = makeModel;
	}
	public String getAirport() {
		return airport;
	}
	public void setAirport(String airport) {
		this.airport = airport;
	}

}
