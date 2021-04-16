package com.leidos.tamp.beans;

import org.bson.codecs.pojo.annotations.BsonProperty;
import org.bson.types.ObjectId;

public class EquipmentPM {
	ObjectId id;
	@BsonProperty(value="MakeModel")
	String makeModel;
	@BsonProperty(value="Name")
	String name;
	@BsonProperty(value="Periodicity")
	String periodicity;
	@BsonProperty(value="AllowedSlack")
	Integer allowedSlack;
	@BsonProperty(value="LaborInitial")
	Double laborInitial;
	@BsonProperty(value="LaborWait")
	Double laborWait;
	@BsonProperty(value="LaborFinal")
	Double laborFinal;
	@BsonProperty(value="ConsumablesCost")
	Double consumablesCost;
	@BsonProperty(value="TechnicianCount")
	Integer technicianCount;
	public ObjectId getId() {
		return id;
	}
	public void setId(ObjectId id) {
		this.id = id;
	}
	public String getMakeModel() {
		return makeModel;
	}
	public void setMakeModel(String makeModel) {
		this.makeModel = makeModel;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getPeriodicity() {
		return periodicity;
	}
	public void setPeriodicity(String periodicity) {
		this.periodicity = periodicity;
	}
	public Integer getAllowedSlack() {
		return allowedSlack;
	}
	public void setAllowedSlack(Integer allowedSlack) {
		this.allowedSlack = allowedSlack;
	}
	public Double getLaborInitial() {
		return laborInitial;
	}
	public void setLaborInitial(Double laborInitial) {
		this.laborInitial = laborInitial;
	}
	public Double getLaborWait() {
		return laborWait;
	}
	public void setLaborWait(Double laborWait) {
		this.laborWait = laborWait;
	}
	public Double getLaborFinal() {
		return laborFinal;
	}
	public void setLaborFinal(Double laborFinal) {
		this.laborFinal = laborFinal;
	}
	public Double getConsumablesCost() {
		return consumablesCost;
	}
	public void setConsumablesCost(Double consumablesCost) {
		this.consumablesCost = consumablesCost;
	}
	public Integer getTechnicianCount() {
		return technicianCount;
	}
	public void setTechnicianCount(Integer technicianCount) {
		this.technicianCount = technicianCount;
	}
	
}
