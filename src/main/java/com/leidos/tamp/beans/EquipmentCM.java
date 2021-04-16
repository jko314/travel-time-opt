package com.leidos.tamp.beans;

import org.bson.codecs.pojo.annotations.BsonProperty;
import org.bson.types.ObjectId;

public class EquipmentCM {
	ObjectId id;
	@BsonProperty(value="MakeModel")
	String makeModel;
	@BsonProperty(value="Name")
	String name;
	@BsonProperty(value="Frequency")
	Double frequency;
	@BsonProperty(value="CMTime")
	Double cmTime;
	@BsonProperty(value="CMStndDev")
	Double cmsStndDev;
	@BsonProperty(value="CMMin")
	Double cmMin;
	@BsonProperty(value="CMMax")
	Double cmMax;
	@BsonProperty(value="PartsCost")
	Double partsCost;
	@BsonProperty(value="PartsTime")
	Double partsTime;
	@BsonProperty(value="PartsStndDev")
	Double partsStndDev;
	@BsonProperty(value="PartsMin")
	Double partsMin;
	@BsonProperty(value="PartsMax")
	Double partsMax;
	@BsonProperty(value="ConsumablesCost")
	Double consumablesCost;
	@BsonProperty(value="TechCount")
	Integer techCount;
	@BsonProperty(value="Index")
	Integer index;
	@BsonProperty(value="Source")
	String source;
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
	public Double getFrequency() {
		return frequency;
	}
	public void setFrequency(Double frequency) {
		this.frequency = frequency;
	}
	public Double getCmTime() {
		return cmTime;
	}
	public void setCmTime(Double cmTime) {
		this.cmTime = cmTime;
	}
	public Double getCmsStndDev() {
		return cmsStndDev;
	}
	public void setCmsStndDev(Double cmsStndDev) {
		this.cmsStndDev = cmsStndDev;
	}
	public Double getCmMin() {
		return cmMin;
	}
	public void setCmMin(Double cmMin) {
		this.cmMin = cmMin;
	}
	public Double getCmMax() {
		return cmMax;
	}
	public void setCmMax(Double cmMax) {
		this.cmMax = cmMax;
	}
	public Double getPartsCost() {
		return partsCost;
	}
	public void setPartsCost(Double partsCost) {
		this.partsCost = partsCost;
	}
	public Double getPartsTime() {
		return partsTime;
	}
	public void setPartsTime(Double partsTime) {
		this.partsTime = partsTime;
	}
	public Double getPartsStndDev() {
		return partsStndDev;
	}
	public void setPartsStndDev(Double partsStndDev) {
		this.partsStndDev = partsStndDev;
	}
	public Double getPartsMin() {
		return partsMin;
	}
	public void setPartsMin(Double partsMin) {
		this.partsMin = partsMin;
	}
	public Double getPartsMax() {
		return partsMax;
	}
	public void setPartsMax(Double partsMax) {
		this.partsMax = partsMax;
	}
	public Double getConsumablesCost() {
		return consumablesCost;
	}
	public void setConsumablesCost(Double consumablesCost) {
		this.consumablesCost = consumablesCost;
	}
	public Integer getTechCount() {
		return techCount;
	}
	public void setTechCount(Integer techCount) {
		this.techCount = techCount;
	}
	public Integer getIndex() {
		return index;
	}
	public void setIndex(Integer index) {
		this.index = index;
	}
	public String getSource() {
		return source;
	}
	public void setSource(String source) {
		this.source = source;
	}
}
