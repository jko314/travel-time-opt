package com.leidos.tamp.beans;

import org.bson.codecs.pojo.annotations.BsonProperty;
import org.bson.types.ObjectId;

public class CMRate {
	ObjectId id;
	@BsonProperty(value="include")
	String include;
	@BsonProperty(value="Index")
	String index;
	@BsonProperty(value="Source")
	String source;
	@BsonProperty(value="ModelNum")
	String modelNum;
	@BsonProperty(value="Category")
	String category;
	@BsonProperty(value="Name")
	String name;
	@BsonProperty(value="WorkType")
	String workType;
	@BsonProperty(value="Mean")
	String mean;
	@BsonProperty(value="StndDev")
	String stndDev;
	@BsonProperty(value="Description")
	String description;
	@BsonProperty(value="CMTime")
	String cmTime;
	@BsonProperty(value="CMTotalTime")
	String cmTotalTime;
	public ObjectId getId() {
		return id;
	}
	public void setId(ObjectId id) {
		this.id = id;
	}
	public String getInclude() {
		return include;
	}
	public void setInclude(String include) {
		this.include = include;
	}
	public String getIndex() {
		return index;
	}
	public void setIndex(String index) {
		this.index = index;
	}
	public String getSource() {
		return source;
	}
	public void setSource(String source) {
		this.source = source;
	}
	public String getModelNum() {
		return modelNum;
	}
	public void setModelNum(String modelNum) {
		this.modelNum = modelNum;
	}
	public String getCategory() {
		return category;
	}
	public void setCategory(String category) {
		this.category = category;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getWorkType() {
		return workType;
	}
	public void setWorkType(String workType) {
		this.workType = workType;
	}
	public String getMean() {
		return mean;
	}
	public void setMean(String mean) {
		this.mean = mean;
	}
	public String getStndDev() {
		return stndDev;
	}
	public void setStndDev(String stndDev) {
		this.stndDev = stndDev;
	}
	public String getDescription() {
		return description;
	}
	public void setDescription(String description) {
		this.description = description;
	}
	public String getCmTime() {
		return cmTime;
	}
	public void setCmTime(String cmTime) {
		this.cmTime = cmTime;
	}
	public String getCmTotalTime() {
		return cmTotalTime;
	}
	public void setCmTotalTime(String cmTotalTime) {
		this.cmTotalTime = cmTotalTime;
	}	

}
