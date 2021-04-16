package com.leidos.tamp.db;

import com.leidos.tamp.beans.Airport;
import com.leidos.tamp.beans.AirportEquipment;
import com.leidos.tamp.beans.AirportServiceArea;
import com.leidos.tamp.beans.AirportTime;
import com.leidos.tamp.beans.AirportWithCargo;
import com.leidos.tamp.beans.CMRate;
import com.leidos.tamp.beans.Equipment;
import com.leidos.tamp.beans.EquipmentCM;
import com.leidos.tamp.beans.EquipmentModel;
import com.leidos.tamp.beans.EquipmentPM;
import com.leidos.tamp.beans.LaborRule;
import com.leidos.tamp.beans.ServiceArea;
import com.leidos.tamp.beans.ServiceAreaModelData;
import com.leidos.tamp.beans.TravelCost;
import com.leidos.tamp.beans.TravelMode;
import com.leidos.tamp.beans.TravelRule;
import com.mongodb.ConnectionString;
import com.mongodb.MongoClientSettings;
import com.mongodb.client.MongoClient;
import com.mongodb.client.MongoClients;
import com.mongodb.client.MongoDatabase;
import org.bson.codecs.configuration.CodecRegistry;
import org.bson.codecs.pojo.PojoCodecProvider;
import org.springframework.stereotype.Repository;

import java.util.logging.Level;
import java.util.logging.Logger;

import static org.bson.codecs.configuration.CodecRegistries.fromProviders;
import static org.bson.codecs.configuration.CodecRegistries.fromRegistries;

@Repository
public class IRADServiceDB {
	
	private static ServiceAreaModelData samData = new ServiceAreaModelData();

//	public static void main(String[] args) {
//		System.out.println("Starting to process IRAD");
//		System.out.println ("Retrieving IRAD JSON Data...");
//		getIRADData();
//	}
//

	public static MongoDatabase getIRADDB() {
		// create a connectstring
		ConnectionString connectionString = new ConnectionString("mongodb://localhost:27017");
		
		// configure the CodecRegistry to include a codec to handle the translation to and from BSON for our POJOs
		CodecRegistry pojoCodecRegistry = fromProviders(PojoCodecProvider.builder().automatic(true).build());

		//add the default codec registry, which contains all the default codecs
        CodecRegistry codecRegistry = fromRegistries(MongoClientSettings.getDefaultCodecRegistry(), pojoCodecRegistry);

        MongoClientSettings clientSettings = MongoClientSettings.builder()
                .applyConnectionString(connectionString)
                .codecRegistry(codecRegistry)
                .build();
        
	    // Creating a Mongo client 
	    MongoClient mongoClient = MongoClients.create(clientSettings);
	    MongoDatabase iraddb = mongoClient.getDatabase("iraddb");
	    
	    return iraddb;
	}
	
	public static ServiceAreaModelData getIRADData() {
	    MongoDatabase iraddb = getIRADDB();
        Logger mongoLogger = Logger.getLogger("org.mongodb.driver");

        mongoLogger.setLevel(Level.SEVERE);

        samData.setAirportWithCargoCollection(iraddb.getCollection("airportWithCargo", AirportWithCargo.class));
        samData.setServiceAreaCollection(iraddb.getCollection("serviceArea", ServiceArea.class));
        samData.setAirportCollection(iraddb.getCollection("airport", Airport.class));
        samData.setAirportServiceAreaCollection(iraddb.getCollection("airportServiceArea", AirportServiceArea.class));
        samData.setTravelCostWithCollection(iraddb.getCollection("travelCost", TravelCost.class));
        samData.setTravelModeCollection(iraddb.getCollection("travelMode", TravelMode.class));
        samData.setAirportTimeCollection(iraddb.getCollection("airportTime", AirportTime.class));
        samData.setEquipmentModelCollection(iraddb.getCollection("equipmentModel", EquipmentModel.class));
        samData.setAirportEquipmentCollection(iraddb.getCollection("airportEquipment", AirportEquipment.class));
        samData.setEquipmentCollection(iraddb.getCollection("equipment", Equipment.class));
        samData.setEquipmentPMCollection(iraddb.getCollection("equipmentPM", EquipmentPM.class));
        samData.setEquipmentCMCollection(iraddb.getCollection("equipmentCM", EquipmentCM.class));
        samData.setCmRateCollection(iraddb.getCollection("cmRate", CMRate.class));
        samData.setLaborRuleCollection(iraddb.getCollection("laborRule", LaborRule.class));
        samData.setTravelRuleCollection(iraddb.getCollection("travelRule", TravelRule.class));
        
	    for (Airport airport : samData.getAirportCollection().find()) {
            System.out.println(airport.getId()+" "+ airport.getCode()+" "+airport.getCity());
        }
        
        System.out.println("airportWithCargoCollection size:"+samData.getAirportWithCargoCollection().countDocuments());
        System.out.println("serviceAreaCollection size:"+samData.getServiceAreaCollection().countDocuments());
        System.out.println("airportCollection size:"+samData.getAirportCollection().countDocuments());
        System.out.println("airportServiceAreaCollection size:"+samData.getAirportServiceAreaCollection().countDocuments());
        System.out.println("travelCostWithCollection size:"+samData.getTravelCostWithCollection().countDocuments());
        System.out.println("travelModeCollection size:"+samData.getTravelModeCollection().countDocuments());
        System.out.println("airportTimeCollection size:"+samData.getAirportTimeCollection().countDocuments());
        System.out.println("equipmentModelCollection size:"+samData.getEquipmentModelCollection().countDocuments());
        System.out.println("airportEquipmentCollection size:"+samData.getAirportEquipmentCollection().countDocuments());
        System.out.println("equipmentCollection size:"+samData.getEquipmentCollection().countDocuments());
        System.out.println("equipmentPMCollection size:"+samData.getEquipmentPMCollection().countDocuments());
        System.out.println("equipmentCMCollection size:"+samData.getEquipmentCMCollection().countDocuments());
        System.out.println("cmRateCollection size:"+samData.getCmRateCollection().countDocuments());
        System.out.println("laborRuleCollection size:"+samData.getLaborRuleCollection().countDocuments());
        System.out.println("travelRuleCollection size:"+samData.getTravelRuleCollection().countDocuments());

        System.out.println("Processing completed");	
        return samData;
	}

	public static ServiceAreaModelData getSamData() {
		return samData;
	}

	public static void setSamData(ServiceAreaModelData samData) {
		IRADServiceDB.samData = samData;
	}

}
