package com.leidos.tamp.type;

public class ModelEnum {
    public enum FILTERTYPE_ENUM {

        FILTER_ModelNum,
    }

    public enum MODELSTEPTYPE_ENUM {

        MDLSTEP_initialize(1),

        MDLSTEP_randomize(2),

        MDLSTEP_loadAirports(32),

        MDLSTEP_loadServiceAreas(40),

        MDLSTEP_loadAirportServiceAreas(41),

        MDLSTEP_assignAirportServiceAreas(42),

        MDLSTEP_loadEquipmentModels(48),

        MDLSTEP_loadEquipment(49),

        MDLSTEP_loadPMRequirements(50),

        MDLSTEP_loadCMRequirements(51),

        MDLSTEP_loadDepotRequirements(52),

        MDLSTEP_loadPMStatus(56),

        MDLSTEP_createPMStatus (57);

        int val;
        MODELSTEPTYPE_ENUM(int val) {
            this.val = val;
        }
    }

    //  Enumerations
    public enum MODELDATATYPE_ENUM {

        MODDATA_Airports(1),

        MODDATA_EquipmentTypes(2),

        MODDATA_Equipment(4);

        int val;
        MODELDATATYPE_ENUM(int val) {
            this.val = val;
        }
    }
    public enum LOCATIONTYPE_ENUM {

        LOCTYPE_Airport(1),

        LOCTYPE_ServiceCenter(2);

        private final int val;

        LOCATIONTYPE_ENUM(int val) {
            this.val = val;
        }
    }
    public enum ACTIVITYTYPE_ENUM {

        //  Activity category
        ACTTYPE_PM(16777216 * 1),

        ACTTYPE_CM(16777216 * 2),

        ACTTYPE_Depot(16777216 * 3),

        //  Repair activities
        ACTTYPE_Diagnosis(1),

        ACTTYPE_PartsRequest(2),

        ACTTYPE_PartsFulfillment(3),

        ACTTYPE_PartsLocalLogistics(4),

        ACTTYPE_Repair(5),

        ACTTYPE_Test(6),

        ACTTYPE_RequestTechSupport(7),

        ACTTYPE_ProvideTechSupport(8),

        ACTTYPE_Signoff(15),

        //  Travel activities
        ACTTYPE_DriveOwnCar(48),

        ACTTYPE_DriveCompanyCar(49),

        ACTTYPE_Fly(50),

        ACTTYPE_Taxi(51),

        ACTTYPE_EnterAirport(52);

        private final int val;

        ACTIVITYTYPE_ENUM(int val) {
            this.val = val;
        }
    }
    public enum PMPERIODICITY_ENUM {

        NoPM(0),

        Daily(1),

        Weekly(2),

        Biweekly(3),

        Monthly(4),

        Quarterly(5),

        SemiAnnually(6),

        Annually(7);

        private final int val;

        PMPERIODICITY_ENUM(int val) {
            this.val = val;
        }

        public int value() { return val;}
    }


    public enum MTRIP_STATUS_ENUM {

        TRIPSTAT_Scheduled(1),

        TRIPSTAT_Active(2),

        TRIPSTAT_Completed(3);

        private final int val;

        MTRIP_STATUS_ENUM(int val) {
            this.val = val;
        }
    }

    public enum MTRIP_ITEMTYPE_ENUM {

        TRIPITEM_Travel(1),

        TRIPITEM_PM(2),

        TRIPITEM_CM(3),

        TRIPITEM_Other(4);

        private final int val;

        MTRIP_ITEMTYPE_ENUM(int val) {
            this.val = val;
        }
    }


    public enum EVOLUTIONAPPROACH_ENUM {

        EVOLAPP_RetainTop,

        EVOLAPP_RetainRandom,

        EVOLAPP_Mate,

        EVOLAPP_InsertDelete,

        EVOLAPP_Random,
    }



}
