package ExcelDriver.Core;
/**
 * Copyright © 2021-2022 The CETC PHM Authors
 *
 */
public enum ReportType {
    
    // FMEA reports
    FMEA_QUALITATIVE(100, "FMEA Qualitative"),
    FMEA_QUANITATIVE(101, "FMEA Quantitative"),
    FMEA_LIST_FAILURE_MODE(102, "List of Type_Failure Mode"),
    FMEA_KEY_FAILURE_MODE_LIST(103, "Key Failure Mode List"),
    FMEA_SINGLE_POINT_FAILURE_LIST(104, "Single Point of Failure List"),
    FMEA_CONTROL_PLAN(105, "FMEA Control Plan Report"),
    FMEA_DVP_R(106, "FMEA Design Verification Plan and Report"),
    FMEA_HAZARD_MATRIX(107, "Hazard Matrix"),
    FMEA_COUNT_CHILD_NODES(108, "Count Child Nodes"),
    FMEA_STATISTIC_CHILD_NODES_LIST(109, "Statistic of Child Nodes by Risk"),
    FMEA_HAZARD_MATRIX_DETAILS(110,"Hazard Matrix Details");
    
    private final int typeIndex;
    private final String typeTitle;
    
    ReportType(int index, String title){
        this.typeIndex = index;  
        this.typeTitle = title;
    }
    
    public int getTypeIndex(){
        return this.typeIndex;
    }
    
    public String getTypeTitle(){
        return this.typeTitle;
    }
}
