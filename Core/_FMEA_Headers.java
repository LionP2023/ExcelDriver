package ExcelDriver.Core;
/**
 * Copyright © 2021-2022 The CETC PHM Authors
 *
 */
import ExcelDriver.Core.ReportType;

import java.util.Arrays;
import java.util.List;

public class _FMEA_Headers {
    
    private static String F_UNIT_NAME = "Unit Name";
    private static String F_TYPE_FMEA = "Type FMEA";
    private static String F_ID = "Identification Number";
    private static String F_ITEM = "Item/Functional Idention";
    private static String F_FAILURE_MODE = "Failure Mode";
    private static String F_FAILURE_CAUSE = "Failure Cause";
    private static String F_LEVEL_S = "Severity level (S)";
    private static String F_LEVEL_O = "Probability level (O)";
    private static String F_LEVEL_D = "Detection difficulty level (D)";
    private static String F_LOCAL_EFFECT = "Local Effects";
    private static String F_NEXT_EFFECT = "Next Higher Level";
    private static String F_END_EFFECT = "End Effects";
    private static String F_RPN = "RPN";
    private static String F_FAILURE_MODE_CRITICAL = "Failure Mode Critical";
    private static String F_SEVERITY_CLASS = "Severity Class";
    private static String F_FAILURE_PROBABILITY = "Failure Probability";
    private static String F_CORRECT_ACTION = "Corrective Action";
    private static String F_NUMBER_FAILURE_MODES = "Numbers of Failure Modes";
    private static String F_PERCENT = "Percentage";
    private static String F_CRITICAL = "Criticality";
    private static String F_PROBABILITY = "Probability";
    private static String F_SEVERITY = "Severity";
    private static String F_QUEUE = "Queue";
    
    
    
    ReportType reportType;
    
    public _FMEA_Headers(ReportType t) {
        this.reportType = t;
    }
        
    private List<String> getHeaderDataList(){
        List<String> headerData = null;
        switch(this.reportType){
            case FMEA_QUALITATIVE:
            case FMEA_QUANITATIVE:
                headerData = Arrays.asList(F_UNIT_NAME, F_TYPE_FMEA, F_ID, F_ITEM, F_FAILURE_MODE, F_FAILURE_CAUSE,
                                           F_LEVEL_S, F_LEVEL_O, F_LEVEL_D, F_RPN);
                break;
            case FMEA_LIST_FAILURE_MODE:
                headerData = Arrays.asList(F_UNIT_NAME, F_TYPE_FMEA, F_ID, F_ITEM, F_FAILURE_MODE, F_FAILURE_CAUSE,
                                           F_LOCAL_EFFECT, F_NEXT_EFFECT, F_END_EFFECT, F_FAILURE_MODE_CRITICAL);
                break;
            case FMEA_KEY_FAILURE_MODE_LIST:
                headerData = Arrays.asList(F_UNIT_NAME, F_TYPE_FMEA, F_ID, F_ITEM, F_LOCAL_EFFECT, F_NEXT_EFFECT, F_END_EFFECT,
                                           F_FAILURE_MODE, F_FAILURE_CAUSE, F_SEVERITY_CLASS, F_LEVEL_S, F_LEVEL_O,
                                           F_FAILURE_PROBABILITY, F_CORRECT_ACTION);
                break;
            case FMEA_SINGLE_POINT_FAILURE_LIST:
                headerData = Arrays.asList(F_UNIT_NAME, F_TYPE_FMEA, F_ID, F_ITEM, F_SEVERITY_CLASS, F_FAILURE_MODE, F_FAILURE_CAUSE,
                                           F_LOCAL_EFFECT, F_NEXT_EFFECT, F_END_EFFECT);
                break;
            case FMEA_COUNT_CHILD_NODES:
                headerData = Arrays.asList(F_UNIT_NAME, F_TYPE_FMEA, F_NUMBER_FAILURE_MODES, F_PERCENT);
                break;
            case FMEA_STATISTIC_CHILD_NODES_LIST:
                headerData = Arrays.asList(F_UNIT_NAME, F_CRITICAL);
                break;
            case FMEA_HAZARD_MATRIX:
                headerData = Arrays.asList(F_PROBABILITY, F_SEVERITY, F_QUEUE);
                break;
            case FMEA_HAZARD_MATRIX_DETAILS:
                headerData = Arrays.asList(F_UNIT_NAME, F_ID, F_ITEM, F_FAILURE_MODE, F_FAILURE_CAUSE, F_SEVERITY_CLASS, F_FAILURE_PROBABILITY);
                break;
        }
        
        return headerData;
    }
    
    private String[] getHeaderDataStringArray(){
        String[] headerData = null;
        switch(this.reportType){
            case FMEA_QUALITATIVE:
            case FMEA_QUANITATIVE:
                headerData = new String[]{F_UNIT_NAME, F_TYPE_FMEA, F_ID, F_ITEM, F_FAILURE_MODE, F_FAILURE_CAUSE,
                                          F_LEVEL_S, F_LEVEL_O, F_LEVEL_D, F_RPN};
                break;
            case FMEA_LIST_FAILURE_MODE:
                headerData = new String[]{F_UNIT_NAME, F_TYPE_FMEA, F_ID, F_ITEM, F_FAILURE_MODE, F_FAILURE_CAUSE,
                                          F_LOCAL_EFFECT, F_NEXT_EFFECT, F_END_EFFECT, F_FAILURE_MODE_CRITICAL};
                break;
            case FMEA_KEY_FAILURE_MODE_LIST:
                headerData = new String[]{F_UNIT_NAME, F_TYPE_FMEA, F_ID, F_ITEM, F_LOCAL_EFFECT, F_NEXT_EFFECT, F_END_EFFECT,
                                          F_FAILURE_MODE, F_FAILURE_CAUSE, F_SEVERITY_CLASS, F_LEVEL_S, F_LEVEL_O,
                                          F_FAILURE_PROBABILITY, F_CORRECT_ACTION};
                break;
            case FMEA_SINGLE_POINT_FAILURE_LIST:
                headerData = new String[]{F_UNIT_NAME, F_TYPE_FMEA, F_ID, F_ITEM, F_SEVERITY_CLASS, F_FAILURE_MODE, F_FAILURE_CAUSE,
                                           F_LOCAL_EFFECT, F_NEXT_EFFECT, F_END_EFFECT};
                break;
            case FMEA_COUNT_CHILD_NODES:
                headerData = new String[]{F_UNIT_NAME, F_TYPE_FMEA, F_NUMBER_FAILURE_MODES, F_PERCENT};
                break;
            case FMEA_STATISTIC_CHILD_NODES_LIST:
                headerData = new String[]{F_UNIT_NAME, F_CRITICAL};
                break;
            case FMEA_HAZARD_MATRIX:
                headerData = new String[]{F_PROBABILITY, F_SEVERITY, F_QUEUE};
                break;
            case FMEA_HAZARD_MATRIX_DETAILS:
                headerData = new String[]{F_UNIT_NAME, F_ID, F_ITEM, F_FAILURE_MODE, F_FAILURE_CAUSE, F_SEVERITY_CLASS, F_FAILURE_PROBABILITY};
                break;

        }
        
        return headerData;
    }
        
    // get header data
    public List<String> getHeaderList(){
        return getHeaderDataList();
    }
    
    public String[] getHeaderString(){
        return getHeaderDataStringArray();
    }
        
        
}
