package ExcelDriver;
/**
 * Copyright © 2021-2022 The CETC PHM Authors
 *
 */
import ExcelDriver.Core.ReportType;

import java.util.List;


public class BuildReport_FMEA {
    
    // path to Excel file on desk
    private String excelFilePath;
    // number of columns in Excel table
    private int columnsNumber;
    // maximum table column width
    private int limitColumnWidth;
    // file excention type
    private boolean isFileType;


    // constructor: set path for Excel file and number columns
    public BuildReport_FMEA(String excelFilePath) {
        super();
        this.excelFilePath = excelFilePath;
        this.limitColumnWidth = 0;
        this.isFileType = true;
    }
    
    // set columns Excel file
    public void setColumns(int colNum){
        this.columnsNumber = colNum;
    }
    
    // set max column width by symbols
    public void setMaxColumnWidth(int limit){
        this.limitColumnWidth = limit;
    }
    
    // set file type
    public void setFileType(boolean isNew){
        this.isFileType = isNew;
    }

    // save FMEA qualitative report
    // headerData - data for table columns
    // tableData - data for table rows
    public void save_FMEA_QUALITATIVE(List<String> headerData, List<List<String>> tableData){
        if(!(this.columnsNumber > 0)){
            this.columnsNumber = headerData.size();
        }
        ExcelCreator creator = new ExcelCreator(this.excelFilePath, this.columnsNumber);
        creator.setFileType(this.isFileType);
        creator.setFileName(ReportType.FMEA_QUALITATIVE);
        if(this.limitColumnWidth > 0){creator.setMaxColumnWidth(this.limitColumnWidth);}
        creator.addHeaderData(headerData);
        creator.addTableData(tableData);
        creator.saveToExcel();
    }
    
    public void save_FMEA_QUALITATIVE(String[] headerData, String[][] tableData){
        if(!(this.columnsNumber > 0)){
            this.columnsNumber = headerData.length;
        }
        ExcelCreator creator = new ExcelCreator(this.excelFilePath, this.columnsNumber);
        creator.setFileType(this.isFileType);
        creator.setFileName(ReportType.FMEA_QUALITATIVE);
        if(this.limitColumnWidth > 0){creator.setMaxColumnWidth(this.limitColumnWidth);}
        creator.addHeaderData(headerData);
        creator.addTableData(tableData);
        creator.saveToExcel();
    }
    
    // save FMEA quanitative report
    // headerData - data for table columns
    // tableData - data for table rows    
    public void save_FMEA_QUANITATIVE(List<String> headerData, List<List<String>> tableData){
        if(!(this.columnsNumber > 0)){
            this.columnsNumber = headerData.size();
        }
        ExcelCreator creator = new ExcelCreator(this.excelFilePath, this.columnsNumber);
        creator.setFileType(this.isFileType);
        creator.setFileName(ReportType.FMEA_QUANITATIVE);
        if(this.limitColumnWidth > 0){creator.setMaxColumnWidth(this.limitColumnWidth);}
         creator.addHeaderData(headerData);
        creator.addTableData(tableData);
        creator.saveToExcel();
    }
    
    public void save_FMEA_QUANITATIVE(String[] headerData, String[][] tableData){
        if(!(this.columnsNumber > 0)){
            this.columnsNumber = headerData.length;
        }
        ExcelCreator creator = new ExcelCreator(this.excelFilePath, this.columnsNumber);
        creator.setFileType(this.isFileType);
        creator.setFileName(ReportType.FMEA_QUANITATIVE);
        if(this.limitColumnWidth > 0){creator.setMaxColumnWidth(this.limitColumnWidth);}
        creator.addHeaderData(headerData);
        creator.addTableData(tableData);
        creator.saveToExcel();
    }
    
    // save List of Type_Failure Mode report
    // headerData - data for table columns
    // tableData - data for table rows
    public void save_FMEA_LIST_FAILURE_MODE(List<String> headerData, List<List<String>> tableData){
        if(!(this.columnsNumber > 0)){
            this.columnsNumber = headerData.size();
        }
        ExcelCreator creator = new ExcelCreator(this.excelFilePath, this.columnsNumber);
        creator.setFileType(this.isFileType);
        creator.setFileName(ReportType.FMEA_LIST_FAILURE_MODE);
        if(this.limitColumnWidth > 0){creator.setMaxColumnWidth(this.limitColumnWidth);}
        creator.addHeaderData(headerData);
        creator.addTableData(tableData);
        creator.saveToExcel();
    }
    
    public void save_FMEA_LIST_FAILURE_MODE(String[] headerData, String[][] tableData){
        if(!(this.columnsNumber > 0)){
            this.columnsNumber = headerData.length;
        }
        ExcelCreator creator = new ExcelCreator(this.excelFilePath, this.columnsNumber);
        creator.setFileType(this.isFileType);
        creator.setFileName(ReportType.FMEA_LIST_FAILURE_MODE);
        if(this.limitColumnWidth > 0){creator.setMaxColumnWidth(this.limitColumnWidth);}
        creator.addHeaderData(headerData);
        creator.addTableData(tableData);
        creator.saveToExcel();
    }
    
    // save Key Failure Mode List report
    // headerData - data for table columns
    // tableData - data for table rows
    public void save_FMEA_KEY_FAILURE_MODE_LIST(List<String> headerData, List<List<String>> tableData){
        if(!(this.columnsNumber > 0)){
            this.columnsNumber = headerData.size();
        }
        ExcelCreator creator = new ExcelCreator(this.excelFilePath, this.columnsNumber);
        creator.setFileType(this.isFileType);
        creator.setFileName(ReportType.FMEA_KEY_FAILURE_MODE_LIST);
        if(this.limitColumnWidth > 0){creator.setMaxColumnWidth(this.limitColumnWidth);}
        creator.addHeaderData(headerData);
        creator.addTableData(tableData);
        creator.saveToExcel();       
    }
    
    public void save_FMEA_KEY_FAILURE_MODE_LIST(String[] headerData, String[][] tableData){
        if(!(this.columnsNumber > 0)){
            this.columnsNumber = headerData.length;
        }
        ExcelCreator creator = new ExcelCreator(this.excelFilePath, this.columnsNumber);
        creator.setFileType(this.isFileType);
        creator.setFileName(ReportType.FMEA_KEY_FAILURE_MODE_LIST);
        if(this.limitColumnWidth > 0){creator.setMaxColumnWidth(this.limitColumnWidth);}
        creator.addHeaderData(headerData);
        creator.addTableData(tableData);
        creator.saveToExcel();
    }
    
    // save Single Point of Failure List report
    // headerData - data for table columns
    // tableData - data for table rows   
    public void save_FMEA_SINGLE_POINT_FAILURE_LIST(List<String> headerData, List<List<String>> tableData){
        if(!(this.columnsNumber > 0)){
            this.columnsNumber = headerData.size();
        }
        ExcelCreator creator = new ExcelCreator(this.excelFilePath, this.columnsNumber);
        creator.setFileType(this.isFileType);
        creator.setFileName(ReportType.FMEA_SINGLE_POINT_FAILURE_LIST);
        if(this.limitColumnWidth > 0){creator.setMaxColumnWidth(this.limitColumnWidth);}
        creator.addHeaderData(headerData);
        creator.addTableData(tableData);
        creator.saveToExcel();       
    }
    
    public void save_FMEA_SINGLE_POINT_FAILURE_LIST(String[] headerData, String[][] tableData){
        if(!(this.columnsNumber > 0)){
            this.columnsNumber = headerData.length;
        }
        ExcelCreator creator = new ExcelCreator(this.excelFilePath, this.columnsNumber);
        creator.setFileType(this.isFileType);
        creator.setFileName(ReportType.FMEA_SINGLE_POINT_FAILURE_LIST);
        if(this.limitColumnWidth > 0){creator.setMaxColumnWidth(this.limitColumnWidth);}
        creator.addHeaderData(headerData);
        creator.addTableData(tableData);
        creator.saveToExcel();
    }
    
    // save Count Child Nodes report
    // headerData - data for table columns
    // tableData - data for table rows  
    public void save_FMEA_COUNT_CHILD_NODES(List<String> headerData, List<List<String>> tableData){
        if(!(this.columnsNumber > 0)){
            this.columnsNumber = headerData.size();
        }
        ExcelCreator creator = new ExcelCreator(this.excelFilePath, this.columnsNumber);
        creator.setFileType(this.isFileType);
        creator.setFileName(ReportType.FMEA_COUNT_CHILD_NODES);
        if(this.limitColumnWidth > 0){creator.setMaxColumnWidth(this.limitColumnWidth);}
        creator.addHeaderData(headerData);
        creator.addTableData(tableData);
        creator.saveToExcel();       
    }
    
    public void save_FMEA_COUNT_CHILD_NODES(String[] headerData, String[][] tableData){
        if(!(this.columnsNumber > 0)){
            this.columnsNumber = headerData.length;
        }
        ExcelCreator creator = new ExcelCreator(this.excelFilePath, this.columnsNumber);
        creator.setFileType(this.isFileType);
        creator.setFileName(ReportType.FMEA_COUNT_CHILD_NODES);
        if(this.limitColumnWidth > 0){creator.setMaxColumnWidth(this.limitColumnWidth);}
        creator.addHeaderData(headerData);
        creator.addTableData(tableData);
        creator.saveToExcel();
    }

    // save Statistic of Child Nodes by Risk report
    // headerData - data for table columns
    // tableData - data for table rows  
    public void save_FMEA_STATISTIC_CHILD_NODES_LIST(List<String> headerData, List<List<String>> tableData){
        if(!(this.columnsNumber > 0)){
            this.columnsNumber = headerData.size();
        }
        ExcelCreator creator = new ExcelCreator(this.excelFilePath, this.columnsNumber);
        creator.setFileType(this.isFileType);
        creator.setFileName(ReportType.FMEA_STATISTIC_CHILD_NODES_LIST);
        if(this.limitColumnWidth > 0){creator.setMaxColumnWidth(this.limitColumnWidth);}
        creator.addHeaderData(headerData);
        creator.addTableData(tableData);
        creator.saveToExcel();       
    }
    
    public void save_FMEA_STATISTIC_CHILD_NODES_LIST(String[] headerData, String[][] tableData){
        if(!(this.columnsNumber > 0)){
            this.columnsNumber = headerData.length;
        }
        ExcelCreator creator = new ExcelCreator(this.excelFilePath, this.columnsNumber);
        creator.setFileType(this.isFileType);
        creator.setFileName(ReportType.FMEA_STATISTIC_CHILD_NODES_LIST);
        if(this.limitColumnWidth > 0){creator.setMaxColumnWidth(this.limitColumnWidth);}
        creator.addHeaderData(headerData);
        creator.addTableData(tableData);
        creator.saveToExcel();
    }
    
    // save Hazard Matrix report
    // headerData - data for table columns
    // tableData - data for table rows  
    public void save_FMEA_HAZARD_MATRIX(List<String> headerData, List<List<String>> tableData){
        if(!(this.columnsNumber > 0)){
            this.columnsNumber = headerData.size();
        }
        ExcelCreator creator = new ExcelCreator(this.excelFilePath, this.columnsNumber);
        creator.setFileType(this.isFileType);
        creator.setFileName(ReportType.FMEA_HAZARD_MATRIX);
        if(this.limitColumnWidth > 0){creator.setMaxColumnWidth(this.limitColumnWidth);}
        creator.addHeaderData(headerData);
        creator.addTableData(tableData);
        creator.saveToExcel();  
    }
    
    public void save_FMEA_HAZARD_MATRIX(String[] headerData, String[][] tableData){
        if(!(this.columnsNumber > 0)){
            this.columnsNumber = headerData.length;
        }
        ExcelCreator creator = new ExcelCreator(this.excelFilePath, this.columnsNumber);
        creator.setFileType(this.isFileType);
        creator.setFileName(ReportType.FMEA_HAZARD_MATRIX);
        if(this.limitColumnWidth > 0){creator.setMaxColumnWidth(this.limitColumnWidth);}
        creator.addHeaderData(headerData);
        creator.addTableData(tableData);
        creator.saveToExcel();
    }
    
    // save Hazard Matrix Details report
    // headerData - data for table columns
    // tableData - data for table rows  
    public void save_FMEA_HAZARD_MATRIX_DETAILS(List<String> headerData, List<List<String>> tableData){
        if(!(this.columnsNumber > 0)){
            this.columnsNumber = headerData.size();
        }
        ExcelCreator creator = new ExcelCreator(this.excelFilePath, this.columnsNumber);
        creator.setFileType(this.isFileType);
        creator.setFileName(ReportType.FMEA_HAZARD_MATRIX_DETAILS);
        if(this.limitColumnWidth > 0){creator.setMaxColumnWidth(this.limitColumnWidth);}
        creator.addHeaderData(headerData);
        creator.addTableData(tableData);
        creator.saveToExcel();  
    }
    
    public void save_FMEA_HAZARD_MATRIX_DETAILS(String[] headerData, String[][] tableData){
        if(!(this.columnsNumber > 0)){
            this.columnsNumber = headerData.length;
        }
        ExcelCreator creator = new ExcelCreator(this.excelFilePath, this.columnsNumber);
        creator.setFileType(this.isFileType);
        creator.setFileName(ReportType.FMEA_HAZARD_MATRIX_DETAILS);
        if(this.limitColumnWidth > 0){creator.setMaxColumnWidth(this.limitColumnWidth);}
        creator.addHeaderData(headerData);
        creator.addTableData(tableData);
        creator.saveToExcel();
    }
    
}
