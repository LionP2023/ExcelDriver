package ExcelDriver;
/**
 * Copyright © 2021-2022 The CETC PHM Authors
 *
 */

import ExcelDriver.Core.ReportType;
import ExcelDriver.Core._FMEA_Headers;

import java.util.List;


public class BuildReportEx_FMEA {
    
    // path to Excel file on desk
    private String excelFilePath;
    // number of columns in Excel table
    private int columnsNumber;
    // maximum table column width
    private int limitColumnWidth;
    // file excention type
    private boolean isFileType;


    // constructor: set path for Excel file and number columns
    public BuildReportEx_FMEA(String excelFilePath) {
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
    // tableData - data for table rows
    public void save_FMEA_QUALITATIVE(List<List<String>> tableData){
        _FMEA_Headers h = new _FMEA_Headers(ReportType.FMEA_QUALITATIVE);
        List<String> headerData = h.getHeaderList();
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
    
    public void save_FMEA_QUALITATIVE(String[][] tableData){
        _FMEA_Headers h = new _FMEA_Headers(ReportType.FMEA_QUALITATIVE);
        String[] headerData = h.getHeaderString();
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
    // tableData - data for table rows    
    public void save_FMEA_QUANITATIVE(List<List<String>> tableData){
        _FMEA_Headers h = new _FMEA_Headers(ReportType.FMEA_QUANITATIVE);
        List<String> headerData = h.getHeaderList();
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
    
    public void save_FMEA_QUANITATIVE(String[][] tableData){
        _FMEA_Headers h = new _FMEA_Headers(ReportType.FMEA_QUANITATIVE);
        String[] headerData = h.getHeaderString();
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
    // tableData - data for table rows
    public void save_FMEA_LIST_FAILURE_MODE(List<List<String>> tableData){
        _FMEA_Headers h = new _FMEA_Headers(ReportType.FMEA_LIST_FAILURE_MODE);
        List<String> headerData = h.getHeaderList();
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
    
    public void save_FMEA_LIST_FAILURE_MODE(String[][] tableData){
        _FMEA_Headers h = new _FMEA_Headers(ReportType.FMEA_LIST_FAILURE_MODE);
        String[] headerData = h.getHeaderString();
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
    // tableData - data for table rows
    public void save_FMEA_KEY_FAILURE_MODE_LIST(List<List<String>> tableData){
        _FMEA_Headers h = new _FMEA_Headers(ReportType.FMEA_KEY_FAILURE_MODE_LIST);
        List<String> headerData = h.getHeaderList();
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
    
    public void save_FMEA_KEY_FAILURE_MODE_LIST(String[][] tableData){
        _FMEA_Headers h = new _FMEA_Headers(ReportType.FMEA_KEY_FAILURE_MODE_LIST);
        String[] headerData = h.getHeaderString();
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
    // tableData - data for table rows  
    public void save_FMEA_SINGLE_POINT_FAILURE_LIST(List<List<String>> tableData){
        _FMEA_Headers h = new _FMEA_Headers(ReportType.FMEA_SINGLE_POINT_FAILURE_LIST);
        List<String> headerData = h.getHeaderList();
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
    
    public void save_FMEA_SINGLE_POINT_FAILURE_LIST(String[][] tableData){
        _FMEA_Headers h = new _FMEA_Headers(ReportType.FMEA_SINGLE_POINT_FAILURE_LIST);
        String[] headerData = h.getHeaderString();
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
    // tableData - data for table rows 
    public void save_FMEA_COUNT_CHILD_NODES(List<List<String>> tableData){
        _FMEA_Headers h = new _FMEA_Headers(ReportType.FMEA_COUNT_CHILD_NODES);
        List<String> headerData = h.getHeaderList();
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
    
    public void save_FMEA_COUNT_CHILD_NODES(String[][] tableData){
        _FMEA_Headers h = new _FMEA_Headers(ReportType.FMEA_COUNT_CHILD_NODES);
        String[] headerData = h.getHeaderString();
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
    // tableData - data for table rows 
    public void save_FMEA_STATISTIC_CHILD_NODES_LIST(List<List<String>> tableData){
        _FMEA_Headers h = new _FMEA_Headers(ReportType.FMEA_STATISTIC_CHILD_NODES_LIST);
        List<String> headerData = h.getHeaderList();
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
    
    public void save_FMEA_STATISTIC_CHILD_NODES_LIST(String[][] tableData){
        _FMEA_Headers h = new _FMEA_Headers(ReportType.FMEA_STATISTIC_CHILD_NODES_LIST);
        String[] headerData = h.getHeaderString();
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
    // tableData - data for table rows  
    public void save_FMEA_HAZARD_MATRIX(List<List<String>> tableData){
        _FMEA_Headers h = new _FMEA_Headers(ReportType.FMEA_HAZARD_MATRIX);
        List<String> headerData = h.getHeaderList();
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
    
    public void save_FMEA_HAZARD_MATRIX(String[][] tableData){
        _FMEA_Headers h = new _FMEA_Headers(ReportType.FMEA_HAZARD_MATRIX);
        String[] headerData = h.getHeaderString();
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
    // tableData - data for table rows  
    public void save_FMEA_HAZARD_MATRIX_DETAILS(List<List<String>> tableData){
        _FMEA_Headers h = new _FMEA_Headers(ReportType.FMEA_HAZARD_MATRIX_DETAILS);
        List<String> headerData = h.getHeaderList();
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
    
    public void save_FMEA_HAZARD_MATRIX_DETAILS(String[][] tableData){
        _FMEA_Headers h = new _FMEA_Headers(ReportType.FMEA_HAZARD_MATRIX_DETAILS);
        String[] headerData = h.getHeaderString();
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