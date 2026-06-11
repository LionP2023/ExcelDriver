package ExcelDriver;
/**
 *  Copyright © 2021-2022 The CETC PHM Authors
 *
 */
import ExcelDriver.Core.ReportType;
import ExcelDriver.Core._Row;
import ExcelDriver.Core._Style;
import ExcelDriver.Core._Workbook;
import ExcelDriver.Core._Worksheet;

import java.text.SimpleDateFormat;

import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;


public class ExcelCreator {
    
    private static final String FORMAT_XLSX = "xlsx";
    private static final String FORMAT_XLS = "xls";
        
    private _Workbook wb;
    private _Worksheet sheet;

    // Excel path file
    private String excelFileName;
    // Excel name file without path
    private String excelSheetName;
    // Excel number of columns
    private int columnsNumber;
    // Excel row position
    private int rowPosition = 0;
    // maximum table column width
    private int maxColumnWidth;
    // file format: old or new
    private boolean isFileType;
    
    public ExcelCreator() {
        super();
        this.wb = new _Workbook();
        this.wb.create();
        this.sheet = new _Worksheet(wb.getWorkbook());
        this.maxColumnWidth = 0;
        this.isFileType = true;
    }
    
    public ExcelCreator(String excelFilePath, int columnsNumber) {
        this.wb = new _Workbook();
        this.wb.create();
        this.wb.setPath(excelFilePath);
        this.sheet = new _Worksheet(wb.getWorkbook());
        this.columnsNumber = columnsNumber;
        this.maxColumnWidth = 0;
        this.isFileType = true;
    }
        
    // set Excel file path (no end slash)
    public void setPath(String path){
        this.wb.setPath(path);
    }
    
    // set Excel file name
    public void setFileName(String fileName){
        if(fileName == null){
            this.excelFileName = this.wb.getName();
        } else {
            this.excelFileName = fileName;
            this.wb.setName(fileName);
        }
    }
    
    // generic Excel file name
    public void setFileName(ReportType type){
        String fn = getNameFileReport(type);
        this.excelFileName = fn + "." + ((this.isFileType)? FORMAT_XLSX:FORMAT_XLS);
        this.wb.setName(excelFileName);
    }
    
    // set Excel sheet name
    public void setSheetName(String sheetName){
        if(sheetName == null){
            this.excelSheetName = this.sheet.getName();
        } else {
            this.excelSheetName = sheetName;
            this.sheet.setName(sheetName);
        }
    }
    
    // set number of columns
    public void setColumnsNumber(int columnsNumber){
        this.columnsNumber = columnsNumber;
    }
    
    // set max column width
    public void setMaxColumnWidth(int numSymbols){
        this.maxColumnWidth = numSymbols;
    }
    
    // set Excel file format
    public void setFileType(boolean isNew){
        this.isFileType = isNew;
    }
    
    // get current row position
    public int getPosition(){
        return this.rowPosition;
    }
    
    // set current row position
    public void setPosition(int newPosition){
        this.rowPosition = newPosition;
    }
    
    // adding HEADER TABLE data
    public void addHeaderData(List<String> data){
        _Row r = new _Row(this.sheet.getSheet(), this.rowPosition);
        _Style styleHeader = new _Style(this.wb.getWorkbook());
        // set style for cells
        CellStyle style = styleHeader.getTableHeaderDefault();
        for (int i = 0; i < data.size(); i++) {
            XSSFCell cell = r.createCell(i);
            cell.setCellValue(data.get(i));
            cell.setCellStyle(style);
        }
        // save last row position
        this.rowPosition = r.getNumber() + 1;
    }
    
    public void addHeaderData(String[] data){
        _Row r = new _Row(this.sheet.getSheet(), this.rowPosition);
        _Style styleHeader = new _Style(this.wb.getWorkbook());
        // set style for cells
        CellStyle style = styleHeader.getTableHeaderDefault();
        for (int i = 0; i < data.length; i++) {
            XSSFCell cell = r.createCell(i);
            cell.setCellValue(data[i]);
            cell.setCellStyle(style);
        }
        // save last row position
        this.rowPosition = r.getNumber() + 1;
    }
    
    // add TABLE DETAILS data
    public void addTableData(List<List<String>> data){
        _Style styleTable = new _Style(this.wb.getWorkbook());
        CellStyle style = styleTable.getTableDetailsDefault();
        //CellStyle styleColor = styleTable.getTableDetailsDefaultColor();
        for (int i = 0; i < data.size(); i++) {
            XSSFRow row = this.sheet.getSheet().createRow(this.rowPosition + i);
            for (int j = 0; j < data.get(i).size(); j++) {
                if(data.get(i).get(j) == null) continue;
                XSSFCell cell = row.createCell(j);
                
                if(this.maxColumnWidth == 0){
                    this.sheet.getSheet().autoSizeColumn(j, true);
                } else {
                    if(data.get(i).get(j).length() > this.maxColumnWidth){
                        this.sheet.setColWidth(j, 256*this.maxColumnWidth);
                    } else {
                        this.sheet.getSheet().autoSizeColumn(j, true);
                    }
                }
                
                cell.setCellValue(data.get(i).get(j));
                
               // if(i%2 == 0) {
               //     cell.setCellStyle(styleColor);
                //} else {
                    cell.setCellStyle(style);
                //}
            }
        }
    }
    
    public void addTableData(String[][] data){
        _Style styleTable = new _Style(this.wb.getWorkbook());
        CellStyle style = styleTable.getTableDetailsDefault();
        //CellStyle styleColor = styleTable.getTableDetailsDefaultColor();
        for (int i = 0; i < data.length; i++) {
            XSSFRow row = this.sheet.getSheet().createRow(this.rowPosition + i);
            for (int j = 0; j < data[i].length; j++) {
                if(data[i][j] == null) continue;
                XSSFCell cell = row.createCell(j);

                if(this.maxColumnWidth == 0){
                    this.sheet.getSheet().autoSizeColumn(j, true);
                } else {
                    if(data[i][j].length() > this.maxColumnWidth){
                        this.sheet.setColWidth(j, 256*this.maxColumnWidth);
                    } else {
                        this.sheet.getSheet().autoSizeColumn(j, true);
                    }
                }

                cell.setCellValue(data[i][j]);
                
                //if((i%2) != 0) {
                //    cell.setCellStyle(styleColor);
                //} else {
                    cell.setCellStyle(style);
                //}
            }
        }
    }
    
    // generate Excel file name
    public String getNameFileReport(ReportType t){
        //String name = t.getTypeTitle() + "_" + System.currentTimeMillis();//(new Date(System.currentTimeMillis())).toString();
        String name = t.getTypeTitle() + "_" + new SimpleDateFormat("yyyy-MM-dd_HH_mm_ss").format(new Date());
        return name;
    }
    
    // save to Excel file
    public void saveToExcel(){
        if(this.excelFileName != null){
            this.wb.save();
        }
    }
    
    // get max column width
//    private int getMaxColumnWidth(List<List<String>> data){
//        int width = 0;
//        for (int i = 0; i < data.size(); i++) {
//            for (int j = 0; j < data.get(i).size(); j++) {
//                if(data.get(i).get(j).length() > width){
//                    width = data.get(i).get(j).length();
//                }
//            }
//        }
//        return width;
//    }
//    
//    private int getMaxColumnWidth(String[][] data){
//        int width = 0;
//        for (int i = 0; i < data.length; i++) {
//            for (int j = 0; j < data[i].length; j++) {
//                if(data[i][j].length() > width){
//                    width = data[i][j].length();
//                }
//            }
//        }
//        return width;
//    }
    
}
