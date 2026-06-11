package ExcelDriver.Core;
/**
 * Copyright © 2021-2022 The CETC PHM Authors
 *
 */

import java.util.ArrayList;
import java.util.Date;

import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class _Cell {
    
    private XSSFSheet sheet;
    private XSSFCell cell;
    
    public _Cell(XSSFSheet sheet, int rowNum, int colNum) {
        this.sheet = sheet;
        this.cell = this.sheet.getRow(rowNum).getCell(colNum);
    }
    
    public _Cell(XSSFSheet sheet){
        this.sheet = sheet;
    }
    
    // return sheet
    public XSSFSheet getSheet(){
        return this.sheet;
    }
    
    // return cell
    public XSSFCell getCell(){
        return this.cell;
    }
    
    public XSSFCell getCell(int rowNum, int colNum){
        this.cell = this.sheet.getRow(rowNum).getCell(colNum);
        return this.cell;
    }
    
    // get Cell style
    public CellStyle getStyle(){
        return this.cell.getCellStyle();
    }
    
    // set Cell style
    public void setStyle(CellStyle style){
        this.cell.setCellStyle(style);
    }
    
    // get Cell type
    public CellType getType(){
        return this.cell.getCellType();
    }
    
    // set Cell type
    public void setType(CellType type){
        this.cell.setCellType(type);
    }
    
    // convert cell value in string
    public String cellToString(){
        if (null != this.cell) {
                switch (this.cell.getCellType()) {
                case BLANK:
                        return "";
                case BOOLEAN:
                    Boolean b = this.cell.getBooleanCellValue();
                    return b==true? "true" : "false";
                case ERROR:
                        return "";
                case FORMULA:
                        return "";
                case NUMERIC:
                        Double d = this.cell.getNumericCellValue();
                        Long l = d.longValue();
                        if (l.doubleValue() == d) {
                                return l.toString();
                        } else {
                                return d.toString();
                        }
                default:
                        return this.cell.getStringCellValue();
                }
        }
        return "";
    }
    
    public String cellToString(XSSFCell cell){
        if (null != cell) {
                switch (cell.getCellType()) {
                case BLANK:
                        return "";
                case BOOLEAN:
                    Boolean b = cell.getBooleanCellValue();
                    return b==true? "true" : "false";
                case ERROR:
                        return "";
                case FORMULA:
                        return "";
                case NUMERIC:
                        Double d = cell.getNumericCellValue();
                        Long l = d.longValue();
                        if (l.doubleValue() == d) {
                                return l.toString();
                        } else {
                                return d.toString();
                        }
                default:
                        return cell.getStringCellValue();
                }
        }
        return "";
    }
    
    // get cell value
    public String getValue(){
        if (this.cell != null)
        {
            switch (this.cell.getCellType())
            {
                case BLANK:
                        return null;
                case BOOLEAN:
                        return String.valueOf(this.cell.getBooleanCellValue() );
                case NUMERIC:
                        return String.valueOf(this.cell.getNumericCellValue() );
                case STRING:
                        return this.cell.getStringCellValue();
            }
        }
        return null;        
    }
    
    public String getValue(XSSFCell cell){
        if (cell != null)
        {
            switch (cell.getCellType())
            {
                case BLANK:
                        return null;
                case BOOLEAN:
                        return String.valueOf(cell.getBooleanCellValue() );
                case NUMERIC:
                        return String.valueOf(cell.getNumericCellValue() );
                case STRING:
                        return cell.getStringCellValue();
            }
        }
        return null;        
    }
    
    // set cell STRING value
    public void setValue(String value){
        this.cell.setCellValue(value);
    }
    
    // set cell NUMERIC value
    public void setValue(double value){
        this.cell.setCellValue(value);
    }
    
    // set cell BOOLEAN value
    public void setValue(boolean value){
        this.cell.setCellValue(value);
    }
    
    // set cell DATE value
    public void setValue(Date value){
        this.cell.setCellValue(value);
    }
    
    // get values list
    public List<List<String>> getValues(){
        int rows = this.sheet.getPhysicalNumberOfRows();
        List<List<String>> rowList = new ArrayList<>();
        for(int i=0;i<rows;i++){
                XSSFRow row = this.sheet.getRow(i);
                if(row==null){
                        continue;
                }
                int cells = row.getPhysicalNumberOfCells();
                List<String> cellList = new ArrayList<>();
                for(int j=0;j<cells;j++){
                        XSSFCell cell = row.getCell(j);
                        String cellValue = cellToString(cell);
                        cellList.add(cellValue);
                }
                rowList.add(cellList);
        }
        return rowList;        
    }
    
    // set values from list
    public void setValues(List<List<String>> values){
        XSSFRow[] rows = new XSSFRow[values.size()];
        for(int i=0;i<values.size();i++){
                List<String> columns = values.get(i);
                rows[i] = this.sheet.createRow(i);
                //this.sheet.setDefaultColumnWidth(columnWidth);
                XSSFCell[] cells = new XSSFCell[columns.size()];
                for(int j=0;j<columns.size();j++){
                    cells[j] = rows[i].createCell(j);
                    //cells[j].setCellStyle(XSSFCellStyle);
                    cells[j].setCellValue(columns.get(j));
                } 
        }        
    }
    
}

