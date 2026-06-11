package ExcelDriver.Core;
/**
 * Copyright © 2021-2022 The CETC PHM Authors
 *
 */
import java.util.ArrayList;

import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class _Row {
    
    XSSFRow row;
    XSSFSheet sheet;
    private static XSSFRow rowTitle;
    
    public _Row(XSSFSheet sheet) {
        this.sheet = sheet;
        this.row = this.sheet.getRow(0);
    }
    
    public _Row(XSSFSheet sheet,  int numRows){
        this.sheet = sheet;
        this.row = this.sheet.createRow(numRows);
    }
    
    //set current row
    public void setRow(int row){
        this.row = this.sheet.getRow(row);
    }
    
    // The top row in the visible view when the sheet is first viewed after opening it in a viewer
    public short getTopRow(){
        return this.sheet.getTopRow();
    }
    
    // Set the row's height in points
    public void setHeightPoint(float height){
        this.row.setHeightInPoints(height);
    }
    
    //Returns row height measured in point size
    public float getHeightPoint(){
        return this.row.getHeightInPoints();
    }
    
    // Set the row number of this row
    public void setNumber(int rowIndex){
        this.row.setRowNum(rowIndex);
    }
    
    public int getNumber(){
        return this.row.getRowNum();
    }
    
    //Get the 0-based number of the first cell contained in this row
    public short getFirstNumber(){
        return this.row.getFirstCellNum();
    }
    
    // Gets the index of the last cell contained in this row PLUS ONE
    public short getLastNumber(){
        return this.row.getLastCellNum();
    }
    
    // Returns the whole-row cell style
    public XSSFCellStyle getStyle(){
        return this.row.getRowStyle();
    }
    
    //Applies a whole-row cell styling to the row
    public void setStyle(CellStyle style){
        this.row.setRowStyle(style);
    }
    
    // get Sheet
    public XSSFSheet getSheet(){
        return this.sheet;
    }
    
    //Returns the cell at the given (0 based) index, with the Row.MissingCellPolicy from the parent Workbook
    public XSSFCell getCell(int cellnum){
        return this.row.getCell(cellnum);
    }
    
    // Use this to create new cells within the row and return it
    public XSSFCell createCell(int columnIndex){
        return this.row.createCell(columnIndex);
    }
    
    //Use this to create new cells within the row and return it
    public XSSFCell createCell(int columnIndex, CellType type){
        return this.row.createCell(columnIndex, type);
    }
    
    //read rows to array
    public ArrayList readRowsToArray(){
        XSSFRow rows;
        this.sheet = this.sheet.getWorkbook().getSheetAt(0);
        String[] cell = new String[this.sheet.getPhysicalNumberOfRows() + 1];
        ArrayList<String> arrayList = new ArrayList<>();
        for (int i = this.sheet.getFirstRowNum() + 1; i < this.sheet.getPhysicalNumberOfRows(); i++) {
            cell[i] = "";
            rows = this.sheet.getRow(i);
            for (int j = rows.getFirstCellNum(); j < rows.getPhysicalNumberOfCells(); j++) {
                cell[i] += rows.getCell(j).toString();
                cell[i] += " | ";
            }
            arrayList.add(cell[i]);
        }
        return arrayList;
    }
    
    //read rows to lists string
    public List<List<String>> readRowsToList(){
        int rows = this.sheet.getPhysicalNumberOfRows();
        List<List<String>> rowList = new ArrayList<>();
        for(int i=0;i<rows;i++){
                XSSFRow r = this.sheet.getRow(i);
                if(r == null){
                        continue;
                }
                int cells = r.getPhysicalNumberOfCells();
                List<String> cellList = new ArrayList<>();
                for(int j=0;j<cells;j++){
                        XSSFCell cell = r.getCell(j);
                        cell.setCellType(CellType.STRING);
                        String cellValue = cell.getStringCellValue();
                        cellList.add(cellValue);
                }
                rowList.add(cellList);
        }
        return rowList;
    }
    
    // is blank rows
    public boolean isBlankRow(int index, int rowCount){
        if(this.row == null)
                return true;
        for(int i=index; i < rowCount; i++){
                if(this.row.getCell(i) != null || 
                                !"".equals(this.row.getCell(i).getStringCellValue().trim())){
                        return false;
                }
        }
        return true;
    }
    
    // TITLE create
    public void titleCreate(String title, int rowStart, int colStart){
        rowTitle = this.sheet.createRow(rowStart);
        rowTitle.createCell(colStart).setCellValue(title);
    }
    
    // TITLE set style
    public void titleSetStyle(CellStyle style){
        if (!(null == rowTitle)){
            rowTitle.setRowStyle(style);
        }
    }
    
    // TITLE set height
    public void titleSetHeight(int heightPoint){
        if (!(null == rowTitle)){
            rowTitle.setHeightInPoints(heightPoint);
        }
    }
    
    // TITLE get next empty row
    public int titleGetNextRow(){
        if (!(null == rowTitle)){
            return rowTitle.getRowNum();
        }
        return 0;
    }
    
    
    
    
    
    

}
