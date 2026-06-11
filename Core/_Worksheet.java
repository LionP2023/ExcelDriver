package ExcelDriver.Core;
/**
 * Copyright © 2021-2022 The CETC PHM Authors
 *
 */
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class _Worksheet {
    
    private static final String DEFAULT_SHEET_NAME = "Sheet";
    
    private static int sheetNumber = 0;
    
    XSSFWorkbook workbook;
    protected XSSFSheet sheet;

    public _Worksheet(XSSFWorkbook workbook) {
        super();
        this.sheet = workbook.createSheet(DEFAULT_SHEET_NAME + ' ' + ExcelDriver.Core
                                                                                ._Worksheet
                                                                                .sheetNumber);
    }
    
    public _Worksheet(XSSFWorkbook workbook, String sheetName) {
        this.sheet = workbook.createSheet(sheetName);
    }
    
    public XSSFSheet getSheet(){
        return this.sheet;
    }
    
    public String getName(){
        return this.sheet.getSheetName();
    }
    
    public void setName(String sheetName){
        this.workbook.setSheetName(ExcelDriver.Core._Worksheet.sheetNumber, sheetName);
    }
    
    public int getNumber(){
        return ExcelDriver.Core._Worksheet.sheetNumber;
    }
    
    public void setNumber(int sheetNumber){
        ExcelDriver.Core._Worksheet.sheetNumber = sheetNumber;
    }
    
    // get rows number
    public int getRows(){
        return this.sheet.getLastRowNum() + 1;
    }

    // get columns number
    public int getColumns(){
        return this.sheet.getRow(0).getLastCellNum();
    }
    
    //public void addSheet(XSSFWorkbook workbook, String sheetName){
    //    this.sheet = workbook.createSheet(sheetName);
    //    Core.Worksheet.sheetNumber++;
    //}
    
    // Sheet protection
    public void setProtection(boolean enable){
        if(enable){
            this.sheet.enableLocking();
        }else{
            this.sheet.disableLocking();
        }
    }
    
    // Returns any tables associated with this Sheet
    public java.util.List<XSSFTable> getTables(){
        return this.sheet.getTables();
    }
    
    // Returns the parent XSSFWorkbook
    public XSSFWorkbook getWorkbook(){
        return this.sheet.getWorkbook();
    }
    
    // Returns a flag indicating whether this sheet is selected
    public boolean isSelectWorksheet(){
        return this.sheet.isSelected(); 
    }
    
    //Sets a flag indicating whether this sheet is selected
    public void setSelectWorksheet(boolean b){
        this.sheet.setSelected(b);
    }
    
    // Set background color of the sheet tab
    public void setBackgroundColor(XSSFColor color){
        this.sheet.setTabColor(color);
    }
    
    //Window zoom magnification for current view representing percent values
    public void setZoom(int percent){
        this.sheet.setZoom(percent);
    }
    
    // get number rows
    //public int getRows(){
    //    return this.sheet.getPhysicalNumberOfRows();
    //}
    
    
    // BEGIN COLUMN
    // Adjusts the column width to fit the contents
    public void columnAutoSize(int column){
        this.sheet.autoSizeColumn(column);
    }
    
    public void columnAutoSize(int column, boolean useMergedCells){
        this.sheet.autoSizeColumn(column, useMergedCells);
    }
    
    //The left col in the visible view when the sheet is first viewed after opening it in a viewer
    public short getLeftCol(){
        return this.sheet.getLeftCol();
    }
    
    // Set the width (in units of 1/256th of a character width)
    public void setColWidth(int columnIndex, int width){
        this.sheet.setColumnWidth(columnIndex, width);
    }
    
    // Specifies the number of characters of the maximum digit width of the normal style's font
    public void setDefaultColWidth(int columnIndex, int width){
        this.sheet.setDefaultColumnWidth(width);
    }
    
    // END COLUMN
    
    
    
    
    
    // BEGIN ROW
    // Create a new row
    public XSSFRow createRow(int rownum){
        return this.sheet.createRow(rownum);
    }
    
    // Get the default row height for the sheet (if the rows do not define their own height) in twips (1/20 of a point)
    public short getRowHeight(){
        return this.sheet.getDefaultRowHeight();
    }
    
    //Set the default row height for the sheet (if the rows do not define their own height) in twips (1/20 of a point)
    public void setRowHeight(short height){
        this.sheet.setDefaultRowHeight(height);
    }
    
    //Get the default row height for the sheet measued in point size (if the rows do not define their own height)
    public float getRowHeightPoint(){
        return this.sheet.getDefaultRowHeightInPoints();
    }
    
    //Sets default row height measured in point size
    public void setRowHeightPoint(float height){
        this.sheet.setDefaultRowHeightInPoints(height);
    }
    
    // Gets the first row on the sheet
    public int getFirstRow(){
        return this.sheet.getFirstRowNum();
    }
    
    // Gets the last row on the sheet
    public int getLastRow(){
        return this.sheet.getLastRowNum();
    }
    
    //The top row in the visible view when the sheet is first viewed after opening it in a viewer
    public short getTopRow(){
        return this.sheet.getTopRow();
    }
    // END ROW
    
    
    
    
    
    // BEGIN OTHER
    // Creates a split (freezepane)
    public void freezeCreate(int colSplit, int rowSplit){
        this.sheet.createFreezePane(colSplit, rowSplit);
    }
    
    public void freezeCreate(int colSplit, int rowSplit, int leftmostColumn, int topRow){
        this.sheet.createFreezePane(colSplit, rowSplit, leftmostColumn, topRow);
    }
    
    // Creates a split pane
    public void splitCreate(int xSplitPos, int ySplitPos, int leftmostColumn, int topRow, int activePane){
        this.sheet.createSplitPane(xSplitPos, ySplitPos, leftmostColumn, topRow, activePane);
    }
    
    // Flag indicating whether the Fit to Page print option is enabled
    public boolean isFitToPage(){
        return this.sheet.getFitToPage();
    }
    
    //Flag indicating whether the Fit to Page print option is enabled
    public void setFotToPage(boolean b){
        this.sheet.setFitToPage(b);
    }
    
    
    
    
    // END OTHER
}
