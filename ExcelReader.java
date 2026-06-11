package ExcelDriver;
/**
 * Copyright © 2021-2022 The CETC PHM Authors
 *
 */

import ExcelDriver.Core._Row;
import ExcelDriver.Core._Workbook;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;


public class ExcelReader {
    
    private _Workbook wb;
    
    public ExcelReader(String excelFilePath) {
        super();
        this.wb = new _Workbook();
        this.wb.open(excelFilePath);
    }
    
    // close Excel file
    public void close(){
        this.wb.close();
    }
    
    // read data to list
    // 1 row - header
    // 2-n rows - data
    public List<List<String>> readToList(){
        _Row r = new _Row(this.wb.getWorkbook().getSheetAt(0));
        List<List<String>> rowList = r.readRowsToList();
        return rowList;
    }
    
    // read data to string
    // 1 row - header
    // 2-n rows - data
    public String[][] readToString(){
        List<List<String>> rowList = readToList();
        String[][] rowString = rowList.stream().map((l) -> l.toArray(new String[l.size()])).collect(Collectors.toList()).toArray(new String[rowList.size()][]);
        return rowString;
    }
    
}
