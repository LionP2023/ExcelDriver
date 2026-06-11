package ExcelDriver.Core;
/**
 * Copyright © 2021-2022 The CETC PHM Authors
 *
 */
public class _Formula {
    
    private static String f_ABS = "ABS";
    private static String f_AVG = "AVERAGE";
    private static String f_COUNT = "COUNT";
    private static String f_MAX = "MAX";
    private static String f_MIN = "MAX";
    private static String f_SUM = "SUM";
    private static String f_CURDATE = "TODAY()";
    
    private String formula;
    private String address;
    
    
    public _Formula() {
        super();
    }
    
    // get last formula
    public String getFormula(){
        return this.formula;
    }
    
    // get last address
    public String getAddress(){
        return this.address;
    }
    
    // create address
    // A1
    public String createAddress(String columnName, int rowNum){
        this.address = columnName + rowNum;
        return this.address;            
    }
    
    // A1:A8
    public String createAddress(String columnName1, int rowNum1, String columnName2, int rowNum2){
        this.address = columnName1 + rowNum1 + ":" + columnName2 + rowNum2;
        return this.address;            
    }
    
    // A1,A2,A3,A4,A5,A6,A7,A8
    public String createAddress(String []array){
        if(array.length>0){
            int cnt = array.length;
            for (int i=0; i<cnt; i++){
                this.address += array[i];
                if (i<cnt-1){
                    this.address += ",";
                }
            }
        }
        return this.address;
    }          
    
    // create SUM formula
    // A1:A2 A1-begin address, A2-end address
    public String createSUM(String addrBegin, String addrEnd){
        this.formula = f_SUM + "(" + addrBegin + ":" + addrEnd + ")";
        return this.formula;
    }
    
    // A1:A2
    public String createSUM(String address){
        this.formula = f_SUM + "(" + address + ")";
        return this.formula;
    }
    
    // A1:B2 A-symbol1, 1-value1, B-symbol2, 2-value2
    public String createSUM(String symbol1, int value1, String symbol2, int value2){
        this.formula = f_SUM + "(" + symbol1 + value1 + ":" + symbol2 + value2 + ")";
        return this.formula;
    }
    
    // adding SUM multiple addresses
    public String createMultiSUM(String[] array){
        if(array.length>0){
            this.formula = f_SUM + "(";
            int cnt = array.length;
            for (int i=0; i<cnt; i++){
                this.formula += array[i];
                if (i<cnt-1){
                    this.formula += ",";
                }
                this.formula += ")";
            }
        }
        return this.formula;
    }
    
    // create AVERAGE formula
    public String createAVG(String addrBegin, String addrEnd){
        this.formula = f_AVG + "(" + addrBegin + ":" + addrEnd + ")";
        return this.formula;
    }
    
    public String createAVG(String address){
        this.formula = f_AVG + "(" + address + ")";
        return this.formula;
    }
    
    public String createAVG(String symbol1, int value1, String symbol2, int value2){
        this.formula = f_AVG + "(" + symbol1 + value1 + ":" + symbol2 + value2 + ")";
        return this.formula;
    }
    
    // adding AVG multiple addresses
    public String createMultiAVG(String[] array){
        if(array.length>0){
            this.formula = f_AVG + "(";
            int cnt = array.length;
            for (int i=0; i<cnt; i++){
                this.formula += array[i];
                if (i<cnt-1){
                    this.formula += ",";
                }
                this.formula += ")";
            }
        }
        return this.formula;
    }
    
    // create ABS formula
    // A1 - address
    public String createABS(String address){
        this.formula = f_ABS + "(" + address + ")";
        return this.formula;
    }
    
    // create MAX formula
    // A1:A2 A1-begin address, A2-end address
    public String createMAX(String addrBegin, String addrEnd){
        this.formula = f_MAX + "(" + addrBegin + ":" + addrEnd + ")";
        return this.formula;
    }
    
    // A1 - address
    public String createMAX(String address){
        this.formula = f_MAX + "(" + address + ")";
        return this.formula;
    }
    
    // adding MAX multiple addresses
    public String createMultiMAX(String[] array){
        if(array.length>0){
            this.formula = f_MAX + "(";
            int cnt = array.length;
            for (int i=0; i<cnt; i++){
                this.formula += array[i];
                if (i<cnt-1){
                    this.formula += ",";
                }
                this.formula += ")";
            }
        }
        return this.formula;
    }
    
    // create MIN formula
    // A1:A2 A1-begin address, A2-end address
    public String createMIN(String addrBegin, String addrEnd){
        this.formula = f_MIN + "(" + addrBegin + ":" + addrEnd + ")";
        return this.formula;
    }
    
    // A1 - address
    public String createMIN(String address){
        this.formula = f_MIN + "(" + address + ")";
        return this.formula;
    }
    
    // adding MIN multiple addresses
    public String createMultiMIN(String[] array){
        if(array.length>0){
            this.formula = f_MIN + "(";
            int cnt = array.length;
            for (int i=0; i<cnt; i++){
                this.formula += array[i];
                if (i<cnt-1){
                    this.formula += ",";
                }
                this.formula += ")";
            }
        }
        return this.formula;
    }
    
    // create COUNT formula
    // A1:A2 A1-begin address, A2-end address
    public String createCOUNT(String addrBegin, String addrEnd){
        this.formula = f_COUNT + "(" + addrBegin + ":" + addrEnd + ")";
        return this.formula;
    }
    
    // A1 - address
    public String createCOUNT(String address){
        this.formula = f_COUNT + "(" + address + ")";
        return this.formula;
    }
    
    // adding COUNT multiple addresses
    public String createMultiCOUNT(String[] array){
        if(array.length>0){
            this.formula = f_COUNT + "(";
            int cnt = array.length;
            for (int i=0; i<cnt; i++){
                this.formula += array[i];
                if (i<cnt-1){
                    this.formula += ",";
                }
                this.formula += ")";
            }
        }
        return this.formula;
    }
    
    // create CURRENT DATE
    public String createCurrentDate(){
        return f_CURDATE;
    }
    
    // create multiplication by a number
    // A1*12
    public String createMultByNumber(String address, int value){
        this.formula = address + "*" + value;
        return this.formula;
    }
    
    // A1*A2*A4*B4*12
    public String createMultByNumber(String []array, int value){
        if(array.length>0){
            int cnt = array.length;
            for (int i=0; i<cnt; i++){
                this.formula += array[i];
                this.formula += "*";
            }
            this.formula += value;
        }
        return this.formula;
    }
    
}
