package ExcelDriver.Core;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Copyright © 2021-2022 The CETC PHM Authors
 *
 */


public class _Workbook {
    
    private static final String DEFAULT_BOOK_NAME = "Book.xlsx";
    
    private String path;
    private String name;
    private XSSFWorkbook workbook;
    
    
    public _Workbook() {
        super();
        //this.workbook = new XSSFWorkbook();
    }
    
    public XSSFWorkbook getWorkbook(){
        return this.workbook;
    }
        
    public String getName(){
        if(this.name == null || this.name.isEmpty()){
            return DEFAULT_BOOK_NAME;
        }else{
            return this.name;
        }
    }
    
    public void setName(String name){
        this.name = name;
    }
    
    public String getPath(){
        return this.path;
    }
    
    public void setPath(String path){
        this.path = path;
    }
    
    public void close() { //)throws Exception{
        try {
            this.workbook.close();
        } catch (IOException e) {
            System.out.println("Error to close Workbook file.");
        }
    }
    
    public void create() {//throws Exception {
        this.workbook = new XSSFWorkbook();
    }
    
    
    public void save() { //throws Exception {
        if(this.name == null || this.name.isEmpty()){
            this.name = DEFAULT_BOOK_NAME;
        }
        try {
            FileOutputStream out = new FileOutputStream(new File(this.path + '\\' + this.name));
            this.workbook.write(out);
            out.close();
            System.out.println(this.name + " save successfully");
        } catch (IOException e) {
            System.out.println(this.name + " not save successfully");
        }            
    }
    
    public void save(String excelFile) { //throws Exception {
        try {
            FileOutputStream out = new FileOutputStream(new File(excelFile));
            this.workbook.write(out);
            out.close();
            System.out.println(this.name + " save successfully");
        } catch (IOException e) {
            System.out.println(this.name + " not save successfully");
        }            
    }
    
    public void open() { //throws Exception {
        if(this.name == null || this.name.isEmpty()){
            this.name = DEFAULT_BOOK_NAME;
        }
        File file = new File(this.path + "/" + this.name);
        try{
        FileInputStream fis = new FileInputStream(file);
        
        //Get the workbook instance for XLSX file 
        this.workbook = new XSSFWorkbook(fis);
        } catch (IOException e) {
            System.out.println("Error to open Workbook file.");
        }            

        if(file.isFile() && file.exists()) {
           System.out.println("openworkbook.xlsx file open successfully.");
        }
    }
    
    public void open(String excelFile) { //throws Exception {
        File file = new File(excelFile);
        try{
        FileInputStream fis = new FileInputStream(file);
        
        //Get the workbook instance for XLSX file 
        this.workbook = new XSSFWorkbook(fis);
        } catch (IOException e) {
            System.out.println("Error to open Workbook file.");
        }            

        if(file.isFile() && file.exists()) {
           System.out.println("openworkbook.xlsx file open successfully.");
        }
    }
    
    // OTHER FUNCTIONAL

    
}
