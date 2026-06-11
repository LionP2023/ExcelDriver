package ExcelDriver.Core;
/**
 * Copyright © 2021-2022 The CETC PHM Authors
 *
 */

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class _Font {
    
    XSSFFont font;
    
    public _Font(XSSFWorkbook workBook) {
        super();
        font = workBook.createFont();
    }
    
    public _Font(XSSFWorkbook workBook, String fontName, short fontHeight){
        font = workBook.createFont();
        font.setFontName(fontName);
        font.setFontHeightInPoints(fontHeight);
    }
    
    public _Font(XSSFCellStyle style){
        font = style.getFont();
    }
    
    public XSSFFont getFont(){
        return this.font;
    }
    
    public String getFontName(){
        return this.font.getFontName();
    }
    
    public void setFontName(String fontName){
        this.font.setFontName(fontName);
    }
    
    public short getFontHeight(){
        return this.font.getFontHeight();
    }
    
    public void setFontHeight(short height){
        this.font.setFontHeight(height);
    }
    
    public short getFontSizePoint(){
        return this.font.getFontHeightInPoints();
    }
    
    public void setFontSizePoint(short fontSize){
        this.font.setFontHeightInPoints(fontSize);
    }
    
    public short getFontColor(){
        return this.font.getColor();
    }
    
    public void setFontColor(short fontColor){
        this.font.setColor(fontColor);
    }
    
    public boolean getFontBold(){
        return this.font.getBold();
    }
    
    public void setFontBold(boolean bold){
        this.font.setBold(bold);
    }
    
    public boolean getFontItalic(){
        return this.font.getItalic();
    }
    
    public void setFontItalic(boolean bold){
        this.font.setItalic(bold);
    }
    
    
}
