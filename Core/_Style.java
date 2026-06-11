package ExcelDriver.Core;
/**
 * Copyright © 2021-2022 The CETC PHM Authors
 *
 */
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class _Style {
    
    private static XSSFWorkbook wb;
    private CellStyle style;
    
    public _Style(XSSFWorkbook workbook) {
        wb = workbook;
        style = wb.createCellStyle();
    }
    
    // get CellStyle object
    public CellStyle getCellStyle(){
        return this.style;
    }
    
    // create BORDER style default
    public CellStyle getBorderDefault(){
        BorderStyle thin = BorderStyle.THIN;
        short black = IndexedColors.BLACK.getIndex();

        style.setBorderRight(thin);
        style.setRightBorderColor(black);
        style.setBorderBottom(thin);
        style.setBottomBorderColor(black);
        style.setBorderLeft(thin);
        style.setLeftBorderColor(black);
        style.setBorderTop(thin);
        style.setTopBorderColor(black);
        return style;
    }
    
    // TITLE
    // create TITLE style default
    public CellStyle getTitleDefault(){
        Font titleFont = wb.createFont();
        titleFont.setBold(true);
        titleFont.setFontHeightInPoints((short)16);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFont(titleFont);        
        return style;
    }
    
    // TABLE HEADER
    // create TABLE HEADER style default
    public CellStyle getTableHeaderDefault(){
        Font headerFont = wb.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short)14);
        style = getBorderDefault();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFont(headerFont);  
        return style;
    }
    
    // TABLE DETAILS
    // create TABLE DETAILS FOR NUMBER style default
    // 0
    // 0.00
    // #,##0
    // #,##0.00
    // 0.00E+00
    // 0%
    // 0.00%
    
    public CellStyle getTableDetailsDefault(){
        Font detailsFont = wb.createFont();
        detailsFont.setFontHeightInPoints((short)14);
        style = getBorderDefault();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setWrapText(true);
        style.setFont(detailsFont);  
        return style; 
    }  
    
    public CellStyle getTableDetailsDefaultColor(){
        Font detailsFont = wb.createFont();
        detailsFont.setFontHeightInPoints((short)14);
        style = getBorderDefault();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setWrapText(true);
        style.setFont(detailsFont);  
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style; 
    }    
    
    public CellStyle getTableDetailsNumberDefault(String format){
        Font detailsFont = wb.createFont();
        DataFormat df = wb.createDataFormat();
        detailsFont.setFontHeightInPoints((short)14);
        style = getBorderDefault();
        style.setDataFormat(df.getFormat(format));
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setFont(detailsFont);  
        return style; 
    }
    
    // create TABLE DETAILS FOR STRING style default
    public CellStyle getTableDetailsStringDefault(){
        Font detailsFont = wb.createFont();
        detailsFont.setFontHeightInPoints((short)14);
        style = getBorderDefault();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setWrapText(true);
        style.setFont(detailsFont);  
        return style; 
    }
    
    // create TABLE DETAILS FOR DATA style default
    public CellStyle getTableDetailsDateDefault(){
        Font detailsFont = wb.createFont();
        DataFormat df = wb.createDataFormat();
        detailsFont.setFontHeightInPoints((short)14);
        style = getBorderDefault();
        style.setDataFormat(df.getFormat("m/d/yy"));
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setFont(detailsFont);  
        return style; 
    }
    
    // create TABLE DETAILS FOR DATA and TIME style default
    public CellStyle getTableDetailsDateTimeDefault(){
        Font detailsFont = wb.createFont();
        DataFormat df = wb.createDataFormat();
        detailsFont.setFontHeightInPoints((short)14);
        style = getBorderDefault();
        style.setDataFormat(df.getFormat("m/d/yy h:mm"));
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setFont(detailsFont);  
        return style; 
    }
    
    // TABLE FOOTER
    // create TABLE FOOTER FOR NUMBER style default
    public CellStyle getTableFooterNumberDefault(String format){
        Font footerFont = wb.createFont();
        DataFormat df = wb.createDataFormat();
        footerFont.setFontHeightInPoints((short)14);
        footerFont.setBold(true);
        style = getBorderDefault();
        style.setDataFormat(df.getFormat(format));
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setFont(footerFont);  
        return style; 
    }
    
    // create TABLE FOOTER FOR STRING style default
    public CellStyle getTableFooterStringDefault(){
        Font footerFont = wb.createFont();
        footerFont.setFontHeightInPoints((short)14);
        footerFont.setBold(true);
        style = getBorderDefault();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setWrapText(true);
        style.setFont(footerFont);  
        return style; 
    }

    // create TABLE FOOTER FOR DATA style default
    public CellStyle getTableFooterDateDefault(){
        Font footerFont = wb.createFont();
        DataFormat df = wb.createDataFormat();
        footerFont.setFontHeightInPoints((short)14);
        footerFont.setBold(true);
        style = getBorderDefault();
        style.setDataFormat(df.getFormat("m/d/yy"));
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setFont(footerFont);  
        return style; 
    }
    
    // create TABLE FOOTER FOR DATA and TIME style default
    public CellStyle getTableFooterDateTimeDefault(){
        Font footerFont = wb.createFont();
        DataFormat df = wb.createDataFormat();
        footerFont.setFontHeightInPoints((short)14);
        footerFont.setBold(true);
        style = getBorderDefault();
        style.setDataFormat(df.getFormat("m/d/yy h:mm"));
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setFont(footerFont);  
        return style; 
    }
    
    
    
    
    
    
    // other
    // create a library of cell styles
    private Map<String, CellStyle> createStyles(){
        Map<String, CellStyle> styles = new HashMap<>();
        DataFormat df = wb.createDataFormat();

        CellStyle stl;
        
        // TITLE
        Font titleFont = wb.createFont();
        titleFont.setBold(true);
        titleFont.setFontHeightInPoints((short)16);
        stl = getBorderDefault();
        stl.setAlignment(HorizontalAlignment.CENTER);
        stl.setFont(titleFont);
        styles.put("title", stl);
        
        // HEADER
        Font headerFont = wb.createFont();
        headerFont.setBold(true);
        titleFont.setFontHeightInPoints((short)14);
        stl = getBorderDefault();
        stl.setAlignment(HorizontalAlignment.CENTER);
        stl.setFont(headerFont);
        styles.put("table_header", stl);

        // CELL BOLD
        Font font1 = wb.createFont();
        font1.setBold(true);
        stl = getBorderDefault();
        stl.setAlignment(HorizontalAlignment.LEFT);
        stl.setFont(font1);
        styles.put("cell_bold", stl);

        // CELL BOLD CENTER
        stl = getBorderDefault();
        stl.setAlignment(HorizontalAlignment.CENTER);
        stl.setFont(font1);
        styles.put("cell_bold_center", stl);

        // CELL NORMAL
        stl = getBorderDefault();
        stl.setAlignment(HorizontalAlignment.LEFT);
        stl.setWrapText(true);
        styles.put("cell_normal", stl);

        // CELL NORMAL CENTER
        stl = getBorderDefault();
        stl.setAlignment(HorizontalAlignment.CENTER);
        stl.setWrapText(true);
        styles.put("cell_normal_center", stl);

        // CELL DATE and TIME
        stl = getBorderDefault();
        stl.setAlignment(HorizontalAlignment.RIGHT);
        stl.setWrapText(true);
        stl.setDataFormat(df.getFormat("m/d/yy h:mm"));
        styles.put("cell_date_time", stl);

        return styles;
    }
    

     
    
    
}
