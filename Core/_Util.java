package ExcelDriver.Core;
/**
 * Copyright © 2021-2022 The CETC PHM Authors
 *
 */
import java.io.BufferedWriter;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import java.io.FilenameFilter;
import java.io.IOException;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.ArrayList;

import javax.xml.transform.TransformerException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.extractor.XSSFExportToXml;
import org.apache.poi.xssf.usermodel.XSSFMap;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.xml.sax.SAXException;


public class _Util {
    
    public static final int EXCEL_STYLE_ESCAPING = 0;
    public static final int UNIX_STYLE_ESCAPING = 1;
    
    public _Util() {
        //super();
    }
    
    // convert to XML
    public void toXML(String pathExcel){
             XSSFWorkbook wb = null;
        try {
            wb = new XSSFWorkbook(new File(pathExcel));
            for (XSSFMap map : wb.getCustomXMLMappings()) {
                    XSSFExportToXml exporter = new XSSFExportToXml(map);

                    ByteArrayOutputStream os = new ByteArrayOutputStream();
                try {
                    exporter.exportToXML(os, true);
                } catch (SAXException | TransformerException e) {
                }
                String xml = os.toString("UTF-8");
                    System.out.println(xml);
                }  
        } catch (IOException | InvalidFormatException e) {
        }
    }
    
    // convert to CSV
    public void toCSV(String pathExcel, String pathCSV, String separator, int format) throws FileNotFoundException, IOException{
        ArrayList<ArrayList<String>> csvData;
        String CSV_FILE_EXTENSION = ".csv";
        Workbook workbook;
        DataFormatter formatter;
        FormulaEvaluator evaluator;
        Sheet sheet;
        Row row;
        int lastRowNum;
        int maxRowWidth = 0;
        int formattingConvention;
        

        // check that the source file/folder exists.
        File source = new File(pathExcel);
        if(!source.exists()) {
            throw new IllegalArgumentException("The source for the Excel " +
                    "file(s) cannot be found at " + source);
        }
        File destination = new File(pathCSV);
        if(!destination.exists()) {
            throw new IllegalArgumentException("The destination directory " + destination + " for the " +
                    "converted CSV file(s) does not exist.");
        }
        if(!destination.isDirectory()) {
            throw new IllegalArgumentException("The destination " + destination + " for the CSV " +
                    "file(s) is not a directory/folder.");
        }
        
        if(format != EXCEL_STYLE_ESCAPING &&
                format != UNIX_STYLE_ESCAPING) {
            throw new IllegalArgumentException("The value passed to the " +
                    "format parameter is out of range: " + format + ", expecting one of " +
                    EXCEL_STYLE_ESCAPING + " or " + UNIX_STYLE_ESCAPING);
        }
        formattingConvention = format;

        // Check to see if the sourceFolder variable holds a reference to
        // a file or a folder full of files.
        final File[] filesList;
        if(source.isDirectory()) {
            // Get a list of all of the Excel spreadsheet files (workbooks) in
            // the source folder/directory
            filesList = source.listFiles(new ExcelFilenameFilter());
        }
        else {
            filesList = new File[]{source};
        }

        if (filesList != null) {
            for(File excelFile : filesList) {
                // open the workbook
                try (FileInputStream fis = new FileInputStream(excelFile)) {
                    workbook = WorkbookFactory.create(fis);
                    evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                    formatter = new DataFormatter(true);
                }

                // Convert it's contents into a CSV file
                csvData = new ArrayList<>();
                
                // Discover how many sheets there are in the workbook....
                int numSheets = workbook.getNumberOfSheets();

                // and then iterate through them.
                for(int i = 0; i < numSheets; i++) {

                    // Get a reference to a sheet and check to see if it contains any rows
                    sheet = workbook.getSheetAt(i);
                    if(sheet.getPhysicalNumberOfRows() > 0) {
                        // from the very first row - number 1 - even if it is missing
                        lastRowNum = sheet.getLastRowNum();
                        for(int j = 0; j <= lastRowNum; j++) {
                            row = sheet.getRow(j);
 
                            Cell cell;
                            int lastCellNum;
                            ArrayList<String> csvLine = new ArrayList<>();

                            if(row != null) {

                                lastCellNum = row.getLastCellNum();
                                for(int k = 0; k <= lastCellNum; k++) {
                                    cell = row.getCell(k);
                                    if(cell == null) {
                                        csvLine.add("");
                                    }
                                    else {
                                        if(cell.getCellType() != CellType.FORMULA) {
                                            csvLine.add(formatter.formatCellValue(cell));
                                        }
                                        else {
                                            csvLine.add(formatter.formatCellValue(cell, evaluator));
                                        }
                                    }
                                }
                                if(lastCellNum > maxRowWidth) {
                                    maxRowWidth = lastCellNum;
                                }
                            }
                            csvData.add(csvLine);
                            //
                        }
                    }
                }

                // Simply replace the .xls or .xlsx file extension with .csv
                String destinationFilename = excelFile.getName();
                destinationFilename = destinationFilename.substring(0, destinationFilename.lastIndexOf('.')) + CSV_FILE_EXTENSION;

                // Save the CSV file
                ArrayList<String> line;
                StringBuilder buffer;
                String csvLineElement;
                
                File save = new File(destination, destinationFilename);

                // Open a writer onto the CSV file.
                try (BufferedWriter bw = Files.newBufferedWriter(save.toPath(), StandardCharsets.ISO_8859_1)) {

                    System.out.println("Saving the CSV file [" + save.getName() + "]");

                    // Step through the elements of the ArrayList
                    for(int i = 0; i < csvData.size(); i++) {
                        buffer = new StringBuilder();

                        // Get an element from the ArrayList
                        line = csvData.get(i);
                        for(int j = 0; j < maxRowWidth; j++) {
                            if(line.size() > j) {
                                csvLineElement = line.get(j);
                                if(csvLineElement != null) {
                                    buffer.append(escapeCharacters(csvLineElement, separator, format));
                                }
                            }
                            if(j < (maxRowWidth - 1)) {
                                buffer.append(separator);
                            }
                        }

                        // Once the line is built, write it away to the CSV file.
                        bw.write(buffer.toString().trim());

                        // Condition the inclusion of new line characters so as to
                        // avoid an additional, superfluous, new line at the end of
                        // the file.
                        if(i < (csvData.size() - 1)) {
                            bw.newLine();
                        }
                    }
                }

            }
        }

    }
    
 
    private String escapeCharacters(String field, String separator, int format) {
        StringBuilder buffer;

        // If the fields contents should be formatted to confrom with Excel's convention....
        if(format == EXCEL_STYLE_ESCAPING) {

            if(field.contains("\"")) {
                buffer = new StringBuilder(field.replaceAll("\"", "\\\"\\\""));
                buffer.insert(0, "\"");
                buffer.append("\"");
            }
            else {
                buffer = new StringBuilder(field);
                if((buffer.indexOf(separator)) > -1 ||
                        (buffer.indexOf("\n")) > -1) {
                    buffer.insert(0, "\"");
                    buffer.append("\"");
                }
            }
            return(buffer.toString().trim());
        }
        else {
            if(field.contains(separator)) {
                field = field.replaceAll(separator, ("\\\\" + separator));
            }
            if(field.contains("\n")) {
                field = field.replaceAll("\n", "\\\\\n");
            }
            return(field);
        }
    }


    // File class and that object refers to a folder/directory
    static class ExcelFilenameFilter implements FilenameFilter {
        @Override
        public boolean accept(File file, String name) {
            return(name.endsWith(".xls") || name.endsWith(".xlsx"));
        }
    }
    
    

    
}
