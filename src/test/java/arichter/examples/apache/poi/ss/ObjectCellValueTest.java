package arichter.examples.apache.poi.ss;

import arichter.examples.apache.poi.ss.ObjectCellValue.DateTimeClass;

import java.io.InputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.*;

import static org.junit.Assert.assertTrue;
import org.junit.Test;

public class ObjectCellValueTest {
    
    @Test
    public void testGetCellValue() throws Exception {
        
        ObjectCellValue objectCellValue = new ObjectCellValue();
        objectCellValue.setDateTimeClass(DateTimeClass.JAVA_TIME_LOCALDATETIME);
        
        try (
            InputStream is = getClass().getResourceAsStream("/ss/ExcelExample.xlsx");
            Workbook workbook = WorkbookFactory.create(is);
            ) {
            
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                for (Cell cell : row ) {
                    Object cellValue = objectCellValue.getCellValue(cell, cell.getCellType());
                    System.out.println(cellValue);
                }
            }                
        }
        
        assertTrue(true);

    }
    
    @Test
    public void testSetCellValue() throws Exception {
        
        ObjectCellValue objectCellValue = new ObjectCellValue();
        
        try (
            InputStream is = getClass().getResourceAsStream("/ss/ExcelExample.xlsx");
            Workbook workbook = WorkbookFactory.create(is);
            FileOutputStream out = new FileOutputStream("./ExcelExampleNew.xlsx");
            ) {
            
            Sheet sheet = workbook.createSheet();
            Row row; int r = 0;
            Cell cell; int c = 0;
            
            row = sheet.createRow(r++);
            cell = row.createCell(c);
            objectCellValue.setCellValue(cell, "Text");
            
            row = sheet.createRow(r++);
            cell = row.createCell(c);
            objectCellValue.setCellValue(cell, null);
            
            cell = workbook.getSheetAt(0).getRow(1).getCell(0);
            RichTextString richTextString = (RichTextString)objectCellValue.getCellValue(cell, cell.getCellType());
            row = sheet.createRow(r++);
            cell = row.createCell(c);
            objectCellValue.setCellValue(cell, richTextString);
           
            row = sheet.createRow(r++);
            cell = row.createCell(c);
            objectCellValue.setCellValue(cell, 123);
            
            row = sheet.createRow(r++);
            cell = row.createCell(c);
            objectCellValue.setCellValue(cell, 123.45);
            
            row = sheet.createRow(r++);
            cell = row.createCell(c);
            objectCellValue.setCellValue(cell, true);
            
            row = sheet.createRow(r++);
            cell = row.createCell(c);
            objectCellValue.setCellValue(cell, new java.util.GregorianCalendar());
            
            row = sheet.createRow(r++);
            cell = row.createCell(c);
            objectCellValue.setCellValue(cell, new java.util.Date());

            row = sheet.createRow(r++);
            cell = row.createCell(c);
            objectCellValue.setCellValue(cell, java.time.LocalDate.now());

            row = sheet.createRow(r++);
            cell = row.createCell(c);
            objectCellValue.setCellValue(cell, java.time.LocalDateTime.now());
            
            workbook.write(out);
        }
        
        assertTrue(true);

    }

}
