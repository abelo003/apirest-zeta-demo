/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package api.rest.zeta.excel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import org.apache.poi.ss.util.CellRangeAddress;

public class ExcelWriter {
    
    private static final SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");

    private static final String[] columns = {"Error Nivelación de datos (PEDT008)", "Error Límite de cajero (MPA1014)", "Error Cambio de PAN T. Física (MPE0026)"};

    public static void main(String[] args) throws IOException, InvalidFormatException {
        // Create a Workbook
        Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

        /* CreationHelper helps us create instances of various things like DataFormat, 
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
//        CreationHelper createHelper = workbook.getCreationHelper();

        // Create a Sheet
        Sheet sheet = workbook.createSheet("Reporte " + sdf.format(new Date()));

        // Create a Font for styling header cells
        Font infoFont = workbook.createFont();
        infoFont.setBold(true);
        infoFont.setFontHeightInPoints((short) 16);
        infoFont.setColor(IndexedColors.BLUE.getIndex());
        
        Font totalFont = workbook.createFont();
        totalFont.setBold(true);
        totalFont.setFontHeightInPoints((short) 13);
        totalFont.setColor(IndexedColors.BLACK.getIndex());
        
        Font headerFont = workbook.createFont();
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        headerCellStyle.setFillBackgroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerCellStyle.setAlignment(HorizontalAlignment.CENTER);
        
        CellStyle infoCellStyle = workbook.createCellStyle();
        infoCellStyle.setFont(infoFont);
        infoCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        infoCellStyle.setAlignment(HorizontalAlignment.CENTER);

        CellStyle totalCellStyle = workbook.createCellStyle();
        totalCellStyle.setFont(totalFont);
        totalCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        totalCellStyle.setAlignment(HorizontalAlignment.CENTER);
        
        // Create a Row
        Row infoRow = sheet.createRow(1);
        Cell cellInfo = infoRow.createCell(1);
        cellInfo.setCellValue("Reporte de errores del día " + sdf.format(new Date()));
        cellInfo.setCellStyle(infoCellStyle);
        
        sheet.addMergedRegion(new CellRangeAddress( 1, 2, 1, 6));
        
        // Create a Row
        Row headerRow = sheet.createRow(3);
        // Create cells
        for(int i = 0; i < columns.length; i++ ) {
            Cell cell = headerRow.createCell(((i + 1) * 2) - 1);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
            sheet.setColumnWidth(cell.getColumnIndex(), 5000);
            sheet.setColumnWidth(cell.getColumnIndex() + 1, 5000);
            //Agregar la combinacion de celdas
            sheet.addMergedRegion(new CellRangeAddress( headerRow.getRowNum(), headerRow.getRowNum() + 1, cell.getColumnIndex(), cell.getColumnIndex() + 1));
            sheet.addMergedRegion(new CellRangeAddress( headerRow.getRowNum() + 2, headerRow.getRowNum() + 2, cell.getColumnIndex(), cell.getColumnIndex() + 1));
        }

        // Create Cell Style for formatting Date
//        CellStyle dateCellStyle = workbook.createCellStyle();
//        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));
        
        
        List<ErrorDetail> errors = new ArrayList<>();
        for (int i = 0; i < 15; i++) {
            errors.add(new ErrorDetail("007635" + String.format("%02d", i), "45672938765263" + String.format("%02d", i)));
        }
        // Create Other rows and cells with employees data
        int rowNum = 6;
        for(ErrorDetail error: errors) {
            Row row = sheet.createRow(rowNum++);
//
            row.createCell(1).setCellValue(error.getBuc());
            row.createCell(2).setCellValue(error.getPan());
//            Cell dateOfBirthCell = row.createCell(2);
//            dateOfBirthCell.setCellValue(employee.getDateOfBirth());
//            dateOfBirthCell.setCellStyle(dateCellStyle);
//
//            row.createCell(3)
//                    .setCellValue(employee.getSalary());
        }

        //Totles
        Row totalRow = sheet.createRow(5);
        totalRow.createCell(1);
        totalRow.createCell(3);
        totalRow.createCell(5);
        totalRow.getCell(1).setCellValue(errors.size());
        totalRow.getCell(1).setCellStyle(totalCellStyle);
        totalRow.getCell(3).setCellValue(errors.size());
        totalRow.getCell(3).setCellStyle(totalCellStyle);
        totalRow.getCell(5).setCellValue(errors.size());
        totalRow.getCell(5).setCellStyle(totalCellStyle);
        // Resize all columns to fit the content size
        for(int i = 0; i < columns.length; i++) {
//            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("C:\\Users\\abelo\\Documents\\excel\\poi-generated-file.xlsx");
        workbook.write(fileOut);
        fileOut.close();
        System.out.println("done!!!");
        // Closing the workbook
        workbook.close();
    }
}
