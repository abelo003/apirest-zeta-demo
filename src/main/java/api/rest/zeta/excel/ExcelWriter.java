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
    
    private static SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");

    private static String[] columns = {"Error Nivelación de datos (PEDT008)", "Error Límite de cajero (MPA1014)", "Error Cambio de PAN T. Física (MPE0026)"};
    private static List<Employee> employees =  new ArrayList<>();

	// Initializing employees data to insert into the excel file
    static {
        Calendar dateOfBirth = Calendar.getInstance();
        dateOfBirth.set(1992, 7, 21);
        employees.add(new Employee("Rajeev Singh", "rajeev@example.com", 
                dateOfBirth.getTime(), 1200000.0));

        dateOfBirth.set(1965, 10, 15);
        employees.add(new Employee("Thomas cook", "thomas@example.com", 
                dateOfBirth.getTime(), 1500000.0));

        dateOfBirth.set(1987, 4, 18);
        employees.add(new Employee("Steve Maiden", "steve@example.com", 
                dateOfBirth.getTime(), 1800000.0));
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        // Create a Workbook
        Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

        /* CreationHelper helps us create instances of various things like DataFormat, 
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        CreationHelper createHelper = workbook.getCreationHelper();

        // Create a Sheet
        Sheet sheet = workbook.createSheet("Reporte " + sdf.format(new Date()));

        // Create a Font for styling header cells
        Font infoFont = workbook.createFont();
        infoFont.setBold(true);
        infoFont.setFontHeightInPoints((short) 16);
        infoFont.setColor(IndexedColors.BLUE.getIndex());
        
        Font headerFont = workbook.createFont();
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        
        CellStyle infoCellStyle = workbook.createCellStyle();
        infoCellStyle.setFont(infoFont);
        infoCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        infoCellStyle.setAlignment(HorizontalAlignment.CENTER);

        // Create a Row
        Row infoRow = sheet.createRow(1);
        Cell cellInfo = infoRow.createCell(2);
        cellInfo.setCellValue("Reporte de errores del día " + sdf.format(new Date()));
        cellInfo.setCellStyle(infoCellStyle);
        sheet.addMergedRegion(new CellRangeAddress( 1, 1, 2, 5));
        
        // Create a Row
        Row headerRow = sheet.createRow(3);

        // Create cells
        for(int i = 0; i < columns.length; i++ ) {
            Cell cell = headerRow.createCell(((i + 1) * 2) - 1);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
            //Agregar la combinacion de celdas
//            sheet.addMergedRegion(new CellRangeAddress( 1, 1, 2, 5));
        }

        // Create Cell Style for formatting Date
//        CellStyle dateCellStyle = workbook.createCellStyle();
//        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));

        // Create Other rows and cells with employees data
//        int rowNum = 4;
//        for(Employee employee: employees) {
//            Row row = sheet.createRow(rowNum++);
//
//            row.createCell(0)
//                    .setCellValue(employee.getName());
//
//            row.createCell(1)
//                    .setCellValue(employee.getEmail());
//
//            Cell dateOfBirthCell = row.createCell(2);
//            dateOfBirthCell.setCellValue(employee.getDateOfBirth());
//            dateOfBirthCell.setCellStyle(dateCellStyle);
//
//            row.createCell(3)
//                    .setCellValue(employee.getSalary());
//        }

		// Resize all columns to fit the content size
        for(int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("C:\\Users\\Z048152\\Documents\\poi-generated-file.xlsx");
        workbook.write(fileOut);
        fileOut.close();
        System.out.println("done!!!");
        // Closing the workbook
        workbook.close();
    }
}
