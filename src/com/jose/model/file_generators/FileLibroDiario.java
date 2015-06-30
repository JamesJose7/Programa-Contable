package com.jose.model.file_generators;

import com.jose.model.LibroDiario;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

/**
 * Created by agua on 30/06/15.
 */
public class FileLibroDiario {
    LibroDiario mLibroDiario;
    File mLibroDiarioFile;
    XSSFWorkbook mWorkbook;


    public void openFile() {
        mLibroDiarioFile = new File("libroDiario.xlsx");

        if (!mLibroDiarioFile.exists()) {
            createSheet();
        } else {
            readSheet();
        }
    }

    private void readSheet() {
        try {

            FileInputStream file = new FileInputStream(new File("libroDiario.xlsx"));

            //Create a workbook from the file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            //Get the desired sheet
            XSSFSheet sheet = workbook.getSheetAt(0);

            //Iterate through rows
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                //Counter for columns
                int colsCounter = 0;

                Row row = rowIterator.next();

                //Iterate through cells
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();



                    //Check cell types
                    switch (cell.getCellType())
                    {
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t\t");
                            break;
                        case Cell.CELL_TYPE_STRING:
                            if (cell.getStringCellValue().equals("-")) {
                                System.out.print("Blank \t\t");
                            } else if (containsHeaders(cell.getStringCellValue())) {
                                //Do Nothing
                            } else {
                                System.out.print(cell.getStringCellValue() + "\t\t");
                            }
                            break;
                    }
                }
                System.out.println();
            }
            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private boolean containsHeaders(String header) {
        String[] headers = {"Libro Diario", "Fecha", "Nombre de Cuentas", "Debitos", "Creditos"};

        for (String element : headers) {
            if (element.equals(header))
                return true;
        }

        return false;
    }

    private void createSheet() {
        //Blank workbook
        mWorkbook = new XSSFWorkbook();

        //Create a blank sheet
        XSSFSheet sheet = mWorkbook.createSheet("Libro Diario");

        //Data to be written

        Row row = sheet.createRow((short) 0);
        Cell cell = row.createCell((short) 0);
        createCell(mWorkbook, row, (short) 0, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER);
        cell.setCellValue("Libro Diario");

        Row row2 = sheet.createRow((short) 1);
        Cell cell2 = row2.createCell((short) 0);
        Cell cell3 = row2.createCell((short) 1);
        Cell cell4 = row2.createCell((short) 2);
        Cell cell5 = row2.createCell((short) 3);
        Cell cell6 = row2.createCell((short) 4);

        cell2.setCellValue("Fecha");
        cell3.setCellValue("Nombre de Cuentas");
        cell4.setCellValue("");
        cell5.setCellValue("Debitos");
        cell6.setCellValue("Creditos");


        sheet.addMergedRegion(new CellRangeAddress(
                0, //first row (0-based)
                0, //last row  (0-based)
                0, //first column (0-based)
                4  //last column  (0-based)
        ));

        try {
            //create a new file
            FileOutputStream fos = new FileOutputStream(mLibroDiarioFile);
            //write workBook
            mWorkbook.write(fos);
            fos.close();
            System.out.println("File created");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }




    /**
     * Creates a cell and aligns it a certain way.
     *
     * @param wb     the workbook
     * @param row    the row to create the cell in
     * @param column the column number to create the cell in
     * @param halign the horizontal alignment for the cell.
     */
    private static void createCell(Workbook wb, Row row, short column, short halign, short valign) {
        Cell cell = row.createCell(column);
        cell.setCellValue("Align It");
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(halign);
        cellStyle.setVerticalAlignment(valign);
        cell.setCellStyle(cellStyle);
    }
}
