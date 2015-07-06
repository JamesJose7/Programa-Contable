package com.jose.model.file_generators;

import com.jose.model.Asiento;
import com.jose.model.LibroDiario;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.print.Doc;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

/**
 * Created by agua on 30/06/15.
 */
public class FileWorkbook {
    LibroDiario mLibroDiario = new LibroDiario();
    File mLibroDiarioFile;
    XSSFWorkbook mWorkbook;


    public void openFile() {
        mLibroDiarioFile = new File("libroDiario.xlsx");
        //Create a new workbook

        if (!mLibroDiarioFile.exists()) {
            mWorkbook = new XSSFWorkbook();
            createLibroDiarioSheet();
        }

        readSheet();
        testPrintLibroDiario();

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


            //Maps for each account
            Map<String, Double> mapDebe = new HashMap<>();
            Map<String, Double> mapHaber = new HashMap<>();

            //When every "asiento" is finished
            boolean isFinishedAsiento = false;
            boolean debeExists = false;
            boolean haberExists = false;

            //Temporal variables for each row
            String fecha = "";
            String cuentaDebe = "";
            String cuentaHaber = "";
            double debito = 0;
            double credito = 0;
            String registro = "";

            while (rowIterator.hasNext()) {

                Row row = rowIterator.next();

                //Iterate through cells
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();

                    //Check cell types
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:

                            if (row.getRowNum() > 1) {
                                //Store values
                                if (debeExists) {
                                    debito = cell.getNumericCellValue();
                                } else if (haberExists) {
                                    credito = cell.getNumericCellValue();
                                }
                            }
                            break;
                        case Cell.CELL_TYPE_STRING:

                            if (cell.getStringCellValue().equals("-")) {
                                isFinishedAsiento = true;
                            }

                            if (row.getRowNum() > 1) {
                                //Store values

                                if (isFinishedAsiento) {
                                    registro = cell.getStringCellValue();
                                    debeExists = true;
                                    haberExists = false;
                                    break;
                                }

                                if (fecha.equals("")) {
                                    fecha = cell.getStringCellValue();
                                    debeExists = true;
                                    haberExists = false;
                                } else {

                                    if (debeExists) {
                                        cuentaDebe = cell.getStringCellValue();

                                    }

                                    if (haberExists) {
                                        cuentaHaber = cell.getStringCellValue();
                                    }

                                    if (cell.getStringCellValue().equals("*")) {
                                        haberExists = true;
                                        debeExists = false;
                                    }

                                }
                            }
                            break;
                        default:
                            break;
                    }
                }

                if (debeExists) {
                    mapDebe.put(cuentaDebe, debito);
                }

                if (haberExists) {
                    mapHaber.put(cuentaHaber, credito);
                }


                if (isFinishedAsiento) {
                    mapDebe.remove("*");

                    mLibroDiario.addAsiento(new Asiento(fecha, mapDebe, mapHaber, registro));
                    fecha = "";

                    mapDebe = new HashMap<>();
                    mapHaber = new HashMap<>();

                    isFinishedAsiento = false;
                }

            }


            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }


    }

    private void testPrintLibroDiario() {
        List<Asiento> asientoList = mLibroDiario.getAsientos();
        for (Asiento asiento : asientoList) {
            testPrintAsiento(asiento);
            System.out.println("\n");
        }
    }

    private void testPrintAsiento(Asiento asiento) {
        System.out.println("Fecha: " + asiento.getFecha());
        //Cuenta debitos
        System.out.println("Cuenta debe:");
        for (Map.Entry<String, Double> entry : asiento.getDebitos().entrySet()) {
            System.out.printf("\tcuenta: %s\n" +
                            "\tvalor: %.2f\n",
                    entry.getKey(),
                    entry.getValue());
        }
        //Cuenta creditos
        System.out.println("Cuenta creditos:");
        for (Map.Entry<String, Double> entry : asiento.getCreditos().entrySet()) {
            System.out.printf("\tcuenta: %s\n" +
                    "\tvalor: %.2f\n",
                    entry.getKey(),
                    entry.getValue());
        }
        System.out.println("Registro: " + asiento.getRegistro());

    }

    private boolean containsHeaders(String header) {
        String[] headers = {"Libro Diario", "Fecha", "Nombre de Cuentas", "Debitos", "Creditos"};

        for (String element : headers) {
            if (element.equals(header)) {
                return true;
            }
        }

        return false;
    }

    private void createLibroDiarioSheet() {

        //Create sheet for "Libro diario"
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
