package com.jose.model.file_generators;

import com.jose.model.Asiento;
import com.jose.model.LibroDiario;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.dev.XSSFSave;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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
        mLibroDiarioFile = new File("Contabilidad.xlsx");
        //Create a new workbook

        if (!mLibroDiarioFile.exists()) {
            mWorkbook = new XSSFWorkbook();
            List<XSSFSheet> sheets = new ArrayList<>();

            createLibroDiarioSheet();
            createPlanDeCuentas();

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

        readLibroDiarioSheet();
        testPrintLibroDiario();

    }


    private void readLibroDiarioSheet() {
        try {

            FileInputStream file = new FileInputStream(new File("Contabilidad.xlsx"));

            //Create a workbook from the file
            mWorkbook = new XSSFWorkbook(file);

            //Get the desired sheet
            XSSFSheet sheet = mWorkbook.getSheetAt(0);

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

    private void createPlanDeCuentas() {
        XSSFSheet sheet = mWorkbook.createSheet("Plan de Cuentas");

        //Data to be written
        Row row = sheet.createRow((short) 0);
        Cell cell = row.createCell((short) 0);
        createCell(mWorkbook, row, (short) 0, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER);
        cell.setCellValue("Plan de cuentas");

        Row row1 = sheet.createRow((short) 1);
        row1.createCell((short) 0);

        sheet.addMergedRegion(new CellRangeAddress(
                0, //first row (0-based)
                1, //last row  (0-based)
                0, //first column (0-based)
                4  //last column  (0-based)
        ));

        Row row2 = sheet.createRow((short) 2);

        Row row3 = sheet.createRow((short) 3);
        Cell cell31 = row3.createCell((short) 0);
        createCell(mWorkbook, row3, (short) 0, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER);
        cell31.setCellValue("Estado de Situación Financiera");

        sheet.addMergedRegion(new CellRangeAddress(
                3, //first row (0-based)
                3, //last row  (0-based)
                0, //first column (0-based)
                4  //last column  (0-based)
        ));

        sheet.createRow((short) 4);
        Row row5 = sheet.createRow((short) 5);
        Cell cell51 = row5.createCell((short) 0);
        Cell cell52 = row5.createCell((short) 1);
        Row row6 = sheet.createRow((short) 6);
        Cell cell61 = row6.createCell((short) 0);
        Cell cell62 = row6.createCell((short) 1);

        cell51.setCellValue("1.");
        cell52.setCellValue("ACTIVO");
        cell61.setCellValue("1.1");
        cell62.setCellValue("ACTIVO CORRIENTE");

//        sheet.createRow((short) 7);
//        sheet.createRow((short) 8);
//        sheet.createRow((short) 9);
//        sheet.createRow((short) 10);
//        sheet.createRow((short) 11);

        Row row12 = sheet.createRow((short) 12);
        Cell cell121 = row12.createCell((short) 0);
        Cell cell122 = row12.createCell((short) 1);

        cell121.setCellValue("1.2");
        cell122.setCellValue("ACTIVO NO CORRIENTE");

//        sheet.createRow((short) 13);
//        sheet.createRow((short) 14);
//        sheet.createRow((short) 15);
//        sheet.createRow((short) 16);
//        sheet.createRow((short) 17);

        sheet.createRow((short) 18);
        Row row19 = sheet.createRow((short) 19);
        Cell cell191 = row19.createCell((short) 0);
        Cell cell192 = row19.createCell((short) 1);
        Row row20 = sheet.createRow((short) 20);
        Cell cell201 = row20.createCell((short) 0);
        Cell cell202 = row20.createCell((short) 1);

        cell191.setCellValue("2.");
        cell192.setCellValue("PASIVO");
        cell201.setCellValue("2.1");
        cell202.setCellValue("PASIVO CORRIENTE");

//        sheet.createRow((short) 21);
//        sheet.createRow((short) 22);
//        sheet.createRow((short) 23);
//        sheet.createRow((short) 24);
//        sheet.createRow((short) 25);

        Row row26 = sheet.createRow((short) 26);
        Cell cell261 = row26.createCell((short) 0);
        Cell cell262 = row26.createCell((short) 1);

        cell261.setCellValue("2.2");
        cell262.setCellValue("PASIVO NO CORRIENTE");

//        sheet.createRow((short) 27);
//        sheet.createRow((short) 28);
//        sheet.createRow((short) 29);
//        sheet.createRow((short) 30);
//        sheet.createRow((short) 31);

        Row row32 = sheet.createRow((short) 32);
        Cell cell321 = row26.createCell((short) 0);
        Cell cell322 = row26.createCell((short) 1);

        cell321.setCellValue("3.");
        cell322.setCellValue("PATRIMONIO");

//        sheet.createRow((short) 33);
//        sheet.createRow((short) 34);
//        sheet.createRow((short) 35);
//        sheet.createRow((short) 36);
//        sheet.createRow((short) 37);




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
