package com.jose.model.file_generators;

import com.jose.model.libro_diario.Asiento;
import com.jose.model.libro_diario.LibroDiario;
import com.jose.model.PlanDeCuentas;
import com.jose.model.libro_mayor.ElementoMayor;
import com.jose.model.libro_mayor.LibroMayor;
import com.jose.model.libro_mayor.LibrosMayores;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.*;
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
    PlanDeCuentas mPlanDeCuentas = new PlanDeCuentas();
    LibrosMayores mLibrosMayores = new LibrosMayores();
    File mLibroDiarioFile;
    XSSFWorkbook mWorkbook;


    public void openFile() {
        mLibroDiarioFile = new File("Contabilidad.xlsx");
        //Create a new workbook
        try {
            if (!mLibroDiarioFile.exists()) {
                mWorkbook = new XSSFWorkbook();
                List<XSSFSheet> sheets = new ArrayList<>();

                createLibroDiarioSheet();
                createPlanDeCuentas();
            }
            FileInputStream file = new FileInputStream(mLibroDiarioFile);

            //Create a workbook from the file
            mWorkbook = new XSSFWorkbook(file);


            readLibroDiarioSheet();
            readPlanDeCuentasSheet();

            createLibroMayor();
            createLibroMayorSheet();

            printLibroDiario();
            printPlanDeCuentas();
            printLibrosMayores();

            createFile();

            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }


    }

    private void createFile() throws Exception{
        //create a new file
        FileOutputStream fos = new FileOutputStream(mLibroDiarioFile);
        //write workBook
        mWorkbook.write(fos);
        fos.close();
        System.out.println("File created");
    }

    private void readPlanDeCuentasSheet() {
        XSSFSheet sheet = mWorkbook.getSheetAt(1);


        //Iterate through rows
        Iterator<Row> rowIterator = sheet.iterator();

        while (rowIterator.hasNext()) {

            Row row = rowIterator.next();

            //Iterate through cells
            Iterator<Cell> cellIterator = row.cellIterator();

            //Temporal variables for each row
            String codigo = "";
            String cuenta = "";

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();

                //Check cell types
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_NUMERIC:
                        if (row.getRowNum() > 1) {
                            //Store values

                        }
                        break;
                    case Cell.CELL_TYPE_STRING:
                        if (row.getRowNum() > 4) {
                            if (!cell.getStringCellValue().equals("Estado de Resultado Integral")) {
                                //Store values
                                switch (cell.getColumnIndex()) {
                                    case 0:
                                        codigo = cell.getStringCellValue();
                                        break;
                                    case 1:
                                        cuenta = cell.getStringCellValue().toUpperCase().trim();
                                        break;
                                }
                            }
                        }
                        break;
                    default:
                        break;
                }
            }
            mPlanDeCuentas.addCuenta(codigo, cuenta);
        }
        mPlanDeCuentas.getCuentas().remove("");
    }

    private void readLibroDiarioSheet() {

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
        int referencia = 1;

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

                mLibroDiario.addAsiento(new Asiento(fecha, mapDebe, mapHaber, registro, referencia));
                fecha = "";
                referencia++;

                mapDebe = new HashMap<>();
                mapHaber = new HashMap<>();

                isFinishedAsiento = false;
            }

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

        Row row12 = sheet.createRow((short) 12);
        Cell cell121 = row12.createCell((short) 0);
        Cell cell122 = row12.createCell((short) 1);

        cell121.setCellValue("1.2");
        cell122.setCellValue("ACTIVO NO CORRIENTE");

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

        Row row26 = sheet.createRow((short) 26);
        Cell cell261 = row26.createCell((short) 0);
        Cell cell262 = row26.createCell((short) 1);

        cell261.setCellValue("2.2");
        cell262.setCellValue("PASIVO NO CORRIENTE");

        Row row32 = sheet.createRow((short) 32);
        Cell cell321 = row32.createCell((short) 0);
        Cell cell322 = row32.createCell((short) 1);

        cell321.setCellValue("3.");
        cell322.setCellValue("PATRIMONIO");

        Row row34 = sheet.createRow((short) 34);
        Cell cell341 = row34.createCell((short) 0);
        createCell(mWorkbook, row34, (short) 0, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER);
        cell341.setCellValue("Estado de Resultado Integral");

        sheet.addMergedRegion(new CellRangeAddress(
                34, //first row (0-based)
                34, //last row  (0-based)
                0, //first column (0-based)
                4  //last column  (0-based)
        ));

        Row row36 = sheet.createRow((short) 36);
        Cell cell361 = row36.createCell((short) 0);
        Cell cell362 = row36.createCell((short) 1);
        Row row37 = sheet.createRow((short) 37);
        Cell cell371 = row37.createCell((short) 0);
        Cell cell372 = row37.createCell((short) 1);

        cell361.setCellValue("4.");
        cell362.setCellValue("INGRESOS");
        cell371.setCellValue("4.1");
        cell372.setCellValue("INGRESOS OPERACIONALES");

        Row row42 = sheet.createRow((short) 42);
        Cell cell421 = row42.createCell((short) 0);
        Cell cell422 = row42.createCell((short) 1);

        cell421.setCellValue("4.2");
        cell422.setCellValue("INGRESOS NO OPERACIONALES");

        Row row47 = sheet.createRow((short) 47);
        Cell cell471 = row47.createCell((short) 0);
        Cell cell472 = row47.createCell((short) 1);
        Row row48 = sheet.createRow((short) 48);
        Cell cell481 = row48.createCell((short) 0);
        Cell cell482 = row48.createCell((short) 1);

        cell471.setCellValue("5.");
        cell472.setCellValue("GASTOS");
        cell481.setCellValue("5.1");
        cell482.setCellValue("GASTOS OPERACIONALES");

        Row row53 = sheet.createRow((short) 53);
        Cell cell531 = row53.createCell((short) 0);
        Cell cell532 = row53.createCell((short) 1);

        cell531.setCellValue("5.2");
        cell532.setCellValue("GASTOS NO OPERACIONALES");

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

    private void createLibroMayorSheet() {
        //Create sheet for "Libro Mayor"
        XSSFSheet sheet;
        try {
            sheet = mWorkbook.createSheet("Libro Mayor");
        } catch (Exception e) {
            sheet = mWorkbook.getSheetAt(2);
        }

        //Data to be writen

        //Rows counter to keep printing downwards
        short rowCounter = 0;

        for (LibroMayor libroMayor : mLibrosMayores.getLibrosMayoresList()) {
            //Headers
            Row row = sheet.createRow(rowCounter);
            Cell cell = row.createCell((short) 0);
            createCell(mWorkbook, row, (short) 0, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER);
            sheet.addMergedRegion(new CellRangeAddress(
                    rowCounter, //first row (0-based)
                    rowCounter, //last row  (0-based)
                    0, //first column (0-based)
                    5  //last column  (0-based)
            ));
            cell.setCellValue("Libro Mayor");
            rowCounter++;

            Row row1 = sheet.createRow(rowCounter);
            Cell cell1 = row1.createCell((short) 0);
            sheet.addMergedRegion(new CellRangeAddress(
                    rowCounter, //first row (0-based)
                    rowCounter, //last row  (0-based)
                    0, //first column (0-based)
                    5  //last column  (0-based)
            ));
            cell1.setCellValue("Cuenta: " + libroMayor.getCuenta().toUpperCase());
            rowCounter++;

            Row row2 = sheet.createRow(rowCounter);
            Cell cell2 = row2.createCell((short) 0);
            sheet.addMergedRegion(new CellRangeAddress(
                    rowCounter, //first row (0-based)
                    rowCounter, //last row  (0-based)
                    0, //first column (0-based)
                    5  //last column  (0-based)
            ));
            cell2.setCellValue("Codigo: " + libroMayor.getCodigo());
            rowCounter++;

            Row row3 = sheet.createRow(rowCounter);
            row3.createCell((short) 0).setCellValue("FECHA");
            row3.createCell((short) 1).setCellValue("DETALLE");
            row3.createCell((short) 2).setCellValue("REF.");
            row3.createCell((short) 3).setCellValue("DEBE");
            row3.createCell((short) 4).setCellValue("HABER");
            row3.createCell((short) 5).setCellValue("SALDO");
            rowCounter++;

            for (ElementoMayor elementoMayor : libroMayor.getElementosMayores()) {
                Row row4 = sheet.createRow(rowCounter);
                row4.createCell((short) 0).setCellValue(elementoMayor.getFecha());
                row4.createCell((short) 1).setCellValue(elementoMayor.getDetalle());
                row4.createCell((short) 2).setCellValue(elementoMayor.getReferencia());
                row4.createCell((short) 3).setCellValue(elementoMayor.getDebe());
                row4.createCell((short) 4).setCellValue(elementoMayor.getHaber());
                row4.createCell((short) 5).setCellValue(elementoMayor.getSaldo());

                rowCounter++;
            }

            rowCounter += 3;

        }
    }

    private void createLibroMayor() {
        for (Map.Entry<String, String> entry : mPlanDeCuentas.getCuentas().entrySet()) {
            LibroMayor libroMayor;
            List<ElementoMayor> elementosMayoresList = new ArrayList<>();
            double saldo = 0;
            boolean cuentaExists = false;

            String cuenta = entry.getValue();

            for (Asiento asiento : mLibroDiario.getAsientos()) {
                ElementoMayor elementoMayor;

                for (Map.Entry<String, Double> entryDebe : asiento.getDebitos().entrySet()) {
                    if (entryDebe.getKey().equalsIgnoreCase(cuenta)) {
                        saldo += entryDebe.getValue();

                        elementoMayor = new ElementoMayor(asiento.getFecha(), asiento.getRegistro(), asiento.getReferencia(),
                                entryDebe.getValue(), 0, saldo);

                        elementosMayoresList.add(elementoMayor);

                        cuentaExists = true;
                    }
                }

                for (Map.Entry<String, Double> entryHaber : asiento.getCreditos().entrySet()) {
                    if (entryHaber.getKey().equalsIgnoreCase(cuenta)) {
                        saldo -= entryHaber.getValue();

                        elementoMayor = new ElementoMayor(asiento.getFecha(), asiento.getRegistro(), asiento.getReferencia(),
                                0, entryHaber.getValue(), saldo);

                        elementosMayoresList.add(elementoMayor);

                        cuentaExists = true;
                    }
                }
            }

            if (cuentaExists) {
                libroMayor = new LibroMayor(cuenta, entry.getKey(), elementosMayoresList);

                mLibrosMayores.addLibroMayor(libroMayor);
            }

        }
    }

    private void printLibrosMayores() {
        for (LibroMayor libroMayor : mLibrosMayores.getLibrosMayoresList()) {
            System.out.printf("\nCuenta: %s\n" +
                    "Codigo: %s\n", libroMayor.getCuenta(), libroMayor.getCodigo());
            for (ElementoMayor elementoMayor : libroMayor.getElementosMayores()) {
                System.out.printf("-----------------------------\n" +
                                "Fecha: %s\n" +
                                "Detalle: %s\n" +
                                "Referencia: %d\n" +
                                "Debe: %.2f\n" +
                                "Haber: %.2f\n" +
                                "Saldo: %.2f\n",
                        elementoMayor.getFecha(),
                        elementoMayor.getDetalle(),
                        elementoMayor.getReferencia(),
                        elementoMayor.getDebe(),
                        elementoMayor.getHaber(),
                        elementoMayor.getSaldo());
            }
            System.out.println("-----------------TERMINADO----------------------------");
        }
    }

    private void printPlanDeCuentas() {
        System.out.println("\n\n\n\nCUENTAS\n");
        for (Map.Entry<String, String> entry : mPlanDeCuentas.getCuentas().entrySet()) {
            System.out.printf("%8s | %s\n",
                    entry.getKey(),
                    entry.getValue());

        }
    }

    private void printLibroDiario() {
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
