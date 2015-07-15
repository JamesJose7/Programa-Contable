package com.jose.model;

import com.jose.model.balance_comprobacion.BalanceDeComprobacion;
import com.jose.model.balance_comprobacion.ElementoBalanceDeComprobacion;
import com.jose.model.estados_financieros.ElementoEF;
import com.jose.model.estados_financieros.estado_resultado_integral.ResultadoIntegral;
import com.jose.model.estados_financieros.estado_situacion_financiera.EstadoSituacionFinanciera;
import com.jose.model.libro_diario.Asiento;
import com.jose.model.libro_diario.LibroDiario;
import com.jose.model.libro_mayor.ElementoMayor;
import com.jose.model.libro_mayor.LibroMayor;
import com.jose.model.libro_mayor.LibrosMayores;
import com.jose.model.plan_de_cuentas.PlanDeCuentas;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.text.DecimalFormat;
import java.util.*;

/**
 * Created by agua on 14/07/15.
 */
public class CicloContable {
    private LibroDiario mLibroDiario = new LibroDiario();
    private PlanDeCuentas mPlanDeCuentas = new PlanDeCuentas();
    private LibrosMayores mLibrosMayores;
    private BalanceDeComprobacion mBalanceDeComprobacion = new BalanceDeComprobacion();
    private ResultadoIntegral mResultadoIntegral;
    private EstadoSituacionFinanciera mEstadoSituacionFinanciera;
    private XSSFWorkbook mWorkbook;

    private DecimalFormat df = new DecimalFormat("#.00");
    private boolean mCerrarCuentas = false;

    public CicloContable(XSSFWorkbook workbook) {
        mWorkbook = workbook;
    }

    public XSSFWorkbook getCicloContable() {
        removeSheets();

        readLibroDiarioSheet();
        readPlanDeCuentasSheet();
        createLibroMayor();
        createBalanceDeComprobacion();
        createResultadoIntegral();
        createEstadoSituacionFinanciera();

        //Generate sheets on workbook
        createBalanceDeComprobacionSheet();

        //Libro mayor con cuentas en cero
        createLibroMayor();
        createLibroMayorSheet();
        createResultadoIntegralSheet();
        createEstadoSituacionFinancieraSheet();

        //printData();

        return mWorkbook;
    }

    public XSSFWorkbook getFirstWorkBook() {
        createLibroDiarioSheet();
        createPlanDeCuentasSheet();

        return mWorkbook;
    }

    public void printData() {
        printLibroDiario();
        printPlanDeCuentas();
        printLibrosMayores();
        printBalanceComprobacion();
    }

    private void removeSheets() {
        try {
            mWorkbook.removeSheetAt(5);
            mWorkbook.removeSheetAt(4);
            mWorkbook.removeSheetAt(3);
            mWorkbook.removeSheetAt(2);
        } catch (Exception e) {
            //e.printStackTrace();
        }
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

    private void createLibroMayor() {
        mLibrosMayores = new LibrosMayores();

        for (Map.Entry<String, String> entry : mPlanDeCuentas.getCuentas().entrySet()) {
            LibroMayor libroMayor;
            List<ElementoMayor> elementosMayoresList = new ArrayList<>();
            double saldo = 0;
            double totalDebe = 0;
            double totalHaber = 0;
            boolean cuentaExists = false;

            String cuenta = entry.getValue();
            ElementoMayor elementoMayor;

            for (Asiento asiento : mLibroDiario.getAsientos()) {

                if (!(asiento.getRegistro().toLowerCase().contains("cierre")) || mCerrarCuentas) {

                    for (Map.Entry<String, Double> entryDebe : asiento.getDebitos().entrySet()) {
                        if (entryDebe.getKey().equalsIgnoreCase(cuenta)) {
                            saldo += entryDebe.getValue();
                            totalDebe += entryDebe.getValue();

                            elementoMayor = new ElementoMayor(asiento.getFecha(), asiento.getRegistro(), asiento.getReferencia(),
                                    entryDebe.getValue(), 0, saldo);

                            elementosMayoresList.add(elementoMayor);

                            cuentaExists = true;
                        }
                    }

                    for (Map.Entry<String, Double> entryHaber : asiento.getCreditos().entrySet()) {
                        if (entryHaber.getKey().equalsIgnoreCase(cuenta)) {
                            saldo -= entryHaber.getValue();

                            totalHaber += entryHaber.getValue();

                            elementoMayor = new ElementoMayor(asiento.getFecha(), asiento.getRegistro(), asiento.getReferencia(),
                                    0, entryHaber.getValue(), saldo);

                            elementosMayoresList.add(elementoMayor);

                            cuentaExists = true;
                        }
                    }

                }
            }

            if (cuentaExists) {
                elementoMayor = new ElementoMayor("TOTAL", "", 0, totalDebe, totalHaber, Math.abs(saldo));
                elementosMayoresList.add(elementoMayor);

                libroMayor = new LibroMayor(cuenta, entry.getKey(), elementosMayoresList);

                mLibrosMayores.addLibroMayor(libroMayor);
            }

        }

        mCerrarCuentas = true;
    }

    private void createBalanceDeComprobacion() {
        int counter = 1;

        List<ElementoBalanceDeComprobacion> elementosList = new ArrayList<>();

        double totalSumasDebe = 0;
        double totalSumasHaber = 0;
        double totalSaldosDebe = 0;
        double totalSaldosHaber = 0;

        for (LibroMayor libroMayor : mLibrosMayores.getLibrosMayoresList()) {
            ElementoBalanceDeComprobacion elementoBalanceDeComprobacion;

            double saldoDebe;
            double saldoHaber;
            ElementoMayor ultimoElemento;

            if (!(libroMayor.getCodigo().charAt(0) == '6')) {

                ultimoElemento = libroMayor.getElementosMayores().get(libroMayor.getElementosMayores().size() - 1);

                if (ultimoElemento.getDebe() > ultimoElemento.getHaber()) {
                    saldoDebe = ultimoElemento.getSaldo();
                    saldoHaber = 0;
                } else {
                    saldoDebe = 0;
                    saldoHaber = ultimoElemento.getSaldo();
                }

                elementoBalanceDeComprobacion = new ElementoBalanceDeComprobacion(counter, libroMayor.getCodigo(), libroMayor.getCuenta(),
                        ultimoElemento.getDebe(), ultimoElemento.getHaber(), saldoDebe, saldoHaber);

                elementosList.add(elementoBalanceDeComprobacion);

                totalSumasDebe += ultimoElemento.getDebe();
                totalSumasHaber += ultimoElemento.getHaber();
                totalSaldosDebe += saldoDebe;
                totalSaldosHaber += saldoHaber;

                counter++;
            }
        }
        mBalanceDeComprobacion.setElementosBalance(elementosList);
        mBalanceDeComprobacion.setTotalSaldosDebe(totalSaldosDebe);
        mBalanceDeComprobacion.setTotalSaldosHaber(totalSaldosHaber);
        mBalanceDeComprobacion.setTotalSumasDebe(totalSumasDebe);
        mBalanceDeComprobacion.setTotalSumasHaber(totalSumasHaber);

    }

    private void createResultadoIntegral() {
        List<ElementoEF> listIngresos = new ArrayList<>();
        List<ElementoEF> listGastos = new ArrayList<>();

        double totalIngresos = 0;
        double totalGastos = 0;
        double utilidad;

        for (ElementoBalanceDeComprobacion elemento : mBalanceDeComprobacion.getElementosBalance()) {
            ElementoEF elementoEF;
            //Add ingresos
            if (elemento.getCodigo().contains("4.1")) {
                elementoEF = new ElementoEF(elemento.getCodigo(), elemento.getCuenta(), elemento.getSaldoHaber());
                listIngresos.add(elementoEF);
                totalIngresos += elemento.getSaldoHaber();
            }

            //Add gastos
            if (elemento.getCodigo().contains("5.1")) {
                elementoEF = new ElementoEF(elemento.getCodigo(), elemento.getCuenta(), elemento.getSaldoDebe());
                listGastos.add(elementoEF);
                totalGastos += elemento.getSaldoDebe();
            }

            //Add costo de venta
            if (elemento.getCodigo().contains("5.") && elemento.getCuenta().toLowerCase().contains("costo")) {
                elementoEF = new ElementoEF(elemento.getCodigo(), elemento.getCuenta(), elemento.getSaldoDebe());
                listIngresos.add(elementoEF);
                totalIngresos -= elemento.getSaldoDebe();
            }
        }

        utilidad = totalIngresos - totalGastos;

        mResultadoIntegral = new ResultadoIntegral(listIngresos, listGastos, totalIngresos, totalGastos, utilidad);


    }

    private void createEstadoSituacionFinanciera() {
        List<ElementoEF> listActivos = new ArrayList<>();
        List<ElementoEF> listPasivos = new ArrayList<>();
        List<ElementoEF> listPatrimonio = new ArrayList<>();

        double totalActivoCorriente = 0;
        double totalActivoNoCorriente = 0;
        double totalPasivoCorriente = 0;
        double totalPasivoNoCorriente = 0;
        double totalPatrimonio = 0;
        double totalPasivoPatrimonio;

        for (ElementoBalanceDeComprobacion elemento : mBalanceDeComprobacion.getElementosBalance()) {
            ElementoEF elementoEF;

            //Add activos
            //Activo corriente
            if (elemento.getCodigo().contains("1.1") && !elemento.getCodigo().contains("5.1")) {
                if (elemento.getSaldoDebe() != 0) {
                    elementoEF = new ElementoEF(elemento.getCodigo(), elemento.getCuenta(), elemento.getSaldoDebe());
                    listActivos.add(elementoEF);
                    totalActivoCorriente += elemento.getSaldoDebe();
                }
            }

            //Activo no corriente
            if (elemento.getCodigo().contains("1.2")) {
                double tempValor;
                if (elemento.getCuenta().toUpperCase().contains("DEPRECIACION")) {
                    tempValor = elemento.getSaldoHaber() * -1;
                } else {
                    tempValor = elemento.getSaldoDebe();
                }

                elementoEF = new ElementoEF(elemento.getCodigo(), elemento.getCuenta(), tempValor);
                listActivos.add(elementoEF);
                totalActivoNoCorriente += tempValor;
            }

            //Add pasivos
            //Pasivo Corriente
            if (elemento.getCodigo().contains("2.1") && elemento.getSaldoHaber() != 0) {
                elementoEF = new ElementoEF(elemento.getCodigo(), elemento.getCuenta(), elemento.getSaldoHaber());
                listPasivos.add(elementoEF);
                totalPasivoCorriente += elemento.getSaldoHaber();
            }

            //Pasivo No corriente
            if (elemento.getCodigo().contains("2.2") && elemento.getSaldoHaber() != 0) {
                elementoEF = new ElementoEF(elemento.getCodigo(), elemento.getCuenta(), elemento.getSaldoHaber());
                listPasivos.add(elementoEF);
                totalPasivoNoCorriente += elemento.getSaldoHaber();
            }

            //Add Patrimonio
            if (elemento.getCodigo().contains("3.1")) {
                elementoEF = new ElementoEF(elemento.getCodigo(), elemento.getCuenta(), elemento.getSaldoHaber());
                listPatrimonio.add(elementoEF);
                totalPatrimonio += elemento.getSaldoHaber();
            }
        }

        totalPasivoPatrimonio = totalPasivoCorriente + totalPasivoNoCorriente + totalPatrimonio;

        mEstadoSituacionFinanciera = new EstadoSituacionFinanciera(listActivos, listPasivos, listPatrimonio,
                totalActivoCorriente, totalActivoNoCorriente, totalPasivoCorriente, totalPasivoNoCorriente,
                totalPatrimonio, totalPasivoPatrimonio);

    }

    private void createEstadoSituacionFinancieraSheet() {
        //Create sheet
        XSSFSheet sheet = mWorkbook.createSheet("Estado de Situacion Financiera");

        short rowCounter = 0;

        //Headers
        Row row = sheet.createRow(rowCounter);
        Cell cell = row.createCell((short) 0);
        createCell(mWorkbook, row, (short) 0, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER);
        sheet.addMergedRegion(new CellRangeAddress(
                rowCounter, //first row (0-based)
                rowCounter, //last row  (0-based)
                0, //first column (0-based)
                3  //last column  (0-based)
        ));
        cell.setCellValue("Empresa \"X\"");
        rowCounter++;

        Row row1 = sheet.createRow(rowCounter);
        Cell cell1 = row1.createCell((short) 0);
        createCell(mWorkbook, row1, (short) 0, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER);
        sheet.addMergedRegion(new CellRangeAddress(
                rowCounter, //first row (0-based)
                rowCounter, //last row  (0-based)
                0, //first column (0-based)
                3  //last column  (0-based)
        ));
        cell1.setCellValue("Estado de Resultado Integral");
        rowCounter++;

        Row row2 = sheet.createRow(rowCounter);
        Cell cell2 = row2.createCell((short) 0);
        createCell(mWorkbook, row2, (short) 0, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER);
        sheet.addMergedRegion(new CellRangeAddress(
                rowCounter, //first row (0-based)
                rowCounter, //last row  (0-based)
                0, //first column (0-based)
                3  //last column  (0-based)
        ));
        cell2.setCellValue("01 - Junio a 30 Junio de 2015");
        rowCounter++;

        Row row3 = sheet.createRow(rowCounter);
        row3.createCell(0).setCellValue("1.");
        row3.createCell(1).setCellValue("Activo");
        rowCounter++;

        Row row4 = sheet.createRow(rowCounter);
        row4.createCell(0).setCellValue("1.1");
        row4.createCell(1).setCellValue("Activo Corriente");
        rowCounter++;

        for (ElementoEF elementoEF : mEstadoSituacionFinanciera.getActivos()) {
            if (elementoEF.getCodigo().contains("1.1")) {
                Row rowA = sheet.createRow(rowCounter);
                rowA.createCell(0).setCellValue(elementoEF.getCodigo());
                rowA.createCell(1).setCellValue(elementoEF.getCuenta());
                rowA.createCell(2).setCellValue(formatDecimal(elementoEF.getValor()));
                rowCounter++;
            }
        }

        Row row5 = sheet.createRow(rowCounter);
        row5.createCell(1).setCellValue("Total activo corriente");
        row5.createCell(3).setCellValue(formatDecimal(mEstadoSituacionFinanciera.getTotalActivoCorriente()));
        rowCounter += 2;

        Row row6 = sheet.createRow(rowCounter);
        row6.createCell(0).setCellValue("1.2.");
        row6.createCell(1).setCellValue("Activo no Corriente");
        rowCounter++;

        for (ElementoEF elementoEF : mEstadoSituacionFinanciera.getActivos()) {
            if (elementoEF.getCodigo().contains("1.2")) {
                Row rowA = sheet.createRow(rowCounter);
                rowA.createCell(0).setCellValue(elementoEF.getCodigo());
                rowA.createCell(1).setCellValue(elementoEF.getCuenta());
                rowA.createCell(2).setCellValue(formatDecimal(elementoEF.getValor()));
                rowCounter++;
            }
        }

        Row row7 = sheet.createRow(rowCounter);
        row7.createCell(1).setCellValue("Total activo no corriente");
        row7.createCell(3).setCellValue(formatDecimal(mEstadoSituacionFinanciera.getTotalActivoNoCorriente()));
        rowCounter += 2;

        Row row8 = sheet.createRow(rowCounter);
        row8.createCell(1).setCellValue("TOTAL ACTIVO");
        row8.createCell(3).setCellValue(formatDecimal(mEstadoSituacionFinanciera.getTotalActivoCorriente() + mEstadoSituacionFinanciera.getTotalActivoNoCorriente()));
        rowCounter += 2;

        Row row9 = sheet.createRow(rowCounter);
        row9.createCell(0).setCellValue("2.");
        row9.createCell(1).setCellValue("Pasivos");
        rowCounter++;

        Row row10 = sheet.createRow(rowCounter);
        row10.createCell(0).setCellValue("2.1");
        row10.createCell(1).setCellValue("Pasivos Corriente");
        rowCounter++;

        for (ElementoEF elementoEF : mEstadoSituacionFinanciera.getPasivos()) {
            if (elementoEF.getCodigo().contains("2.1")) {
                Row rowP = sheet.createRow(rowCounter);
                rowP.createCell(0).setCellValue(elementoEF.getCodigo());
                rowP.createCell(1).setCellValue(elementoEF.getCuenta());
                rowP.createCell(2).setCellValue(formatDecimal(elementoEF.getValor()));
                rowCounter++;
            }
        }

        Row row11 = sheet.createRow(rowCounter);
        row11.createCell(1).setCellValue("Total pasivo corriente");
        row11.createCell(3).setCellValue(formatDecimal(mEstadoSituacionFinanciera.getTotalPasivoCorriente()));
        rowCounter += 2;

        Row row12 = sheet.createRow(rowCounter);
        row12.createCell(0).setCellValue("2.2");
        row12.createCell(1).setCellValue("Pasivo no Corriente");
        rowCounter++;

        for (ElementoEF elementoEF : mEstadoSituacionFinanciera.getPasivos()) {
            if (elementoEF.getCodigo().contains("2.2")) {
                Row rowP = sheet.createRow(rowCounter);
                rowP.createCell(0).setCellValue(elementoEF.getCodigo());
                rowP.createCell(1).setCellValue(elementoEF.getCuenta());
                rowP.createCell(2).setCellValue(formatDecimal(elementoEF.getValor()));
                rowCounter++;
            }
        }

        Row row13 = sheet.createRow(rowCounter);
        row13.createCell(1).setCellValue("Total pasivo no corriente");
        row13.createCell(3).setCellValue(formatDecimal(mEstadoSituacionFinanciera.getTotalPasivoNoCorriente()));
        rowCounter += 2;

        Row row14 = sheet.createRow(rowCounter);
        row14.createCell(1).setCellValue("TOTAL PASIVO");
        row14.createCell(3).setCellValue(formatDecimal(mEstadoSituacionFinanciera.getTotalPasivoCorriente() + mEstadoSituacionFinanciera.getTotalPasivoNoCorriente()));
        rowCounter += 2;

        Row row15 = sheet.createRow(rowCounter);
        row15.createCell(0).setCellValue("3.");
        row15.createCell(1).setCellValue("Patrimonio");
        rowCounter++;

        for (ElementoEF elementoEF : mEstadoSituacionFinanciera.getPatrimonio()) {
            Row rowPT = sheet.createRow(rowCounter);
            rowPT.createCell(0).setCellValue(elementoEF.getCodigo());
            rowPT.createCell(1).setCellValue(elementoEF.getCuenta());
            rowPT.createCell(2).setCellValue(formatDecimal(elementoEF.getValor()));
            rowCounter++;
        }

        Row row16 = sheet.createRow(rowCounter);
        row16.createCell(1).setCellValue("Total patrimonio");
        row16.createCell(3).setCellValue(formatDecimal(mEstadoSituacionFinanciera.getTotalPatrimonio()));
        rowCounter += 2;

        Row row17 = sheet.createRow(rowCounter);
        row17.createCell(1).setCellValue("TOTAL PASIVO Y PATRIMONIO");
        row17.createCell(3).setCellValue(formatDecimal(mEstadoSituacionFinanciera.getTotalPasivoPatrimonio()));
    }

    private void createResultadoIntegralSheet() {
        //Create sheet
        XSSFSheet sheet = mWorkbook.createSheet("Resultado Integral");

        short rowCounter = 0;

        //Headers
        Row row = sheet.createRow(rowCounter);
        Cell cell = row.createCell((short) 0);
        createCell(mWorkbook, row, (short) 0, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER);
        sheet.addMergedRegion(new CellRangeAddress(
                rowCounter, //first row (0-based)
                rowCounter, //last row  (0-based)
                0, //first column (0-based)
                3  //last column  (0-based)
        ));
        cell.setCellValue("Empresa \"X\"");
        rowCounter++;

        Row row1 = sheet.createRow(rowCounter);
        Cell cell1 = row1.createCell((short) 0);
        createCell(mWorkbook, row1, (short) 0, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER);
        sheet.addMergedRegion(new CellRangeAddress(
                rowCounter, //first row (0-based)
                rowCounter, //last row  (0-based)
                0, //first column (0-based)
                3  //last column  (0-based)
        ));
        cell1.setCellValue("Estado de Resultado Integral");
        rowCounter++;

        Row row2 = sheet.createRow(rowCounter);
        Cell cell2 = row2.createCell((short) 0);
        createCell(mWorkbook, row2, (short) 0, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER);
        sheet.addMergedRegion(new CellRangeAddress(
                rowCounter, //first row (0-based)
                rowCounter, //last row  (0-based)
                0, //first column (0-based)
                3  //last column  (0-based)
        ));
        cell2.setCellValue("01 - Junio a 30 Junio de 2015");
        rowCounter++;

        Row row3 = sheet.createRow(rowCounter);
        row3.createCell(0).setCellValue("4.");
        row3.createCell(1).setCellValue("Ingresos");
        rowCounter++;

        for (ElementoEF elementoEF : mResultadoIntegral.getIngresos()) {
            Row rowI = sheet.createRow(rowCounter);
            rowI.createCell(0).setCellValue(elementoEF.getCodigo());
            rowI.createCell(1).setCellValue(elementoEF.getCuenta());
            rowI.createCell(2).setCellValue(elementoEF.getValor());
            rowCounter++;
        }

        Row row4 = sheet.createRow(rowCounter);
        row4.createCell(1).setCellValue("-Utilidad de venta");
        row4.createCell(3).setCellValue(mResultadoIntegral.getTotalIngresos());
        rowCounter += 2;

        Row row5 = sheet.createRow(rowCounter);
        row5.createCell(0).setCellValue("5.");
        row5.createCell(1).setCellValue("Gastos");
        rowCounter++;

        Row row6 = sheet.createRow(rowCounter);
        row6.createCell(0).setCellValue("5.1.");
        row6.createCell(1).setCellValue("Gastos Operativos");
        rowCounter++;

        for (ElementoEF elementoEF : mResultadoIntegral.getGastos()) {
            Row rowG = sheet.createRow(rowCounter);
            rowG.createCell(0).setCellValue(elementoEF.getCodigo());
            rowG.createCell(1).setCellValue(elementoEF.getCuenta());
            rowG.createCell(2).setCellValue(elementoEF.getValor());
            rowCounter++;
        }

        Row row7 = sheet.createRow(rowCounter);
        row7.createCell(1).setCellValue("Total Gastos");
        row7.createCell(3).setCellValue(mResultadoIntegral.getTotalGastos());
        rowCounter++;

        Row row8 = sheet.createRow(rowCounter);
        row8.createCell(1).setCellValue("-Utilidad antes de participacion");
        row8.createCell(3).setCellValue(mResultadoIntegral.getUtilidad());
    }

    private void createLibroMayorSheet() {

        //Create sheet for "Libro Mayor"
        XSSFSheet sheet = mWorkbook.createSheet("Libro Mayor");

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
                row4.createCell((short) 3).setCellValue(formatDecimal(elementoMayor.getDebe()));
                row4.createCell((short) 4).setCellValue(formatDecimal(elementoMayor.getHaber()));
                row4.createCell((short) 5).setCellValue(formatDecimal(elementoMayor.getSaldo()));

                rowCounter++;
            }

            rowCounter += 3;

        }
    }

    private void createBalanceDeComprobacionSheet() {
        //Create sheet for "Libro Mayor"
        XSSFSheet sheet = mWorkbook.createSheet("Balance de Comprobacion");

        int rowCounter = 0;

        Row row = sheet.createRow(rowCounter);
        Cell cell = row.createCell((short) 0);
        createCell(mWorkbook, row, (short) 0, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER);
        sheet.addMergedRegion(new CellRangeAddress(
                rowCounter, //first row (0-based)
                (rowCounter + 1), //last row  (0-based)
                0, //first column (0-based)
                7  //last column  (0-based)
        ));
        cell.setCellValue("Balance de comprobacion");
        rowCounter += 2;

        Row row1 = sheet.createRow(rowCounter);
        row1.createCell((short) 0).setCellValue("No");
        row1.createCell((short) 1).setCellValue("Codigo");
        row1.createCell((short) 2).setCellValue("Cuenta");
        row1.createCell((short) 3).setCellValue("Suma Debe");
        row1.createCell((short) 4).setCellValue("Suma Haber");
        row1.createCell((short) 5).setCellValue("Saldo Debe");
        row1.createCell((short) 6).setCellValue("Saldo Haber");
        rowCounter++;

        for (ElementoBalanceDeComprobacion elemento : mBalanceDeComprobacion.getElementosBalance()) {
            Row rowE = sheet.createRow(rowCounter);
            rowE.createCell((short) 0).setCellValue(elemento.getNumero());
            rowE.createCell((short) 1).setCellValue(elemento.getCodigo());
            rowE.createCell((short) 2).setCellValue(elemento.getCuenta());
            rowE.createCell((short) 3).setCellValue(formatDecimal(elemento.getSumaDebe()));
            rowE.createCell((short) 4).setCellValue(formatDecimal(elemento.getSumaHaber()));
            rowE.createCell((short) 5).setCellValue(formatDecimal(elemento.getSaldoDebe()));
            rowE.createCell((short) 6).setCellValue(formatDecimal(elemento.getSaldoHaber()));
            rowCounter++;
        }

        Row lastRow = sheet.createRow(rowCounter);
        lastRow.createCell((short) 2).setCellValue("TOTALES");
        lastRow.createCell((short) 3).setCellValue(formatDecimal(mBalanceDeComprobacion.getTotalSumasDebe()));
        lastRow.createCell((short) 4).setCellValue(formatDecimal(mBalanceDeComprobacion.getTotalSumasHaber()));
        lastRow.createCell((short) 5).setCellValue(formatDecimal(mBalanceDeComprobacion.getTotalSaldosDebe()));
        lastRow.createCell((short) 6).setCellValue(formatDecimal(mBalanceDeComprobacion.getTotalSaldosHaber()));


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
        createCell(mWorkbook, row2, (short) 1, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER);

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

        sheet.addMergedRegion(new CellRangeAddress(
                1, //first row (0-based)
                1, //last row  (0-based)
                1, //first column (0-based)
                2  //last column  (0-based)
        ));

    }

    private void createPlanDeCuentasSheet() {
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

    private void printLibroDiario() {
        List<Asiento> asientoList = mLibroDiario.getAsientos();
        for (Asiento asiento : asientoList) {
            printAsiento(asiento);
            System.out.println("\n");
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

    private void printBalanceComprobacion() {
        System.out.printf("| %2s | %8s | %-50s | %10s | %10s | %10s | %10s |\n",
                "No", "Codigo", "Cuentas", "Debe", "Haber", "Debe", "Haber");
        for (ElementoBalanceDeComprobacion elementoBalanceDeComprobacion : mBalanceDeComprobacion.getElementosBalance()) {
            System.out.printf("| %2d | %8s | %-50s | %10.2f | %10.2f | %10.2f | %10.2f |\n",
                    elementoBalanceDeComprobacion.getNumero(),
                    elementoBalanceDeComprobacion.getCodigo(),
                    elementoBalanceDeComprobacion.getCuenta(),
                    elementoBalanceDeComprobacion.getSumaDebe(),
                    elementoBalanceDeComprobacion.getSumaHaber(),
                    elementoBalanceDeComprobacion.getSaldoDebe(),
                    elementoBalanceDeComprobacion.getSaldoHaber());
        }

        System.out.printf("| %2s | %8s | %-50s | %10.2f | %10.2f | %10.2f | %10.2f |\n",
                "", "", "TOTAL",
                mBalanceDeComprobacion.getTotalSumasDebe(),
                mBalanceDeComprobacion.getTotalSumasHaber(),
                mBalanceDeComprobacion.getTotalSaldosDebe(),
                mBalanceDeComprobacion.getTotalSaldosHaber());
    }


    private void printAsiento(Asiento asiento) {
        System.out.println("Fecha: " + asiento.getFecha());
        //Cuenta debitos
        System.out.println("Cuenta debe:");
        for (Map.Entry<String, Double> entry : asiento.getDebitos().entrySet()) {
            System.out.printf("\tcuenta: %s\n" +
                            "\tvalor: %.2f\n",
                    entry.getKey(),
                    entry.getValue());
        }
        //Cuenta haber
        System.out.println("Cuenta haber:");
        for (Map.Entry<String, Double> entry : asiento.getCreditos().entrySet()) {
            System.out.printf("\tcuenta: %s\n" +
                            "\tvalor: %.2f\n",
                    entry.getKey(),
                    entry.getValue());
        }
        System.out.println("Registro: " + asiento.getRegistro());

    }

    private String formatDecimal(double d) {
        if (d != 0) {
            return String.format("%.2f", d);
        }
        return "0";
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
