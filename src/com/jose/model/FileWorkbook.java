package com.jose.model;

import com.jose.model.balance_comprobacion.BalanceDeComprobacion;
import com.jose.model.balance_comprobacion.ElementoBalanceDeComprobacion;
import com.jose.model.libro_diario.Asiento;
import com.jose.model.libro_diario.LibroDiario;
import com.jose.model.plan_de_cuentas.PlanDeCuentas;
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
    public final int FILE_CREATED = 0;
    public final int FILE_UPDATED = 1;
    public final int FILE_ERROR = 2;


    File mLibroDiarioFile = new File("Contabilidad.xlsx");
    XSSFWorkbook mWorkbook;
    CicloContable mCicloContable;


    public int saveWorkBook() {
        try {
            //Create a new workbook
            if (!mLibroDiarioFile.exists()) {
                mWorkbook = new XSSFWorkbook();
                mCicloContable = new CicloContable(mWorkbook);

                mWorkbook = mCicloContable.getFirstWorkBook();

                createFile();

                return FILE_CREATED;
            } else {
                //Open the workbook
                FileInputStream file = new FileInputStream(mLibroDiarioFile);

                //Create a workbook from the file
                mWorkbook = new XSSFWorkbook(file);

                mCicloContable = new CicloContable(mWorkbook);

                mWorkbook = mCicloContable.getCicloContable();

                createFile();

                return FILE_UPDATED;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return FILE_ERROR;
    }

    private void createFile() throws Exception{
        //create a new file
        FileOutputStream fos = new FileOutputStream(mLibroDiarioFile);
        //write workBook
        mWorkbook.write(fos);
        fos.close();
        System.out.println("File created");
    }




}
