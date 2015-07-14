package com.jose.model;


import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * Created by agua on 30/06/15.
 */
public class FileWorkbook {
    public final int FILE_CREATED = 0;
    public final int FILE_UPDATED = 1;
    public final int FILE_ERROR = 2;


    File mLibroDiarioFile = new File("Excel/Contabilidad.xlsx");
    XSSFWorkbook mWorkbook;
    CicloContable mCicloContable;


    public int saveWorkBook() {
        try {
            //Create a new workbook
            if (!mLibroDiarioFile.exists()) {
                mWorkbook = new XSSFWorkbook();
                mCicloContable = new CicloContable(mWorkbook);

                mWorkbook = mCicloContable.getFirstWorkBook();

                writeToFile();

                return FILE_CREATED;
            } else {
                //Open the workbook
                FileInputStream file = new FileInputStream(mLibroDiarioFile);

                //Create a workbook from the file
                mWorkbook = new XSSFWorkbook(file);

                mCicloContable = new CicloContable(mWorkbook);

                mWorkbook = mCicloContable.getCicloContable();

                writeToFile();

                return FILE_UPDATED;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return FILE_ERROR;
    }

    private void writeToFile() throws Exception{
        //create a new file
        FileOutputStream fos = new FileOutputStream(mLibroDiarioFile);
        //write workBook
        mWorkbook.write(fos);
        fos.close();
        System.out.println("File written");
    }




}
