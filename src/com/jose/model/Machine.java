package com.jose.model;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by agua on 09/07/15.
 */
public class Machine {
    /*private FileWorkbook mFileWorkbook = new FileWorkbook();
    private BufferedReader mReader;
    private Map<String, String> mMenu;

    public Machine() {
        mReader = new BufferedReader(new InputStreamReader(System.in));
        mMenu = new HashMap<>();
        mMenu.put("Crear", "Crea un nuevo nuevo archivo Excel");
        mMenu.put("Abrir", "Lee el archivo modificado");
        mMenu.put("Mayorizar", "Genera el libro mayor de todas las cuentas");
        mMenu.put("Balance", "Genera el balance de comprobacion");
        mMenu.put("Salir", "Sale del programa");
    }

    private String promptAction() throws IOException {
        System.out.printf("\nMenu: \n\n");

        for (Map.Entry<String, String> option : mMenu.entrySet()) {
            System.out.printf("%s - %s \n",
                    option.getKey(),
                    option.getValue());

        }
        System.out.print("Que desea hacer: ");
        String choice = mReader.readLine();
        return choice.trim().toLowerCase();
    }

    public void run() {
        String choice = "";
        do {
            try {
                choice = promptAction();
                switch (choice) {
                    case "crear":
                        mFileWorkbook.createWorkBook();
                        break;
                    case "abrir":
                        mFileWorkbook.openFile();
                        break;
                    case "mayorizar":
                        mFileWorkbook.generateLibroMayor();
                        break;
                    case "balance":
                        mFileWorkbook.generateBalanceComprobacion();
                        break;
                    case "salir":
                        System.out.println(" - Programa terminado - ");
                        break;
                    default:
                        System.out.printf("%s - Opcion no valida. Intente de nuevo por favor",
                                choice);
                }
            } catch (Exception e) {
                System.out.println("Input problem");
                e.printStackTrace();
            }
        } while(!choice.equalsIgnoreCase("salir"));
    }*/

}
