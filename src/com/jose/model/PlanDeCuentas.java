package com.jose.model;

import java.util.HashMap;
import java.util.Map;
import java.util.TreeMap;

/**
 * Created by agua on 07/07/15.
 */
public class PlanDeCuentas {
    Map<String, String> mCuentas;

    public PlanDeCuentas() {
        mCuentas = new TreeMap<>();
    }

    public void addCuenta(String codigo, String cuenta) {
        mCuentas.put(codigo, cuenta);
    }

    public String getCuenta(String codigo) {
        return mCuentas.get(codigo);
    }

    public void setCuentas(Map<String, String> cuentas) {
        mCuentas = cuentas;
    }

    public Map<String, String> getCuentas() {
        return mCuentas;
    }

    public String getCodigo(String cuenta) {
        String cuentaMayus = cuenta.toUpperCase();
        for (Map.Entry<String, String> entry : mCuentas.entrySet()) {
            if (entry.getValue().equals(cuentaMayus)) {
                return entry.getKey();
            }
        }
        return "";
    }
}
