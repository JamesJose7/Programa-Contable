package com.jose.model.balance_comprobacion;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by agua on 09/07/15.
 */
public class BalanceDeComprobacion {
    private List<ElementoBalanceDeComprobacion> mElementosBalance;
    private double mTotalSumas;
    private double mTotalSaldos;

    public BalanceDeComprobacion(List<ElementoBalanceDeComprobacion> elementosBalance, double totalSumas, double totalSaldos) {
        mElementosBalance = new ArrayList<>(elementosBalance);
        mTotalSumas = totalSumas;
        mTotalSaldos = totalSaldos;
    }

    public List<ElementoBalanceDeComprobacion> getElementosBalance() {
        return mElementosBalance;
    }

    public void setElementosBalance(List<ElementoBalanceDeComprobacion> elementosBalance) {
        mElementosBalance = elementosBalance;
    }

    public double getTotalSumas() {
        return mTotalSumas;
    }

    public void setTotalSumas(double totalSumas) {
        mTotalSumas = totalSumas;
    }

    public double getTotalSaldos() {
        return mTotalSaldos;
    }

    public void setTotalSaldos(double totalSaldos) {
        mTotalSaldos = totalSaldos;
    }
}
