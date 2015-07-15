package com.jose.model.estados_financieros.estado_resultado_integral;

import java.util.List;

/**
 * Created by agua on 15/07/15.
 */
public class ResultadoIntegral {
    private List<ElementoRI> mIngresos;
    private List<ElementoRI> mGastos;
    private double mTotalIngresos;
    private double mTotalGastos;
    private double mUtilidad;

    public ResultadoIntegral(List<ElementoRI> ingresos, List<ElementoRI> gastos, double totalIngresos, double totalGastos, double utilidad) {
        mIngresos = ingresos;
        mGastos = gastos;
        mTotalIngresos = totalIngresos;
        mTotalGastos = totalGastos;
        mUtilidad = utilidad;
    }

    public List<ElementoRI> getIngresos() {
        return mIngresos;
    }

    public void setIngresos(List<ElementoRI> ingresos) {
        mIngresos = ingresos;
    }

    public List<ElementoRI> getGastos() {
        return mGastos;
    }

    public void setGastos(List<ElementoRI> gastos) {
        mGastos = gastos;
    }

    public double getTotalIngresos() {
        return mTotalIngresos;
    }

    public void setTotalIngresos(double totalIngresos) {
        mTotalIngresos = totalIngresos;
    }

    public double getTotalGastos() {
        return mTotalGastos;
    }

    public void setTotalGastos(double totalGastos) {
        mTotalGastos = totalGastos;
    }

    public double getUtilidad() {
        return mUtilidad;
    }

    public void setUtilidad(double utilidad) {
        mUtilidad = utilidad;
    }
}
