package com.jose.model.estados_financieros.estado_resultado_integral;

import com.jose.model.estados_financieros.ElementoEF;

import java.util.List;

/**
 * Created by agua on 15/07/15.
 */
public class ResultadoIntegral {
    private List<ElementoEF> mIngresos;
    private List<ElementoEF> mGastos;
    private double mTotalIngresos;
    private double mTotalGastos;
    private double mUtilidad;

    public ResultadoIntegral(List<ElementoEF> ingresos, List<ElementoEF> gastos, double totalIngresos, double totalGastos, double utilidad) {
        mIngresos = ingresos;
        mGastos = gastos;
        mTotalIngresos = totalIngresos;
        mTotalGastos = totalGastos;
        mUtilidad = utilidad;
    }

    public List<ElementoEF> getIngresos() {
        return mIngresos;
    }

    public void setIngresos(List<ElementoEF> ingresos) {
        mIngresos = ingresos;
    }

    public List<ElementoEF> getGastos() {
        return mGastos;
    }

    public void setGastos(List<ElementoEF> gastos) {
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
