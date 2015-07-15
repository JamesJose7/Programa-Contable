package com.jose.model.estados_financieros.estado_situacion_financiera;

import com.jose.model.estados_financieros.ElementoEF;

import java.util.List;

/**
 * Created by agua on 15/07/15.
 */
public class EstadoSituacionFinanciera {
    private List<ElementoEF> mActivos;
    private List<ElementoEF> mPasivos;
    private List<ElementoEF> mPatrimonio;
    private double mTotalActivoCorriente;
    private double mTotalActivoNoCorriente;
    private double mTotalPasivoCorriente;
    private double mTotalPasivoNoCorriente;
    private double mTotalPatrimonio;
    private double mTotalPasivoPatrimonio;

    public EstadoSituacionFinanciera(List<ElementoEF> activos, List<ElementoEF> pasivos, List<ElementoEF> patrimonio, double totalActivoCorriente,
                                     double totalActivoNoCorriente, double totalPasivoCorriente, double totalPasivoNoCorriente, double totalPatrimonio, double totalPasivoPatrimonio) {
        mActivos = activos;
        mPasivos = pasivos;
        mPatrimonio = patrimonio;
        mTotalActivoCorriente = totalActivoCorriente;
        mTotalActivoNoCorriente = totalActivoNoCorriente;
        mTotalPasivoCorriente = totalPasivoCorriente;
        mTotalPasivoNoCorriente = totalPasivoNoCorriente;
        mTotalPatrimonio = totalPatrimonio;
        mTotalPasivoPatrimonio = totalPasivoPatrimonio;
    }

    public List<ElementoEF> getActivos() {
        return mActivos;
    }

    public void setActivos(List<ElementoEF> activos) {
        mActivos = activos;
    }

    public List<ElementoEF> getPasivos() {
        return mPasivos;
    }

    public void setPasivos(List<ElementoEF> pasivos) {
        mPasivos = pasivos;
    }

    public List<ElementoEF> getPatrimonio() {
        return mPatrimonio;
    }

    public void setPatrimonio(List<ElementoEF> patrimonio) {
        mPatrimonio = patrimonio;
    }

    public double getTotalActivoCorriente() {
        return mTotalActivoCorriente;
    }

    public void setTotalActivoCorriente(double totalActivoCorriente) {
        mTotalActivoCorriente = totalActivoCorriente;
    }

    public double getTotalActivoNoCorriente() {
        return mTotalActivoNoCorriente;
    }

    public void setTotalActivoNoCorriente(double totalActivoNoCorriente) {
        mTotalActivoNoCorriente = totalActivoNoCorriente;
    }

    public double getTotalPasivoCorriente() {
        return mTotalPasivoCorriente;
    }

    public void setTotalPasivoCorriente(double totalPasivoCorriente) {
        mTotalPasivoCorriente = totalPasivoCorriente;
    }

    public double getTotalPasivoNoCorriente() {
        return mTotalPasivoNoCorriente;
    }

    public void setTotalPasivoNoCorriente(double totalPasivoNoCorriente) {
        mTotalPasivoNoCorriente = totalPasivoNoCorriente;
    }

    public double getTotalPatrimonio() {
        return mTotalPatrimonio;
    }

    public void setTotalPatrimonio(double totalPatrimonio) {
        mTotalPatrimonio = totalPatrimonio;
    }

    public double getTotalPasivoPatrimonio() {
        return mTotalPasivoPatrimonio;
    }

    public void setTotalPasivoPatrimonio(double totalPasivoPatrimonio) {
        mTotalPasivoPatrimonio = totalPasivoPatrimonio;
    }
}
