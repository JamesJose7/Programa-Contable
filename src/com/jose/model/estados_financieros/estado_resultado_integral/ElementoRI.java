package com.jose.model.estados_financieros.estado_resultado_integral;

/**
 * Created by agua on 15/07/15.
 */
public class ElementoRI {
    private String mCodigo;
    private String mCuenta;
    private double mValor;

    public ElementoRI(String codigo, String cuenta, double valor) {
        mCodigo = codigo;
        mCuenta = cuenta;
        mValor = valor;
    }

    public String getCodigo() {
        return mCodigo;
    }

    public void setCodigo(String codigo) {
        mCodigo = codigo;
    }

    public String getCuenta() {
        return mCuenta;
    }

    public void setCuenta(String cuenta) {
        mCuenta = cuenta;
    }

    public double getValor() {
        return mValor;
    }

    public void setValor(double valor) {
        mValor = valor;
    }
}
