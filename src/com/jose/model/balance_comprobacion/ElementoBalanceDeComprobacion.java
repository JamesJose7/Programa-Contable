package com.jose.model.balance_comprobacion;

/**
 * Created by agua on 09/07/15.
 */
public class ElementoBalanceDeComprobacion {
    private int mNumero;
    private String mCodigo;
    private String mCuenta;
    private double mSumaDebe;
    private double mSumaHaber;
    private double mSaldoDebe;
    private double mSaldoHaber;

    public ElementoBalanceDeComprobacion(int numero, String codigo, String cuenta, double sumaDebe, double sumaHaber, double saldoDebe, double saldoHaber) {
        mNumero = numero;
        mCodigo = codigo;
        mCuenta = cuenta;
        mSumaDebe = sumaDebe;
        mSumaHaber = sumaHaber;
        mSaldoDebe = saldoDebe;
        mSaldoHaber = saldoHaber;
    }

    public int getNumero() {
        return mNumero;
    }

    public void setNumero(int numero) {
        mNumero = numero;
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

    public double getSumaDebe() {
        return mSumaDebe;
    }

    public void setSumaDebe(double sumaDebe) {
        mSumaDebe = sumaDebe;
    }

    public double getSumaHaber() {
        return mSumaHaber;
    }

    public void setSumaHaber(double sumaHaber) {
        mSumaHaber = sumaHaber;
    }

    public double getSaldoDebe() {
        return mSaldoDebe;
    }

    public void setSaldoDebe(double saldoDebe) {
        mSaldoDebe = saldoDebe;
    }

    public double getSaldoHaber() {
        return mSaldoHaber;
    }

    public void setSaldoHaber(double saldoHaber) {
        mSaldoHaber = saldoHaber;
    }
}
