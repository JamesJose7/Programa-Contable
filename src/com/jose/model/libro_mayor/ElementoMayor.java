package com.jose.model.libro_mayor;

/**
 * Created by agua on 08/07/15.
 */
public class ElementoMayor {
    private String mFecha;
    private String mDetalle;
    private int mReferencia;
    private double mDebe;
    private double mHaber;
    private double mSaldo;

    public ElementoMayor(String fecha, String detalle, int referencia, double debe, double haber, double saldo) {
        mFecha = fecha;
        mDetalle = detalle;
        mReferencia = referencia;
        mDebe = debe;
        mHaber = haber;
        mSaldo = saldo;
    }

    public String getFecha() {
        return mFecha;
    }

    public void setFecha(String fecha) {
        mFecha = fecha;
    }

    public String getDetalle() {
        return mDetalle;
    }

    public void setDetalle(String detalle) {
        mDetalle = detalle;
    }

    public int getReferencia() {
        return mReferencia;
    }

    public void setReferencia(int referencia) {
        mReferencia = referencia;
    }

    public double getDebe() {
        return mDebe;
    }

    public void setDebe(double debe) {
        mDebe = debe;
    }

    public double getHaber() {
        return mHaber;
    }

    public void setHaber(double haber) {
        mHaber = haber;
    }

    public double getSaldo() {
        return mSaldo;
    }

    public void setSaldo(double saldo) {
        mSaldo = saldo;
    }
}
