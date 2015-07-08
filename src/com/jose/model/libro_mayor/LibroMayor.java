package com.jose.model.libro_mayor;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by agua on 08/07/15.
 */
public class LibroMayor {
    private String mCuenta;
    private String mCodigo;
    private List<ElementoMayor> mElementosMayores;

    public LibroMayor(String cuenta, String codigo, List<ElementoMayor> elementosMayores) {
        mElementosMayores = new ArrayList<>(elementosMayores);
        mCuenta = cuenta;
        mCodigo = codigo;
    }

    public String getCuenta() {
        return mCuenta;
    }

    public void setCuenta(String cuenta) {
        mCuenta = cuenta;
    }

    public String getCodigo() {
        return mCodigo;
    }

    public void setCodigo(String codigo) {
        mCodigo = codigo;
    }

    public List<ElementoMayor> getElementosMayores() {
        return mElementosMayores;
    }

    public void addElementoMayor(ElementoMayor elementoMayor) {
        mElementosMayores.add(elementoMayor);
    }
}
