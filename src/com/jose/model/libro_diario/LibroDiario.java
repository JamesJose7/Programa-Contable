package com.jose.model.libro_diario;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by agua on 30/06/15.
 */
public class LibroDiario {
    private List<Asiento> mAsientos;

    public LibroDiario() {
        mAsientos = new ArrayList<>();
    }

    public void addAsiento(Asiento asiento) {
        mAsientos.add(asiento);
    }

    public List<Asiento> getAsientos() {
        return mAsientos;
    }
}
