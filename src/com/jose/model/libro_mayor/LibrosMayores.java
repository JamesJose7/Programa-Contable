package com.jose.model.libro_mayor;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by agua on 08/07/15.
 */
public class LibrosMayores {
    private List<LibroMayor> mLibrosMayoresList;

    public LibrosMayores() {
        mLibrosMayoresList = new ArrayList<>();
    }

    public void addLibroMayor(LibroMayor libroMayor) {
        mLibrosMayoresList.add(libroMayor);
    }

    public LibroMayor getLibroMayor(int index) {
        return mLibrosMayoresList.get(index);
    }

    public List<LibroMayor> getLibrosMayoresList() {
        return mLibrosMayoresList;
    }
}
