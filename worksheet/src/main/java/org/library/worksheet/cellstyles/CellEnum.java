package org.library.worksheet.cellstyles;

import java.util.Arrays;

public enum CellEnum {
    DEFAULT_HEADER, DEFAULT_TITLE, DEFAULT_CELL,

    TEAL_HEADER, TEAL_TITLE, TEAL_CELL,

    FORMULA_1, FORMULA_2, FORMULA_3,

    LINE_1, LINE_2;

    public static boolean contains(String test) {

        for (CellEnum c : CellEnum.values()) {
            if (c.name().equals(test)) {
                return true;
            }
        }

        return false;
    }
}
