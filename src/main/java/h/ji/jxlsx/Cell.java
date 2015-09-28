package h.ji.jxlsx;

public class Cell {

    private Worksheet sh;
    private String cellRef;

    private String s;
    private String t;
    private String v;

    public Cell(Worksheet sh, String cellRef, String s, String t, String v) {
        this.sh = sh;
        this.cellRef = cellRef;
        this.s = s;
        this.t = t;
        this.v = v;
    }

    @SuppressWarnings("unchecked")
    public <T> T getValue() {
        if ("s".equals(t)) {
            int index = Integer.parseInt(v);
            return (T) sh.getWorkbook().getSharedString(index);
        }
        return (T) v;
    }

    public String getCellRef() {
        return cellRef;
    }

    public static String getCellRef(int row, int column) {
        return convertNumToColString(column) + (row + 1);
    }

    public static String convertNumToColString(int col) {
        int excelColNum = col + 1;

        StringBuilder colRef = new StringBuilder(2);
        int colRemain = excelColNum;

        while (colRemain > 0) {
            int thisPart = colRemain % 26;
            if (thisPart == 0) {
                thisPart = 26;
            }
            colRemain = (colRemain - thisPart) / 26;

            // The letter A is at 65
            char colChar = (char) (thisPart + 64);
            colRef.insert(0, colChar);
        }

        return colRef.toString();
    }

    public static int convertColStringToIndex(String ref) {
        int retval = 0;
        char[] refs = ref.toUpperCase().toCharArray();
        for (int k = 0; k < refs.length; k++) {
            char thechar = refs[k];
            // Character is uppercase letter, find relative value to A
            retval = retval * 26 + thechar - 'A' + 1;
        }
        return retval - 1;
    }

}
