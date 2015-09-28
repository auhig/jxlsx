package h.ji.jxlsx;

public class Cell {

    private final Worksheet sh;

    private String v;

    public Cell(Worksheet sh, String v) {
        this.sh = sh;
        this.v = v;
    }

    public Object getValue() {
        return v;
    }

    public static String getCellRef(int row, int column) {
        StringBuilder sb = new StringBuilder(4);
        int colRemain = column + 1;
        while (colRemain > 0) {
            int thisPart = colRemain % 26;
            if (thisPart == 0) {
                thisPart = 26;
            }
            colRemain = (colRemain - thisPart) / 26;
            char colChar = (char) (thisPart + 64);  // The letter A is at 65
            sb.insert(0, colChar);
        }
        sb.append(row + 1);
        return sb.toString();
    }

}
