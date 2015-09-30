package h.ji.jxlsx;

import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.TimeZone;

import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class Cell {

    public static final int SECONDS_PER_MINUTE = 60;
    public static final int MINUTES_PER_HOUR = 60;
    public static final int HOURS_PER_DAY = 24;
    public static final int SECONDS_PER_DAY = HOURS_PER_DAY * MINUTES_PER_HOUR * SECONDS_PER_MINUTE;
    public static final long DAY_MILLISECONDS = SECONDS_PER_DAY * 1000L;

    private Row row;
    private String ref;
    private int columnNumber;
    private String s;
    private String t;
    private String v;

    public Cell(Row row, Node node) {
        this.row = row;
        this.ref = XmlUtil.getAttributeValue(node, "r"); // 例如“C3”
        this.columnNumber = convertCellRef(ref)[1];
        this.s = XmlUtil.getAttributeValue(node, "s");
        this.t = XmlUtil.getAttributeValue(node, "t");
        NodeList nl = node.getChildNodes();
        for (int i = 0; i < nl.getLength(); i++) {
            if ("v".equals(nl.item(i).getNodeName())) {
                v = nl.item(i).getTextContent();
            }
        }
    }

    public Row getRow() {
        return row;
    }

    public String getRef() {
        return ref;
    }

    public int getColumnNumber() {
        return columnNumber;
    }

    public int getRowNumber() {
        return row.getNumber();
    }

    public String getStyle() {
        return s;
    }

    public String getRawValue() {
        return v;
    }

    public String getValue() {
        if ("s".equals(t)) {
            int index = Integer.parseInt(v);
            return row.getWorksheet().getWorkbook().getSharedString(index);
        }
        return v;
    }

    public Date getDateValue() {
        double date = getNumericValue();
        boolean use1904windowing = row.getWorksheet().getWorkbook().isDate1904();
        Calendar cal = getJavaCalendar(date, use1904windowing, null, false);
        return cal.getTime();
    }

    public static Calendar getJavaCalendar(double date, boolean use1904windowing, TimeZone timeZone, boolean roundSeconds) {
        if (!isValidExcelDate(date)) {
            return null;
        }
        int wholeDays = (int)Math.floor(date);
        int millisecondsInDay = (int)((date - wholeDays) * DAY_MILLISECONDS + 0.5);
        Calendar calendar;
        if (timeZone != null) {
            calendar = new GregorianCalendar(timeZone);
        } else {
            calendar = new GregorianCalendar();     // using default time-zone
        }
        setCalendar(calendar, wholeDays, millisecondsInDay, use1904windowing, roundSeconds);
        return calendar;
    }

    public static boolean isValidExcelDate(double value) {
        return value > -Double.MIN_VALUE;
    }

    public static void setCalendar(Calendar calendar, int wholeDays,
            int millisecondsInDay, boolean use1904windowing,
            boolean roundSeconds) {
        int startYear = 1900;
        int dayAdjust = -1; // Excel thinks 2/29/1900 is a valid date, which it isn't
        if (use1904windowing) {
            startYear = 1904;
            dayAdjust = 1;  // 1904 date windowing uses 1/2/1904 as the first day
        } else if (wholeDays < 61) {
            // Date is prior to 3/1/1900, so adjust because Excel thinks 2/29/1900 exists
            // If Excel date == 2/29/1900, will become 3/1/1900 in Java representation
            dayAdjust = 0;
        }
        calendar.set(startYear, Calendar.JANUARY, wholeDays + dayAdjust, 0, 0, 0);
        calendar.set(Calendar.MILLISECOND, millisecondsInDay);
        if (roundSeconds) {
            calendar.add(Calendar.MILLISECOND, 500);
            calendar.clear(Calendar.MILLISECOND);
        }
    }

    public double getNumericValue() {
        return Double.parseDouble(v);
    }

    public String getStringValue() {
        return getValue();
    }

    public boolean isNull() {
        return v == null;
    }

    public static String getCellRef(int row, int column) {
        return convertNumToColString(column) + (row + 1);
    }

    public static int[] convertCellRef(String ref) {
        char[] ca = ref.toCharArray();
        String rowStr = null;
        String colStr = null;
        for (int i = 0; i < ca.length; i++) {
            char c = ca[i];
            if (c >= '0' && c <= '9') {
                rowStr = ref.substring(i);
                colStr = ref.substring(0, i);
                break;
            }
        }
        if (rowStr == null || colStr == null) {
            throw new IllegalArgumentException("Error has occurred while parsing " + ref);
        }
        int row = Integer.parseInt(rowStr) - 1;
        int col = convertColStringToIndex(colStr);
        return new int[] { row, col };
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
        for (char thechar : refs) {
            // Character is uppercase letter, find relative value to A
            retval = retval * 26 + thechar - 'A' + 1;
        }
        return retval - 1;
    }

}
