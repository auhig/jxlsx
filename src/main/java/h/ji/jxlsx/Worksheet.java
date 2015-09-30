package h.ji.jxlsx;

import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.Map;

import org.w3c.dom.Document;
import org.w3c.dom.Node;

public class Worksheet {

    private Workbook wb;
    private String name;
    private String tabColor;
    private int firstRowNum;
    private int firstColumnNum;
    private int lastRowNum;
    private int lastColumnNum;
    private Map<Integer, Row> rows = new LinkedHashMap<>();

    public Worksheet(Workbook wb, String name, Document doc) {
        this.wb = wb;
        this.name = name;

        // Sheet颜色
        Node tabColorNode = XmlUtil.getByXPath(doc, "/worksheet/sheetPr/tabColor");
        if (tabColorNode != null) {
            this.tabColor = XmlUtil.getAttributeValue(tabColorNode, "rgb");
        }

        // dimension
        Node dimensionNode = XmlUtil.getByXPath(doc, "/worksheet/dimension");
        String dimension = XmlUtil.getAttributeValue(dimensionNode, "ref");
        int splitIndex = dimension.indexOf(':');
        if (splitIndex > 0) {
            String first = dimension.substring(0, splitIndex);
            int[] fa = Cell.convertCellRef(first);
            this.firstRowNum = fa[0];
            this.firstColumnNum = fa[1];
            String last = dimension.substring(splitIndex + 1, dimension.length());
            int[] la = Cell.convertCellRef(last);
            this.lastRowNum = la[0];
            this.lastColumnNum = la[1];
        } else {
            int[] da = Cell.convertCellRef(dimension);
            this.firstRowNum = da[0];
            this.firstColumnNum = da[1];
            this.lastRowNum = da[0];
            this.lastColumnNum = da[1];
        }

        // Row
        XmlUtil.forEachByXPath(doc, "/worksheet/sheetData/row", n -> {
            Row r = new Row(this, n);
            rows.put(r.getNum(), r);
        });
    }

    public Workbook getWorkbook() {
        return wb;
    }

    public String getName() {
        return name;
    }

    public String getTabColor() {
        return tabColor;
    }

    public int getFirstRowNum() {
        return firstRowNum;
    }

    public int getFirstColumnNum() {
        return firstColumnNum;
    }

    public int getLastRowNum() {
        return lastRowNum;
    }

    public int getLastColumnNum() {
        return lastColumnNum;
    }

    public Cell getCell(int row, int column) {
        Row r = rows.get(row);
        return r == null ? null : r.getCell(column);
    }

    public Collection<Row> getRows() {
        return rows.values();
    }

}
