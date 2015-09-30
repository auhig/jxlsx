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
    private Map<Integer, Row> rows = new LinkedHashMap<>();

    public Worksheet(Workbook wb, String name, Document doc) {
        this.wb = wb;
        this.name = name;
        Node tabColor = XmlUtil.getByXPath(doc, "/worksheet/sheetPr/tabColor");
        if (tabColor != null) {
            this.tabColor = XmlUtil.getAttributeValue(tabColor, "rgb");
        }
        XmlUtil.forEachByXPath(doc, "/worksheet/sheetData/row", n -> {
            Row r = new Row(this, n);
            rows.put(r.getNumber(), r);
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

    public Cell getCell(int row, int column) {
        Row r = rows.get(row);
        return r == null ? null : r.getCell(column);
    }

    public Collection<Row> getRows() {
        return rows.values();
    }

}
