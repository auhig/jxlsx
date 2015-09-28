package h.ji.jxlsx;

import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.Map;

import javax.xml.xpath.XPathExpressionException;

import org.w3c.dom.Document;

public class Worksheet {

    private Workbook wb;
    private Map<String, Cell> cells = new LinkedHashMap<>();
    private String name;

    public Worksheet(Workbook wb, String name, Document doc) throws XPathExpressionException {
        this.wb = wb;
        this.name = name;
        XMLUtil.forEachByXPath(doc, "/worksheet/sheetData/row", n -> {
            XMLUtil.forEach(n.getChildNodes(), c -> {
                String ref = XMLUtil.getAttributeValue(c, "r"); // 例如“C3”
                String s = XMLUtil.getAttributeValue(c, "s");
                String t = XMLUtil.getAttributeValue(c, "t");
                String v = null;
                if (c.getChildNodes().getLength() == 1) {
                    v = c.getChildNodes().item(0).getTextContent();
                } else if (c.getChildNodes().getLength() != 0) {
                    throw new IllegalArgumentException("Size of child node must equal 1 or 0");
                }
                Cell cell = new Cell(this, ref, s, t, v);
                this.cells.put(ref, cell);
            });
        });
    }

    public Workbook getWorkbook() {
        return wb;
    }

    public Collection<Cell> getCells() {
        return cells.values();
    }

}
