package h.ji.jxlsx;

import javax.xml.xpath.XPathExpressionException;

import java.util.LinkedHashMap;
import java.util.Map;

import org.w3c.dom.Document;

public class Worksheet {

    private final Workbook wb;
    public final Map<String, Cell> cells = new LinkedHashMap<>();
    private String name;

    public Worksheet(Workbook wb, String name, Document doc) throws XPathExpressionException {
        this.wb = wb;
        this.name = name;
        XMLUtil.forEachByXPath(doc, "/worksheet/sheetData/row", n -> {
            XMLUtil.forEach(n.getChildNodes(), c -> {
                String ref = XMLUtil.getAttributeValue(c, "r"); // 例如“C3”
                String v = null;
                if (c.getChildNodes().getLength() == 1) {
                    v = c.getChildNodes().item(0).getTextContent();
                } else if (c.getChildNodes().getLength() != 0) {
                    throw new IllegalArgumentException("Size of child node must equal 1 or 0");
                }
                Cell cell = new Cell(this, v);
                this.cells.put(ref, cell);
            });
        });
    }

}
