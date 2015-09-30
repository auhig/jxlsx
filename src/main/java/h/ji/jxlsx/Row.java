package h.ji.jxlsx;

import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.Map;

import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class Row {

    private Worksheet sh;
    private int num;
    private Map<Integer, Cell> cells = new LinkedHashMap<>();

    public Row(Worksheet sh, Node node) {
        this.sh = sh;
        String r = XmlUtil.getAttributeValue(node, "r");
        this.num = Integer.parseInt(r);
        NodeList children = node.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node n = children.item(i);
            if (!"c".equals(n.getNodeName())) {
                continue;
            }
            Cell cell = new Cell(this, n);
            this.cells.put(cell.getColumnNum(), cell);
        }
    }

    public Worksheet getWorksheet() {
        return sh;
    }

    public int getNum() {
        return num;
    }

    public Cell getCell(int column) {
        return cells.get(column);
    }

    public Collection<Cell> getCells() {
        return cells.values();
    }

}
