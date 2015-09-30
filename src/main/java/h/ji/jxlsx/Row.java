package h.ji.jxlsx;

import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.Map;

import org.w3c.dom.Node;

public class Row {

    private Worksheet sh;
    private int number;
    private Map<Integer, Cell> cells = new LinkedHashMap<>();

    public Row(Worksheet sh, Node node) {
        this.sh = sh;
        String r = XmlUtil.getAttributeValue(node, "r");
        this.number = Integer.parseInt(r);
        XmlUtil.forEachByXPath(node, "c", n -> {
            Cell cell = new Cell(this, n);
            this.cells.put(cell.getColumnNumber(), cell);
        });
    }

    public Worksheet getWorksheet() {
        return sh;
    }

    public int getNumber() {
        return number;
    }

    public Cell getCell(int column) {
        return cells.get(column);
    }

    public Collection<Cell> getCells() {
        return cells.values();
    }

}
