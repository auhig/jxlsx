package h.ji.jxlsx;

import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.Map;

import org.w3c.dom.Document;

public class Worksheet {

    private Workbook wb;
    private String name;
    private Map<String, Cell> cells = new LinkedHashMap<>();

    public Worksheet(Workbook wb, String name, Document doc) {
        this.wb = wb;
        this.name = name;
        XmlUtil.forEachByXPath(doc, "/worksheet/sheetData/row", n -> {
            XmlUtil.forEach(n.getChildNodes(), c -> {
                Cell cell = new Cell(this, c);
                this.cells.put(cell.getRef(), cell);
            });
        });
    }

    public Workbook getWorkbook() {
        return wb;
    }

    public String getName() {
        return name;
    }

    public Cell getCell(int row, int column) {
        String ref = Cell.getCellRef(row, column);
        return cells.get(ref);
    }

    public Cell getCell(String ref) {
        return cells.get(ref);
    }

    public Collection<Cell> getCells() {
        return cells.values();
    }

    public int getRowNum() {
        return 0;
    }

    public int getColumnNum() {
        return 0;
    }

}
