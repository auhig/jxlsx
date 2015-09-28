package h.ji.jxlsx;

import static org.junit.Assert.assertEquals;

import org.junit.Test;

public class CellTest {

    @Test
    public void testGetCellRef_normal() {
        String result = Cell.getCellRef(1, 2);
        assertEquals("C2", result);
    }

    @Test
    public void testConvertNumToColString_normal() {
        assertEquals("A", Cell.convertNumToColString(0));
        assertEquals("B", Cell.convertNumToColString(1));
        assertEquals("Z", Cell.convertNumToColString(25));
        assertEquals("AA", Cell.convertNumToColString(26));
    }

    @Test
    public void testConvertColStringToIndex_normal() {
        assertEquals(0, Cell.convertColStringToIndex("A"));
        assertEquals(1, Cell.convertColStringToIndex("B"));
        assertEquals(25, Cell.convertColStringToIndex("Z"));
        assertEquals(26, Cell.convertColStringToIndex("AA"));
    }

}