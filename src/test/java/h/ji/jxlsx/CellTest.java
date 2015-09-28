package h.ji.jxlsx;

import static org.junit.Assert.assertEquals;

import org.junit.Test;

public class CellTest {

    @Test
    public void testGetCellRef_normal() {
        String result = Cell.getCellRef(1, 2);
        assertEquals("C2", result);
    }

}