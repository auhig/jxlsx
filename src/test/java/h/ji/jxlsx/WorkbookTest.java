package h.ji.jxlsx;

import java.io.InputStream;

import org.junit.Test;

public class WorkbookTest {

    @Test
    public void test_normal() throws Exception {

        try (InputStream is = WorkbookTest.class.getResourceAsStream("/Test.xlsx")) {
            Workbook wb = new Workbook(is);
            Worksheet sh = wb.getSheet("Sheet1");
            sh.getCells().forEach((c) -> {
                System.out.println(c.getCellRef() + " = " + c.getValue());
            });
        }

    }
}
