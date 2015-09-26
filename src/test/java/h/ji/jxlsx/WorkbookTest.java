package h.ji.jxlsx;

import java.io.IOException;
import java.io.InputStream;

import org.junit.Test;

public class WorkbookTest {

    @Test
    public void test_normal() throws IOException {

        try (InputStream is = WorkbookTest.class.getResourceAsStream("/Test.xlsx")) {
            new Workbook(is);
        }

    }
}
