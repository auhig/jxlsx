package h.ji.jxlsx;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PoiTest {

    public static void main(String[] args) throws IOException {

        InputStream is = PoiTest.class.getResourceAsStream("/Test.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook(is)) {
            System.out.println(wb);
        }

    }

}
