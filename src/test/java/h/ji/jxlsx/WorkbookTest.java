package h.ji.jxlsx;

import java.io.InputStream;

import org.junit.Test;

public class WorkbookTest {

    @Test
    public void test_normal() throws Exception {

        try (InputStream is = WorkbookTest.class.getResourceAsStream("/Test.xlsx")) {
            Workbook wb = new Workbook(is);
            wb.getSheetNames().forEach(n -> {
                Worksheet sh = wb.getSheet(n);
                System.out.println(sh.getName() + " ------------------------");
                sh.getCells().forEach((c) -> {
                    System.out.println(c.getRef() + " = " + c.getValue());
                });
                System.out.println();

                if (n.equals("Sheet1")) {
                    System.out.println(sh.getCell("C6").getDateValue());
                    System.out.println(sh.getCell("C7").getDateValue());
                }
            });
        }

        //        try (InputStream is = WorkbookTest.class.getResourceAsStream("/EquipLevelPackage.xlsx")) {
        //            Workbook wb = new Workbook(is);
        //            wb.getSheetNames().forEach(n -> {
        //                Worksheet sh = wb.getSheet(n);
        //                System.out.println(sh.getName() + " ------------------------");
        //                sh.getCells().forEach((c) -> {
        //                    System.out.println(c.getRef() + " = " + c.getValue());
        //                });
        //                System.out.println();
        //            });
        //        }

    }
}
