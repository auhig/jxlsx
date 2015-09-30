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
                System.out.println(sh.getName() + " " + sh.getTabColor() + " ------------------------");
                System.out.println(sh.getFirstRowNum());
                System.out.println(sh.getFirstColumnNum());
                System.out.println(sh.getLastRowNum());
                System.out.println(sh.getLastColumnNum());
                sh.getRows().forEach(r -> {
                    r.getCells().forEach(c -> {
                        System.out.println(c.getRef() + " = " + c.getValue());
                    });
                });
                System.out.println();
            });
        }

        //        try (InputStream is = WorkbookTest.class.getResourceAsStream("/EquipLevelPackage.xlsx")) {
        //            Workbook wb = new Workbook(is);
        //            wb.getSheetNames().forEach(n -> {
        //                Worksheet sh = wb.getSheet(n);
        //                System.out.println(sh.getName() + " ------------------------");
        //                System.out.println(sh.getFirstRowNum());
        //                System.out.println(sh.getFirstColumnNum());
        //                System.out.println(sh.getLastRowNum());
        //                System.out.println(sh.getLastColumnNum());
        //                sh.getRows().forEach(r -> {
        //                    r.getCells().forEach(c -> {
        //                        System.out.println(c.getRef() + " = " + c.getValue());
        //                    });
        //                });
        //                System.out.println();
        //            });
        //        }

        System.out.println();

    }
}
