package h.ji.jxlsx;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.Charset;
import java.util.stream.Stream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

public class Workbook {

    private static final String[] IGNORE_FILE_PREFIX = { "xl/theme/", "xl/styles.xml" };

    public Workbook(InputStream is) throws IOException {

        ZipInputStream zis = new ZipInputStream(is);
        ZipEntry ze;
        while ((ze = zis.getNextEntry()) != null) {

            if (ze.isDirectory()) {
                continue;
            }

            String name = ze.getName();
            boolean ignore = Stream.of(IGNORE_FILE_PREFIX).anyMatch(p -> name.toLowerCase().startsWith(p));
            if (ignore) {
                continue;
            }

            byte[] ba = readZip(zis, ze);
            System.out.println(name);
            System.out.println(new String(ba, Charset.forName("UTF-8")));
            System.out.println();
        }

    }

    private byte[] readZip(ZipInputStream is, ZipEntry ze) throws IOException {
        if (ze.getSize() > Integer.MAX_VALUE) {
            throw new IllegalArgumentException("Extra data too large, size is " + ze.getSize());
        }
        byte[] result = new byte[(int) ze.getSize()];
        int count = 0;
        while (count < result.length) {
            count += is.read(result, count, result.length - count);
        }
        return result;
    }

}
