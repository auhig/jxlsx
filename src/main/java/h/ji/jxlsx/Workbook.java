package h.ji.jxlsx;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Stream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPathExpressionException;

import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.xml.sax.SAXException;

public class Workbook {

    private static final String[] IGNORE_FILE_PREFIX = {"xl/theme/", "xl/styles.xml"};

    private boolean date1904 = false;

    //    private final Map<String, String> rels = new HashMap<>();
    private final List<String> sharedStrings = new ArrayList<>();
    private final Map<String, Worksheet> sheets = new LinkedHashMap<>();

    public Workbook(InputStream is) throws IOException, SAXException, ParserConfigurationException, XPathExpressionException {
        Map<String, Document> documents = new HashMap<>();
        ZipInputStream zis = new ZipInputStream(is);
        for (ZipEntry ze = zis.getNextEntry(); ze != null; ze = zis.getNextEntry()) {
            if (ze.isDirectory()) {
                continue;
            }
            String name = ze.getName();
            if (Stream.of(IGNORE_FILE_PREFIX).anyMatch(name::startsWith)) {
                continue;
            }
            byte[] ba = readZip(zis, ze);
            Document doc = XmlUtil.parse(ba);
            documents.put(name, doc);
        }
        parse(documents);
    }

    private byte[] readZip(ZipInputStream is, ZipEntry ze) throws IOException {
        if (ze.getSize() > Integer.MAX_VALUE) {
            throw new IllegalArgumentException("Extra data too large, size is " + ze.getSize());
        }
        byte[] ba = new byte[(int) ze.getSize()];
        int n = 0;
        while (n < ba.length) {
            n += is.read(ba, n, ba.length - n);
        }
        return ba;
    }

    private void parse(Map<String, Document> documents) {

        // 读取workbook.xml.rels
        Map<String, String> relsIdTargetMap = new HashMap<>();
        Document relsDoc = documents.get("xl/_rels/workbook.xml.rels");
        XmlUtil.forEachByTagName(relsDoc, "Relationship", n -> {
            String id = XmlUtil.getAttributeValue(n, "Id");
            String target = XmlUtil.getAttributeValue(n, "Target");
            relsIdTargetMap.put(id, target);
        });

        // 读取sharedStrings.xml
        Document sharedStringsDoc = documents.get("xl/sharedStrings.xml");
        if (sharedStringsDoc != null) {
            XmlUtil.forEachByXPath(sharedStringsDoc, "/sst/si/t", n -> sharedStrings.add(n.getTextContent()));
        }

        // 读取workbook.xml
        Document workbookDoc = documents.get("xl/workbook.xml");
        XmlUtil.forEachByXPath(workbookDoc, "/workbook/sheets/sheet", n -> {
            String name = XmlUtil.getAttributeValue(n, "name"); // Sheet的名字
            String rId = XmlUtil.getAttributeValue(n, "r:id");  // rId
            String target = relsIdTargetMap.get(rId);           // rId对应的Target
            Document sheetDoc = documents.get("xl/" + target);  // 获取该Sheet对应的Document
            Worksheet sh = new Worksheet(this, name, sheetDoc);
            sheets.put(name, sh);
        });
        // workbookPr
        Node workbookPr = XmlUtil.getByXPath(workbookDoc, "/workbook/workbookPr");
        if (workbookPr != null) {
            String v = XmlUtil.getAttributeValue(workbookPr, "date1904");
            if ("1".equals(v)) {
                date1904 = true;
            }
        }

    }

    String getSharedString(int index) {
        return sharedStrings.get(index);
    }

    boolean isDate1904() {
        return date1904;
    }

    public Set<String> getSheetNames() {
        return sheets.keySet();
    }

    public Worksheet getSheet(String name) {
        return sheets.get(name);
    }

}
