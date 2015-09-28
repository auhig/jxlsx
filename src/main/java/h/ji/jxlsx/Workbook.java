package h.ji.jxlsx;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPathExpressionException;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import org.w3c.dom.Document;
import org.xml.sax.SAXException;

public class Workbook {

    private static final String[] IGNORE_FILE_PREFIX = {"xl/theme/", "xl/styles.xml"};

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
            Document doc = XMLUtil.parse(ba);
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

    private void parse(Map<String, Document> documents) throws XPathExpressionException {

        // 读取workbook.xml.rels
        Map<String, String> relsIdTargetMap = new HashMap<>();
        Document relsDoc = documents.get("xl/_rels/workbook.xml.rels");
        XMLUtil.forEachByTagName(relsDoc, "Relationship", n -> {
            String id = XMLUtil.getAttributeValue(n, "Id");
            String target = XMLUtil.getAttributeValue(n, "Target");
            relsIdTargetMap.put(id, target);
        });

        // 读取sharedStrings.xml
        Document sharedStringsDoc = documents.get("xl/sharedStrings.xml");
        XMLUtil.forEachByXPath(sharedStringsDoc, "/sst/si/t", n -> sharedStrings.add(n.getTextContent()));

        // 读取workbook.xml
        Document workbookDoc = documents.get("xl/workbook.xml");
        XMLUtil.forEachByXPath(workbookDoc, "/workbook/sheets/sheet", n -> {
            String name = XMLUtil.getAttributeValue(n, "name"); // Sheet的名字
            String rId = XMLUtil.getAttributeValue(n, "r:id");  // rId
            String target = relsIdTargetMap.get(rId);           // rId对应的Target
            Document sheetDoc = documents.get("xl/" + target);  // 获取该Sheet对应的Document
            Worksheet sh = null;
            try {
                sh = new Worksheet(this, name, sheetDoc);
            } catch (XPathExpressionException e) {
                e.printStackTrace();
            }
            sheets.put(name, sh);
        });

    }

    String getSharedString(int index) {
        return sharedStrings.get(index);
    }

    public Worksheet getSheet(String name) {
        return sheets.get(name);
    }

}
