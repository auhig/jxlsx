package h.ji.jxlsx;

import java.io.IOException;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.junit.Test;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

public class XMLUtilTest {

    @Test
    public void test_normal() throws SAXException, IOException, ParserConfigurationException, XPathExpressionException {
        Document doc = XMLUtil.parse(XMLUtilTest.class.getResourceAsStream("/Test/xl/workbook.xml"));
        XPath path = XPathFactory.newInstance().newXPath();
        XPathExpression expression = path.compile("//workbook/sheets/sheet");
        NodeList nl = (NodeList) expression.evaluate(doc, XPathConstants.NODESET);
        System.out.println(nl.getLength());
        XMLUtil.forEach(nl, n -> {
            System.out.println(XMLUtil.getAttribute(n, "name"));
            System.out.println(XMLUtil.getAttribute(n, "r:id"));
        });
    }

}
