package h.ji.jxlsx;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.function.Consumer;

import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

public class XMLUtil {

    private static DocumentBuilderFactory documentBuilderFactory = DocumentBuilderFactory.newInstance();
    private static XPathFactory xPathFactory = XPathFactory.newInstance();

    public static Document parse(byte[] ba) throws SAXException, IOException, ParserConfigurationException {
        return parse(new ByteArrayInputStream(ba));
    }

    public static Document parse(InputStream is) throws SAXException, IOException, ParserConfigurationException {
        return documentBuilderFactory.newDocumentBuilder().parse(is);
    }

    public static void forEach(NodeList nl, Consumer<Node> action) {
        for (int i = 0; i < nl.getLength(); i++) {
            action.accept(nl.item(i));
        }
    }

    public static void forEachByTagName(Document doc, String tagname, Consumer<Node> action) {
        NodeList nl = doc.getElementsByTagName(tagname);
        forEach(nl, action);
    }

    public static void forEachByXPath(Node n, String expression, Consumer<Node> action) throws XPathExpressionException {
        XPath path = xPathFactory.newXPath();
        NodeList nl = (NodeList) path.evaluate(expression, n, XPathConstants.NODESET);
        forEach(nl, action);
    }

    public static Node getAttribute(Node n, String name) {
        return n.getAttributes().getNamedItem(name);
    }

    public static String getAttributeValue(Node n, String name) {
        Node a = getAttribute(n, name);
        return a == null ? null : a.getNodeValue();
    }

}
