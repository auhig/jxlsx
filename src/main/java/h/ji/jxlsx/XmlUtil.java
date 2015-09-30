package h.ji.jxlsx;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.function.Consumer;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

public class XmlUtil {

    private static DocumentBuilderFactory documentBuilderFactory = DocumentBuilderFactory.newInstance();
    private static XPathFactory xPathFactory = XPathFactory.newInstance();

    public static Document parse(byte[] ba) throws IOException {
        return parse(new ByteArrayInputStream(ba));
    }

    public static Document parse(InputStream is) throws IOException {
        try {
            return documentBuilderFactory.newDocumentBuilder().parse(is);
        } catch (SAXException | ParserConfigurationException e) {
            throw new RuntimeException(e);
        }
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

    public static void forEachByXPath(Node n, String expression, Consumer<Node> action) {
        XPath path = xPathFactory.newXPath();
        try {
            NodeList nl = (NodeList) path.evaluate(expression, n, XPathConstants.NODESET);
            forEach(nl, action);
        } catch (XPathExpressionException e) {
            throw new RuntimeException(e);
        }
    }

    public static Node getByXPath(Node n, String expression) {
        XPath path = xPathFactory.newXPath();
        try {
            return (Node) path.evaluate(expression, n, XPathConstants.NODE);
        } catch (XPathExpressionException e) {
            throw new RuntimeException(e);
        }
    }

    public static Node getAttribute(Node n, String name) {
        return n.getAttributes().getNamedItem(name);
    }

    public static String getAttributeValue(Node n, String name) {
        Node a = getAttribute(n, name);
        return a == null ? null : a.getNodeValue();
    }

}
