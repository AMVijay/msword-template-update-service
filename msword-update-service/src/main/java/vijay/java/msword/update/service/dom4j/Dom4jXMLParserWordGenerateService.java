package vijay.java.msword.update.service.dom4j;

import java.io.File;
import java.util.Iterator;
import java.util.List;

import org.dom4j.Attribute;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.Node;
import org.dom4j.io.SAXReader;

import vijay.java.msword.update.service.IWordUpdateService;
import vijay.java.msword.update.service.WordUpdateModel;

public class Dom4jXMLParserWordGenerateService implements IWordUpdateService {

    @Override
    public void updateWordDocument(WordUpdateModel wordGenerateModel) {

        String filePath = System.getProperty("user.dir") + "/test-templates/template-bookmark.xml";
        try {

            File inputFile = new File(filePath);
            SAXReader reader = new SAXReader();
            Document document = reader.read(inputFile);

            List<Node> nodes = document.selectNodes("//*[local-name()='bookmarkStart']");
            for (Node node : nodes) {
                System.out.println("Node Present");
                Element element = (Element) node;
                System.out.println("node name " + element.getName());

                for(Attribute attribute : element.attributes()){
                    System.out.println("Attribute details :: " + attribute.getName() + "-" + attribute.getValue());
                }
               
                Iterator<Element> iterator = element.elementIterator("xmlData");

                while (iterator.hasNext()) {
                    Element marksElement = (Element) iterator.next();
                    System.out.println("Child Element " + marksElement.getName());
                    //marksElement.setText("80");
                }
            }

            // Pretty print the document to System.out
            // OutputFormat format = OutputFormat.createPrettyPrint();
            // XMLWriter writer;
            // writer = new XMLWriter(System.out, format);
            // writer.write(document);
        } catch (DocumentException e) {
            e.printStackTrace();
        } finally {

        }
    }

}