package vijay.poc.msword.update.service.docx4j;

import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.xml.bind.JAXBException;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Text;

import vijay.poc.msword.update.service.IWordUpdateService;
import vijay.poc.msword.update.service.WordUpdateModel;

/**
 * Method to update Text nodes in word by matching the exact phrase.
 * @author Vijay
 *
 */
public class TextNodeUpdateService implements IWordUpdateService {

    @Override
    public void updateWordDocument(WordUpdateModel wordGenerateModel) {

        try {

            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(wordGenerateModel.getInputFilePath());
            MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

            replaceTextContent(documentPart, wordGenerateModel.getWordContent());

            // save the docx...
            wordMLPackage.save(wordGenerateModel.getOutputFilePath());

        } catch (Docx4JException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Method to replace the Text File in Word Document.
     */
    private void replaceTextContent(MainDocumentPart documentPart, Map<String, String> map)
            throws XPathBinderAssociationIsPartialException, JAXBException {

        Set<org.docx4j.wml.Text> textNodeSet = new HashSet<org.docx4j.wml.Text>();
        map.entrySet().forEach(entry -> {

            // Exact Match
            //String xpath = "//w:t[(text()='" + entry.getKey() + "')]";

            // Contains
            String xpath = "//w:t[contains(text(),'" + entry.getKey() + "')]";

            //System.out.println("XPath :: " + xpath);
            try {
                List<Object> list = documentPart.getJAXBNodesViaXPath(xpath, false);
                System.out.println("xPATH :: " + xpath + " Node Count " + list.size());
                list.forEach(textNode -> {
                    Object object = XmlUtils.unwrap(textNode);
                    if (object instanceof org.docx4j.wml.Text) {
                        textNodeSet.add((Text) object);
                    }
                });

            } catch (XPathBinderAssociationIsPartialException | JAXBException e) {
                e.printStackTrace();
            }

        });

        for (Text textNode : textNodeSet) {
            String textNodeValue = textNode.getValue();
            String[] stringTokenArray = textNodeValue.split("\\s+");

            for (String stringToken : stringTokenArray) {
                System.out.println("String Token :: " + stringToken);
                String tokenValue = map.get(stringToken);
                if(tokenValue != null)
                    textNodeValue = textNodeValue.replace(stringToken, tokenValue);
            }

            textNode.setValue(textNodeValue);
        }
    }
}