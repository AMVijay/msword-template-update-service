package vijay.poc.msword.update.service.docx4j;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.xml.bind.JAXBException;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.model.fields.merge.DataFieldName;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Text;

import vijay.poc.msword.update.service.IWordUpdateService;
import vijay.poc.msword.update.service.WordUpdateModel;

/**
 * Service Class to update Content Controls
 * 
 * @author Vijay
 *
 */
public class WordContentControlUpdateService implements IWordUpdateService {

	@Override
	public void updateWordDocument(WordUpdateModel wordUpdateModel) {

		try {

			Map<DataFieldName, String> map = new HashMap<DataFieldName, String>();
			wordUpdateModel.getBookmarkContent().entrySet().forEach(entry -> {
				map.put(new DataFieldName(entry.getKey()), entry.getValue());
			});

			System.out.println("Load Docx Content as Object");
			WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(wordUpdateModel.getInputFilePath());
			MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
			
			System.out.println("Going to update Content Controls");
			replaceContentControls(documentPart, wordUpdateModel.getBookmarkContent());

			// save the docx...
			wordMLPackage.save(wordUpdateModel.getOutputFilePath());

		} catch (Docx4JException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	/**
	 * Method to replace Content Controls
	 * 
	 * @param documentPart
	 * @param wordContentMap
	 */
	private void replaceContentControls(MainDocumentPart documentPart, Map<String, String> wordContentMap) {
		wordContentMap.entrySet().forEach(entry -> {
			try {

				String xpath = "//w:sdt/w:sdtPr/w:id[@w:val='" + entry.getKey() + "']/../../w:sdtContent/w:r/w:t";
				List<Object> xPathMatchedObjectList = documentPart.getJAXBNodesViaXPath(xpath, false);

				boolean textReplaced = false;
				for(Object object : xPathMatchedObjectList) {
					System.out.println("Iterating the Content Control Objects");
					Object unwrappedObject = XmlUtils.unwrap(object);
					
					if(unwrappedObject instanceof Text) {
						Text textObject = (Text)unwrappedObject;
						if(!textReplaced) {
							System.out.println("Text is updated with value :: " + entry.getValue());
							textObject.setValue(entry.getValue());
							textReplaced = true;
						}
						else {
							textObject.setValue("");
						}
					}	
					
				}

			} catch (XPathBinderAssociationIsPartialException | JAXBException e) {
				e.printStackTrace();
			}
		});

	}
}
