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
import org.docx4j.wml.CTBookmark;
import org.docx4j.wml.P;

import vijay.poc.msword.update.service.IWordUpdateService;
import vijay.poc.msword.update.service.WordUpdateModel;

public class BookmarkUpdateService implements IWordUpdateService {

	@Override
	public void updateWordDocument(WordUpdateModel wordGenerateModel) {

		try {

			Map<DataFieldName, String> map = new HashMap<DataFieldName, String>();
			wordGenerateModel.getWordContent().entrySet().forEach(entry -> {
				map.put(new DataFieldName(entry.getKey()), entry.getValue());
			});

			WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(wordGenerateModel.getInputFilePath());
			MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

			replaceBookmarkContents(documentPart, wordGenerateModel.getWordContent());

			// save the docx...
			wordMLPackage.save(wordGenerateModel.getOutputFilePath());

		} catch (Docx4JException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private void replaceBookmarkContents(MainDocumentPart documentPart, Map<String, String> map) {

		map.entrySet().forEach(entry -> {
			try {

				String xpath = "//w:bookmarkStart[@w:name=\"" + entry.getKey() + "\"]";
				List<Object> xPathMatchedObjectList = documentPart.getJAXBNodesViaXPath(xpath, false);
				System.out.println("xPath with Name :: " + xpath);
				xPathMatchedObjectList.forEach(object -> {
					Object unwrappedObject = XmlUtils.unwrap(object);
					if (unwrappedObject instanceof CTBookmark) {
						CTBookmark ctBookmark = (CTBookmark) (unwrappedObject);
						System.out.println("Bookmark ID :: " + ctBookmark.getId());

//						 if (ctBookmark.getParent() instanceof P) {
//		                    List<Object> paragraphObjectList = ((ContentAccessor) (ctBookmark.getParent())).getContent();
//		                    System.out.println("paragraphObjectList size :: " + paragraphObjectList.size());
//		                    
//		                    // Update Text Content in W:t field
//		                    updateTextField(entry.getValue(),ctBookmark, paragraphObjectList);
//						 }

						updateBookmarkTextFields(ctBookmark, documentPart, entry.getValue());

					}
				});

			} catch (XPathBinderAssociationIsPartialException | JAXBException e) {
				e.printStackTrace();
			}
		});

	}

	private void updateBookmarkTextFields(CTBookmark ctBookmark, MainDocumentPart documentPart, String value) {

		try {

			// String xpath = "//w:t[../../w:bookmarkStart[@w:id=\"" + ctBookmark.getId() +
			// "\"] and //w:t[../../w:bookmarkEnd[@w:id=\"" + ctBookmark.getId() + "\"]]";
//			String xpath = "//w:bookmarkStart[@w:id='" + ctBookmark.getId() + "']/../w:r/w:t and //w:bookmarkEnd[@w:id='" + ctBookmark.getId() + "']";
			String xpath = "//w:bookmarkStart[@w:id='" + ctBookmark.getId() + "']/..";
			System.out.println("xPath with ID :: " + xpath);
			List<Object> xPathMatchedObjectList = documentPart.getJAXBNodesViaXPath(xpath, false);

			System.out.println("xPath List :: " + xPathMatchedObjectList.size());
			boolean textReplaced = false;
			for (Object object : xPathMatchedObjectList) {

				Object unwrappedObject = XmlUtils.unwrap(object);
				if (unwrappedObject instanceof P) {
					P element = (P) unwrappedObject;
					System.out.println("Element Name " + element.getParaId());
				}

//				if (unwrappedObject instanceof Text) {
//					Text text = (Text) unwrappedObject;
//					if (!textReplaced) {
//						text.setValue(value);
//						textReplaced = true;
//					} else {
//						text.setValue("");
//					}
//				}
			}
		} catch (XPathBinderAssociationIsPartialException | JAXBException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

//	/**
//	 * Method to traverse the w:paragraph and identify text nodes within bookmark and update text.
//	 * @param replaceContent
//	 * @param ctBookmark
//	 * @param theList
//	 */
//	private void updateTextField(String replaceContent, CTBookmark ctBookmark, List<Object> theList) {
//		
//		boolean replaceStartFlag = false;
//		for(Object object : theList) {
//			Object unwrappedObject = XmlUtils.unwrap(object);
//			if(unwrappedObject instanceof CTBookmark && ((CTBookmark)unwrappedObject).getId() == ctBookmark.getId()) {
//				replaceStartFlag = true;
//				System.out.println("Matched the bookmark");
//			}
//			
//			if(unwrappedObject instanceof R) {
//				R rObject = (R)unwrappedObject;
//				List<Object> rContentList = rObject.getContent();
//				System.out.println("rContentList size" + rContentList.size());
//				if(replaceStartFlag)
//					updateWTextField(replaceContent,rContentList);
//			}
//		}
//		
//	}
//
//	
//	private void updateWTextField(String replaceContent, List<Object> objectList) {
//		for(Object object : objectList) {
//			Object unwrappedObject = XmlUtils.unwrap(object);
//			if(unwrappedObject instanceof Text) {
//				System.out.println("Text Value :: " + ((Text)unwrappedObject).getValue());
//			}
//		}
//		
//	}
}
