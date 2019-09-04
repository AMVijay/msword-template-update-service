package vijay.poc.msword.update.service.docx4j;

import java.util.List;
import java.util.Map;

import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.dml.CTTable;
import org.docx4j.finders.RangeFinder;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.CTBookmark;
import org.docx4j.wml.CTMarkupRange;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.docx4j.wml.R;

import vijay.poc.msword.update.service.IWordUpdateService;
import vijay.poc.msword.update.service.WordUpdateModel;

/**
 * TableContentUpdateService - To update Table Content in Word Template.
 * 
 * @author Vijay
 *
 */
public class WordTableContentUpdateService implements IWordUpdateService {

	/**
	 * Method to update Word Template Table Content.
	 * 
	 * @param word
	 * @throws Docx4JException
	 */
	@Override
	public void updateWordDocument(WordUpdateModel wordUpdateModel) {

		try {

			System.out.println("Service is creating WordML from docx...");
			WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(wordUpdateModel.getInputFilePath());

			System.out.println("Service is loading the MainDocument Object from WordML...");
			MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

			org.docx4j.wml.Document wmlDocumentEl = (org.docx4j.wml.Document) documentPart.getJaxbElement();
			Body body = wmlDocumentEl.getBody();

			System.out.println("Service is replacing the Table Content...");
			updateTableContents(body.getContent(), wordUpdateModel.getWordContent());

			System.out.println("Service completed Bookmark - Text Content");

			// save the docx...
			wordMLPackage.save(wordUpdateModel.getOutputFilePath());

		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * Method to update Table Contents.
	 * 
	 * @param wordObjects as List of Objects.
	 * @param wordContent
	 */
	private void updateTableContents(List<Object> wordObjects, Map<String, String> wordContent) {

		RangeFinder rangeFinder = new RangeFinder("CTBookmark", "CTMarkupRange");
		new TraversalUtil(wordObjects, rangeFinder);

		for (CTBookmark bookmark : rangeFinder.getStarts()) {
			if (wordContent.get(bookmark.getName()) != null) {
				System.out.println("Bookmark Name :: " + bookmark.getName() + " ;Bookmark ID :: " + bookmark.getId());
				if (bookmark.getParent() instanceof P) {

					List<Object> objectList = ((ContentAccessor) (bookmark.getParent())).getContent();
					updateParagraphContent(objectList, bookmark);
				}
			}
		}
	}

	/**
	 * Update Paragraph Content.
	 * 
	 * @param objectList as List of Object.
	 * @param bookmark   as CTBookmark.
	 */
	private void updateParagraphContent(List<Object> objectList, CTBookmark bookmark) {

		boolean bookmarkRangeStarted = false;
		boolean textUpdated = false;
		
		for (Object object : objectList) {
			Object unwrappedObject = XmlUtils.unwrap(object);
			if (unwrappedObject instanceof CTBookmark) {
				CTBookmark bookmarkStart = (CTBookmark) unwrappedObject;
				if (bookmarkStart.getId() == bookmark.getId()) {
					System.out.println("BOOKMARK Name " + bookmarkStart.getName() + " ; BOOKMARK ID " + bookmarkStart.getId());
					bookmarkRangeStarted = true;
				}
			} else if (unwrappedObject instanceof CTMarkupRange && bookmarkRangeStarted) {
				CTMarkupRange bookmarkEnd = (CTMarkupRange) unwrappedObject;
				// System.out.println("BOOKMARK End :: " + bookmarkEnd.getId());
				if (bookmark.getId().intValue() == bookmarkEnd.getId().intValue()) {
					System.out.println("BOOKMARK Name " + bookmark.getName() + " text update completed");
					bookmarkRangeStarted = false;
				}
			} else if (bookmarkRangeStarted && unwrappedObject instanceof R) {
				System.out.println("Word Runtime Node matched");
				R runTime = (R) unwrappedObject;
				if (!textUpdated) {
					System.out.println("Replacing actual Text Content");
					updateRunObjectTableContent(runTime.getContent());
					textUpdated = true;
				} 
			}
		}
	}

	
	private void updateRunObjectTableContent(List<Object> objectList) {
		
		for (Object object : objectList) {
			Object unwrappedObject = XmlUtils.unwrap(object);
		}
		
	}

}
