package vijay.poc.msword.update.service.docx4j;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.finders.RangeFinder;
import org.docx4j.model.fields.merge.DataFieldName;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.CTBookmark;
import org.docx4j.wml.CTMarkupRange;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;

import vijay.poc.msword.update.service.IWordUpdateService;
import vijay.poc.msword.update.service.WordUpdateModel;

/**
 * CTBookmarkUpdateService - class to update bookmark text content.
 * @author Vijay
 *
 */
public class CTBookmarkUpdateService implements IWordUpdateService {

	/**
	 * Method to update bookmark text content.
	 * @param wordUpdateModel as WordUpdateModel.
	 */
	@Override
	public void updateWordDocument(WordUpdateModel wordUpdateModel) {

		try {

			Map<DataFieldName, String> map = new HashMap<DataFieldName, String>();
			wordUpdateModel.getWordContent().entrySet().forEach(entry -> {
				map.put(new DataFieldName(entry.getKey()), entry.getValue());
			});

			System.out.println("Service is loading the Word Document as WordML Package...");
			WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(wordUpdateModel.getInputFilePath());
			System.out.println("Service is loading the MainDocument Object...");
			MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

			org.docx4j.wml.Document wmlDocumentEl = (org.docx4j.wml.Document) documentPart.getJaxbElement();
			Body body = wmlDocumentEl.getBody();

			System.out.println("Service is replacing the bookmarks - Text Content...");
			replaceBookmarkContents(body.getContent(), map);
			System.out.println("Service completed Bookmark - Text Content");
			// save the docx...
			wordMLPackage.save(wordUpdateModel.getOutputFilePath());

		} catch (Docx4JException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {

		}
	}

	/**
	 * Method to update bookmark.
	 * @param paragraphs
	 * @param data
	 * @throws Exception
	 */
	private void replaceBookmarkContents(List<Object> paragraphs, Map<DataFieldName, String> data) throws Exception {

		RangeFinder rt = new RangeFinder("CTBookmark", "CTMarkupRange");
		new TraversalUtil(paragraphs, rt);

		for (CTBookmark bm : rt.getStarts()) {

			// do we have data for this one?
			if (bm.getName() == null)
				continue;
			String value = data.get(new DataFieldName(bm.getName()));
			System.out.println("BOOKMARK name :: " + bm.getName() + " Text Content Content is going to be replaced...");
			if (value == null)
				continue;

			try {
				// Can't just remove the object from the parent,
				// since in the parent, it may be wrapped in a JAXBElement
				List<Object> theList = null;
				if (bm.getParent() instanceof P) {
//                    System.out.println("OK!");
					theList = ((ContentAccessor) (bm.getParent())).getContent();
//                    System.out.println("theList size :: " + theList.size());
					System.out.println("BOOKMARK name :: " + bm.getName() + " Paragraph Content is going to be changed...");
					updateParagraphContent(bm, theList, value);
				} else {
					continue;
				}

			} catch (ClassCastException cce) {
				cce.printStackTrace();
			}
		}

	}

	private static final String BLANK_VALUE = "";

	/**
	 * Method to find text element withing bookmark.
	 * @param ctBookmark
	 * @param objectList
	 * @param textContent
	 */
	private void updateParagraphContent(CTBookmark ctBookmark, List<Object> objectList, String textContent) {

		boolean bookmarkRangeStarted = false;
		boolean textUpdated = false;

		for (Object object : objectList) {
			Object unwrappedObject = XmlUtils.unwrap(object);
			if (unwrappedObject instanceof CTBookmark) {
				CTBookmark bookmarkStart = (CTBookmark) unwrappedObject;
				if (bookmarkStart.getId() == ctBookmark.getId()) {
					System.out.println("BOOKMARK Name " + bookmarkStart.getName() + " ; BOOKMARK ID " + bookmarkStart.getId());
					bookmarkRangeStarted = true;
				}
			} else if (unwrappedObject instanceof CTMarkupRange && bookmarkRangeStarted) {
				CTMarkupRange bookmarkEnd = (CTMarkupRange) unwrappedObject;
				// System.out.println("BOOKMARK End :: " + bookmarkEnd.getId());
				if (ctBookmark.getId().intValue() == bookmarkEnd.getId().intValue()) {
					System.out.println("BOOKMARK Name " + ctBookmark.getName() + " text update completed");
					bookmarkRangeStarted = false;
				}
			} else if (bookmarkRangeStarted && unwrappedObject instanceof R) {
				System.out.println("Word Runtime Node matched");
				R runTime = (R) unwrappedObject;
				if (!textUpdated) {
					System.out.println("Replacing actual Text Content");
					updateRunObjectTextContent(runTime.getContent(), textContent);
					textUpdated = true;
				} else {
					System.out.println("As actual text content replaced, now replacing BLANK content");
					updateRunObjectTextContent(runTime.getContent(), BLANK_VALUE);
				}
			}
		}
	}

	/**
	 * Method to update Text Element.
	 * @param objectList
	 * @param textContent
	 */
	private void updateRunObjectTextContent(List<Object> objectList, String textContent) {
		boolean textUpdateCompleted = false;
		for (Object object : objectList) {
			Object unwrappedObject = XmlUtils.unwrap(object);
			if (unwrappedObject instanceof Text) {
				Text text = (Text) unwrappedObject;
				if (!textUpdateCompleted) {
					text.setValue(textContent);
					textUpdateCompleted = true;
				} else {
					text.setValue(BLANK_VALUE);
				}
			}
		}
	}
}