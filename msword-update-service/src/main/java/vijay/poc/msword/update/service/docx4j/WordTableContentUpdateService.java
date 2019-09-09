package vijay.poc.msword.update.service.docx4j;

import java.util.ArrayList;
import java.util.List;

import javax.xml.bind.JAXBException;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;

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
	 * @param wordUpdateModel as {@link WordUpdateModel}
	 * @throws Docx4JException.
	 */
	@Override
	public void updateWordDocument(WordUpdateModel wordUpdateModel) {

		try {

			System.out.println("Service is creating WordML from docx...");
			WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(wordUpdateModel.getInputFilePath());

			System.out.println("Service is loading the MainDocument Object from WordML...");
			MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

			System.out.println("Service is replacing the Table Content...");
			// updateTableContents(documentPart, wordUpdateModel.getTableContentList());

			updateTableElementWithinBookmark(documentPart, wordUpdateModel.getTableBookmarkName(), wordUpdateModel.getTableContentList());

			System.out.println("Service completed Bookmark - Text Content");

			// save the docx...
			wordMLPackage.save(wordUpdateModel.getOutputFilePath());

		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * Method to update WordML Table Content after the bookmark.
	 * 
	 * @param documentPart     as {@link MainDocumentPart}
	 * @param bookmarkName     as {@link String}
	 * @param tableContentList as {@link List} of {@link List} of {@link String}.
	 */
	private void updateTableElementWithinBookmark(MainDocumentPart documentPart, String bookmarkName, List<List<String>> tableContentList) {

		System.out.println("findTableElementWithinBookmark start");

		String xpath = "//w:bookmarkStart[@w:name=\"" + bookmarkName + "\"]/../following-sibling::*[1][self::w:tbl]";
		try {
			List<Object> xPathMatchedObjectList = documentPart.getJAXBNodesViaXPath(xpath, false);
			xPathMatchedObjectList.forEach(object -> {
				System.out.println("");
				Object unwrappedObject = XmlUtils.unwrap(object);
				if (unwrappedObject instanceof Tbl) {
					Tbl table = (Tbl) unwrappedObject;
					System.out.println("Table Object Matched");
					updateTableRowContent(table, tableContentList);
				}
			});
		} catch (XPathBinderAssociationIsPartialException | JAXBException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * Method to update WordML Table Row Content.
	 * 
	 * @param table            as {@link Tbl}.
	 * @param tableContentList as {@link List} of {@link List} of {@link String}.
	 */
	private void updateTableRowContent(Tbl table, List<List<String>> tableContentList) {
		if (table.getContent() != null && !table.getContent().isEmpty()) {
			List<Tr> newTableRows = new ArrayList<Tr>();
			for (int i = 0; i < tableContentList.size(); i++) {
				if (i < table.getContent().size()) {
					Object unwrappedObject = XmlUtils.unwrap(table.getContent().get(i));
					if (unwrappedObject instanceof Tr) {
						Tr tableRow = (Tr) unwrappedObject;
						updateTableColumn(tableRow, tableContentList.get(i));
					}
				} else {
					if (table.getContent().get(table.getContent().size() - 1) != null) {
						Object unwrappedObject = XmlUtils.unwrap(table.getContent().get(table.getContent().size() - 1));
						if (unwrappedObject instanceof Tr) {
							Tr workingRow = (Tr) XmlUtils.deepCopy((Tr) unwrappedObject);
							updateTableColumn(workingRow, tableContentList.get(i));
							newTableRows.add(workingRow);
						}
					}
				}
			}
			table.getContent().addAll(newTableRows);
		}
	}

	/**
	 * Method to update WordML Table Column.
	 * 
	 * @param tableRow            as {@link Tr}
	 * @param tableRowContentList as {@link List} of {@link String}.
	 */
	private void updateTableColumn(Tr tableRow, List<String> tableRowContentList) {
		if (tableRow.getContent() != null && !tableRow.getContent().isEmpty()) {
			int i = 0;
			for (Object object : tableRow.getContent()) {
				Object unwrappedObject = XmlUtils.unwrap(object);
				if (unwrappedObject instanceof Tc) {
					Tc tableColumn = (Tc) unwrappedObject;
					updateParagraphContent(tableColumn.getContent(), tableRowContentList.get(i));
					i++;
				}
			}
		}
	}

	/**
	 * Method to update Word Text Value in Paragraph.
	 * 
	 * @param objectList    as {@link List} of {@link Object}.
	 * @param columnContent as {@link String}.
	 */
	private void updateParagraphContent(List<Object> objectList, String columnContent) {

		for (Object object : objectList) {
			Object unwrappedObject = XmlUtils.unwrap(object);
			if (unwrappedObject instanceof P) {
				P p = (P) unwrappedObject;
				updateRunObjectTextContent(p.getContent(), columnContent);
			}
		}
	}

	private static final String BLANK_VALUE = "";

	/**
	 * Method to update Text Element.
	 * 
	 * @param objectList  as {@link List} of {@link Object}.
	 * @param textContent as {@link String}
	 */
	private void updateRunObjectTextContent(List<Object> objectList, String textContent) {
		boolean textUpdateCompleted = false;
		for (Object object : objectList) {
			Object unwrappedObject = XmlUtils.unwrap(object);
			if (unwrappedObject instanceof R) {
				R r = (R) unwrappedObject;
				for (Object object1 : r.getContent()) {
					Object unwrappedObject1 = XmlUtils.unwrap(object1);
					if (unwrappedObject1 instanceof Text) {
						Text text = (Text) unwrappedObject1;
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
	}

}
