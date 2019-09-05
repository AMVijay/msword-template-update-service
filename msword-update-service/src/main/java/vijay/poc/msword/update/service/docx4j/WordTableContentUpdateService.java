package vijay.poc.msword.update.service.docx4j;

import java.util.ArrayList;
import java.util.List;

import javax.xml.bind.JAXBElement;

import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.ContentAccessor;
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
			updateTableContents(documentPart, wordUpdateModel.getTableContentList());
//			updateTableContents(body.getContent(), wordUpdateModel.getWordContent());

			System.out.println("Service completed Bookmark - Text Content");

			// save the docx...
			wordMLPackage.save(wordUpdateModel.getOutputFilePath());

		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private void updateTableContents(MainDocumentPart documentPart, List<List<String>> tableContentList) {
		List<Object> tableList = getAllElementFromObject(documentPart, Tbl.class);
		for (Object object : tableList) {
			Object unwrappedObject = XmlUtils.unwrap(object);
			if (unwrappedObject instanceof Tbl) {
				Tbl table = (Tbl) unwrappedObject;
				System.out.println("Table Object Matched");
				updateTableRowContent(table, tableContentList);
			}
		}
	}

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
						Object unwrappedObject = XmlUtils.unwrap(table.getContent().get(table.getContent().size()-1));
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
	 * @param objectList
	 * @param textContent
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

	private static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
		List<Object> result = new ArrayList<Object>();
		if (obj instanceof JAXBElement)
			obj = ((JAXBElement<?>) obj).getValue();

		if (obj.getClass().equals(toSearch))
			result.add(obj);
		else if (obj instanceof ContentAccessor) {
			List<?> children = ((ContentAccessor) obj).getContent();
			for (Object child : children) {
				result.addAll(getAllElementFromObject(child, toSearch));
			}

		}
		return result;
	}

}
