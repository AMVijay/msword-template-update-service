package vijay.java.service.msword.update.service.docx4j;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import vijay.java.msword.update.service.WordUpdateModel;
import vijay.java.msword.update.service.docx4j.CTBookmarkUpdateService;

public class CTBookmarkUpdateServiceTest {

	@Test
	public void testUpdateWordDocument() {
		
		String templateFolder = System.getProperty("user.dir") + "/test-templates/";
		String outputFolder =  System.getProperty("user.dir") + "/test-output/updated-";

		WordUpdateModel wordGenerateModel = new WordUpdateModel();
//		wordGenerateModel.setInputFilePath(new File(templateFolder + "template-firstpage-bookmark.docx"));
		wordGenerateModel.setInputFilePath(new File(templateFolder + "template-bookmark.docx"));
		wordGenerateModel.setOutputFilePath(new File(outputFolder + wordGenerateModel.getInputFilePath().getName()));

		Map<String, String> wordContentMap = new HashMap<String, String>();

		wordContentMap.put("salutation", "Mr.");
		wordContentMap.put("firstName", "Test User First Name");
		wordContentMap.put("MiddleName", " ");
		wordContentMap.put("LastName", "Test Last Name");
		wordContentMap.put("Suffix", " ");
		wordContentMap.put("columnname", "record1");
		wordContentMap.put("columndescription", "record description");

		wordGenerateModel.setWordContent(wordContentMap);

		CTBookmarkUpdateService service = new CTBookmarkUpdateService();
		service.updateWordDocument(wordGenerateModel);

		Assertions.assertTrue(true);

	}

}
