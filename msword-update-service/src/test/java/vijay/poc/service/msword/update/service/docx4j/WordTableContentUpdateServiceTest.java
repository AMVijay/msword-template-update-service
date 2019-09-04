package vijay.poc.service.msword.update.service.docx4j;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

import org.junit.jupiter.api.Test;

import vijay.poc.msword.update.service.WordUpdateModel;
import vijay.poc.msword.update.service.docx4j.WordTableContentUpdateService;

public class WordTableContentUpdateServiceTest {

	@Test
	public void test() {

		String templateFolder = System.getProperty("user.dir") + "/test-templates/";
		String outputFolder = System.getProperty("user.dir") + "/test-output/updated-";

		WordUpdateModel wordUpdateModel = new WordUpdateModel();

		wordUpdateModel.setInputFilePath(new File(templateFolder + "template-bookmark.docx"));
		wordUpdateModel.setOutputFilePath(new File(outputFolder + wordUpdateModel.getInputFilePath().getName()));

		Map<String, String> wordContentMap = new HashMap<String, String>();
		wordContentMap.put("tablecontent", "tablecontent");		
		wordUpdateModel.setWordContent(wordContentMap);

		WordTableContentUpdateService service = new WordTableContentUpdateService();
		service.updateWordDocument(wordUpdateModel);

	}
}
