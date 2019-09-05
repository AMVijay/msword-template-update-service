package vijay.poc.service.msword.update.service.docx4j;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
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

		List<List<String>> tableContentList = new ArrayList<List<String>>();
		for (int i = 0; i < 20; i++) {
			List<String> tableRow = new ArrayList<String>();
			for (int j = 0; j < 5; j++) {
				tableRow.add("content " + i + " " + j);
			}
			tableContentList.add(tableRow);
		}

		wordUpdateModel.setTableContentList(tableContentList);
		WordTableContentUpdateService service = new WordTableContentUpdateService();
		service.updateWordDocument(wordUpdateModel);

	}
}
