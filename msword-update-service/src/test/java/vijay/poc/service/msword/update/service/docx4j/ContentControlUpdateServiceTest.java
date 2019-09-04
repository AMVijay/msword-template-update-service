package vijay.poc.service.msword.update.service.docx4j;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

import org.junit.jupiter.api.Test;

import vijay.poc.msword.update.service.WordUpdateModel;
import vijay.poc.msword.update.service.docx4j.ContentControlUpdateService;

public class ContentControlUpdateServiceTest {
	
	@Test
	public void test() {
		
		
		WordUpdateModel wordGenerateModel = new WordUpdateModel();
		
		String templateFolder = System.getProperty("user.dir") + "/test-templates/";
		String outputFolder =  System.getProperty("user.dir") + "/test-output/updated-";
		
		wordGenerateModel.setInputFilePath(new File(templateFolder + "template-placeholder.docx"));
		wordGenerateModel.setOutputFilePath(new File(outputFolder + wordGenerateModel.getInputFilePath().getName()));
		
        Map<String, String> wordContentMap = new HashMap<String, String>();

        wordContentMap.put("-433902062", "Mr.");
        wordContentMap.put("-1031495007", "Test User First Name");

        wordGenerateModel.setWordContent(wordContentMap);
        
        ContentControlUpdateService service = new ContentControlUpdateService();
        service.updateWordDocument(wordGenerateModel);
	}

}
