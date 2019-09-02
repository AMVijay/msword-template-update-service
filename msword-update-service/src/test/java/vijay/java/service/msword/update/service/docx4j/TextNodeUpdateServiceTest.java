package vijay.java.service.msword.update.service.docx4j;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import vijay.java.msword.update.service.WordUpdateModel;
import vijay.java.msword.update.service.docx4j.TextNodeUpdateService;

/**
 * Unit test for simple App.
 */
public class TextNodeUpdateServiceTest {

    @Test
    public void testGenerateWordDocumentService(){

    	WordUpdateModel wordUpdateModel = new WordUpdateModel();
    	
    	String templateFolder = System.getProperty("user.dir") + "/test-templates/";
		String outputFolder =  System.getProperty("user.dir") + "/test-output/updated-";
		
		wordUpdateModel.setInputFilePath(new File(templateFolder + "template-placeholder.docx"));
		wordUpdateModel.setOutputFilePath(new File(outputFolder + wordUpdateModel.getInputFilePath().getName()));
        
        Map<String, String> wordContentMap = new HashMap<String, String>();

        wordContentMap.put("Salutation", "Mr.");
        wordContentMap.put("FirstName", "Test User First Name");
        wordContentMap.put("MiddleName", " ");
        wordContentMap.put("LastName", "Test Last Name");
        wordContentMap.put("Suffix"," ");
        wordContentMap.put("columnname", "record1");
        wordContentMap.put("columndescription", "record description");
        
        wordUpdateModel.setWordContent(wordContentMap);

        TextNodeUpdateService service = new TextNodeUpdateService();
        service.updateWordDocument(wordUpdateModel);

        Assertions.assertTrue(true);
    }

}
