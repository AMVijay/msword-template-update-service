package vijay.poc.msword.update.service;

import java.io.File;
import java.util.List;
import java.util.Map;

import lombok.Data;

@Data
public class WordUpdateModel {

	private File inputFilePath;

	private File outputFilePath;

	private Map<String, String> bookmarkContent;

	private String tableBookmarkName;
	private List<List<String>> tableContentList;

}
