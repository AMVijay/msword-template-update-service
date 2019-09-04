package vijay.poc.msword.update.service.regex;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.Writer;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import vijay.poc.msword.update.service.IWordUpdateService;
import vijay.poc.msword.update.service.WordUpdateModel;

public class RegularExpressionWordGenerateService implements IWordUpdateService {

	@Override
	public void updateWordDocument(WordUpdateModel wordGenerateModel) {

		String filePath = System.getProperty("user.dir") + "/test-templates/template.xml";

		String outputfile = System.getProperty("user.dir") + "/test-ouput/template-updated.xml";

		List<String> regularExpressionPatternList = new ArrayList<String>();
		regularExpressionPatternList.add("(\\{.*?\\})");
		// regularExpressionPatternList.add("(\\{).*?<w:t>(.*?)</w:t>.*(\\})");

		try {

			BufferedReader reader = new BufferedReader(new FileReader(filePath));
			Writer writer = new FileWriter(outputfile);

			String contentLine = reader.readLine();

			while (contentLine != null) {

				for (String patternString : regularExpressionPatternList) {
					contentLine = checkContentAndReplaceUsingPattern(contentLine, patternString, wordGenerateModel.getWordContent());
				}
				System.out.println("SEcond Time Check");
				// checkContentAndReplaceUsingPattern(contentLine,
				// regularExpressionPatternList.get(0), wordContentMap);
				writer.write(contentLine + "\n");
				contentLine = reader.readLine();
			}

			reader.close();
			writer.close();

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private String checkContentAndReplaceUsingPattern(String contentLine, String patternString, Map<String, String> wordContentMap) {

		// System.out.println("Content Line :: " + contentLine);
		System.out.println("Pattern :: " + patternString);

		Pattern pattern = Pattern.compile(patternString);
		// Now create matcher object.
		Matcher matcher = pattern.matcher(contentLine);
		while (matcher.find()) {
			if (!matcher.group().isEmpty()) {
				int groupCount = matcher.groupCount();
				StringBuffer keyName = new StringBuffer();
				for (int i = 1; i <= groupCount; i++) {
					keyName.append(matcher.group(i));
				}

				System.out.println("Key Name :: " + keyName);
				if (wordContentMap.get(keyName.toString()) != null && groupCount == 1) {
					System.out.println("KeyValue" + wordContentMap.get(keyName.toString()));
					System.out.println("To Be Replaced String " + matcher.group());
					contentLine = contentLine.replaceFirst(patternString, wordContentMap.get(keyName.toString()));
				}
			}
		}

		return contentLine;
	}

}