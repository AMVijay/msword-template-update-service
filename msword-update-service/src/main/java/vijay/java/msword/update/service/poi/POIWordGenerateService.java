package vijay.java.msword.update.service.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class POIWordGenerateService {

	public void generateWordDocumentAsDocx(Map<String, String> wordContent) {

		String filePath = System.getProperty("user.dir") + "/test-templates/template.docx";

		String outputfile = System.getProperty("user.dir") + "/test-output/template-updated.docx";

		try {

			final XWPFDocument document = new XWPFDocument(new FileInputStream(filePath));

			wordContent.entrySet().forEach(entry -> {
				for (XWPFParagraph paragraph : document.getParagraphs()) {
					for (XWPFRun run : paragraph.getRuns()) {
						String text = run.text();
						System.out.println("Read Text" + text);
						if (text.contains(entry.getKey())) {
							text = text.replace(entry.getKey(), entry.getValue());
							run.setText(text, 0);
							System.out.println(text);
						}
					}
				}
			});

			document.write(new FileOutputStream(outputfile));
			document.close();

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}