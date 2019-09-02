package vijay.java.msword.update.service.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import vijay.java.msword.update.service.IWordUpdateService;
import vijay.java.msword.update.service.WordUpdateModel;

public class POIBookmarkWordGenerateService implements IWordUpdateService {

    @Override
    public void updateWordDocument(WordUpdateModel wordGenerateModel) {

        String filePath = System.getProperty("user.dir") + "/test-templates/template-bookmark.docx";

        String outputfile = System.getProperty("user.dir") + "/test-output/template-bookmark-updated.docx";

        try {

            
            final XWPFDocument document = new XWPFDocument(new FileInputStream(filePath));

            wordGenerateModel.getWordContent().entrySet().forEach(entry -> {
                for (XWPFParagraph paragraph : document.getParagraphs()) {

                    for(CTBookmark bookmark: paragraph.getCTP().getBookmarkStartList()){

                        System.out.println("Bookmark Name :: " +  bookmark.getName());

                        NodeList nodeList = bookmark.getDomNode().getChildNodes();

                        for(int i=0;i<nodeList.getLength();i++){
                            Node node = nodeList.item(i);
                            System.out.println("Node Details" + node.getNodeName() + "-" + node.getNodeValue());
                            System.out.println("Node Text Content" + node.getTextContent());
                        }
                        

                    }

                    // for (XWPFRun run : paragraph.getRuns()) {
                    //     String text = run.text();
                    //     System.out.println("Read Text" + text);
                    //     if(text.contains(entry.getKey())) {
                    //         text = text.replace(entry.getKey(), entry.getValue());
                    //         run.setText(text, 0);
                    //         System.out.println(text);
                    //     }
                    // }
                }
            });

            document.write(new FileOutputStream(outputfile));
            document.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    

    

}