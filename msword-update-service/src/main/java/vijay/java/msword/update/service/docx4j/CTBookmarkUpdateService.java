package vijay.java.msword.update.service.docx4j;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.finders.RangeFinder;
import org.docx4j.jaxb.Context;
import org.docx4j.model.fields.merge.DataFieldName;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.CTBookmark;
import org.docx4j.wml.CTMarkupRange;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;

import vijay.java.msword.update.service.IWordUpdateService;
import vijay.java.msword.update.service.WordUpdateModel;

public class CTBookmarkUpdateService implements IWordUpdateService {

    private static boolean DELETE_BOOKMARK = false;

    private static org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();

    @Override
    public void updateWordDocument(WordUpdateModel wordGenerateModel) {

//        String filePath = "C:\\Users\\135670\\Vijay\\github\\java\\word-generation\\src\\main\\resources\\templates\\template-firstpage-bookmark.docx";
//        String outputfile = "C:\\Users\\135670\\Vijay\\github\\java\\word-generation\\src\\main\\resources\\template-firstpage-bookmark-updated.docx";

        try {

            Map<DataFieldName, String> map = new HashMap<DataFieldName, String>();
            wordGenerateModel.getWordContent().entrySet().forEach(entry -> {
                map.put(new DataFieldName(entry.getKey()), entry.getValue());
            });
            
            System.out.println("Service is loading the Word Document as WordML Package...");
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(wordGenerateModel.getInputFilePath());
            System.out.println("Service is loading the MainDocument Object...");
            MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

            org.docx4j.wml.Document wmlDocumentEl = (org.docx4j.wml.Document) documentPart.getJaxbElement();
            Body body = wmlDocumentEl.getBody();

            System.out.println("Service is replacing the bookmarks - Text Content...");
            replaceBookmarkContents(body.getContent(), map);
            System.out.println("Service completed Bookmark - Text Content");
            // save the docx...
            wordMLPackage.save(wordGenerateModel.getOutputFilePath());

        } catch (Docx4JException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {

        }
    }

    private void replaceBookmarkContents(List<Object> paragraphs, Map<DataFieldName, String> data) throws Exception {

        RangeFinder rt = new RangeFinder("CTBookmark", "CTMarkupRange");
        new TraversalUtil(paragraphs, rt);

        for (CTBookmark bm : rt.getStarts()) {

            // do we have data for this one?
            if (bm.getName() == null)
                continue;
            String value = data.get(new DataFieldName(bm.getName()));
            System.out.println("BOOKMARK name :: " + bm.getName() + " Text Content Content is going to be replaced...");
            if (value == null)
                continue;

            try {
                // Can't just remove the object from the parent,
                // since in the parent, it may be wrapped in a JAXBElement
                List<Object> theList = null;
                if (bm.getParent() instanceof P) {
//                    System.out.println("OK!");
                    theList = ((ContentAccessor) (bm.getParent())).getContent();
//                    System.out.println("theList size :: " + theList.size());
                    System.out.println("BOOKMARK name :: " + bm.getName() + " Paragraph Content is going to be changed...");
                    updateParagraphContent(bm,theList,value);
                } else {
                    continue;
                }
//
//                int rangeStart = -1;
//                int rangeEnd = -1;
//                int i = 0;
//                int textUpdatedIndex = -1;
//                for (Object ox : theList) {
//                    Object listEntry = XmlUtils.unwrap(ox);
//                    if (listEntry.equals(bm)) {
//                        System.out.println("Bookmark Matched");
//                        if (DELETE_BOOKMARK) {
//                            rangeStart = i;
//                        } else {
//                            rangeStart = i + 1;
//                        }
//                    } else if (listEntry instanceof CTMarkupRange) {
//                        if (((CTMarkupRange) listEntry).getId().equals(bm.getId())) {
//                            if (DELETE_BOOKMARK) {
//                                rangeEnd = i;
//                            } else {
//                                rangeEnd = i - 1;
//                            }
//                            break;
//                        }
//                    } else if (listEntry instanceof org.docx4j.wml.Text){
//                        System.out.println(" ListEntry is Text Type");
//
//                        org.docx4j.wml.Text t = (org.docx4j.wml.Text) listEntry;
//                        String text = t.getValue();
//
//                        if (text != null)
//                        {
//                           // t.setSpace("preserve"); // needed?
//                            t.setValue(value);
//                            textUpdatedIndex = i;
//                        }
//                    }                
//                    i++;
//                }
//
//                if(textUpdatedIndex > 0 ){
//                    System.out.println("textUpdatedIndex " + textUpdatedIndex);
//                     // Delete the bookmark range
//                     for (int j = 0; j <= theList.size(); j++) {
//                         if(j != textUpdatedIndex){
//                            theList.remove(j);
//                         }
//                    }
//                }
//                else if (rangeStart > 0 && rangeEnd <= rangeStart && theList.size() > rangeStart) {
//
//                    org.docx4j.wml.R run = factory.createR();
//                    org.docx4j.wml.Text t = factory.createText();
//                    run.getContent().add(t);
//                    t.setValue(value);
//                    theList.add(rangeStart, run);
//
//                    // Delete the bookmark range
//                    for (int j = rangeStart + 1; j <= theList.size(); j++) {
//                        theList.remove(j);
//                    }
//                } else if (rangeStart > 0 && rangeEnd > rangeStart) {
//
//                    // Delete the bookmark range
//                    for (int j = rangeEnd; j >= rangeStart; j--) {
//                        theList.remove(j);
//                    }
//
//                    // now add a run
//                    org.docx4j.wml.R run = factory.createR();
//                    org.docx4j.wml.Text t = factory.createText();
//                    run.getContent().add(t);
//                    t.setValue(value);
//                    System.out.println("Setting Value");
//                    theList.add(rangeStart, run);
//                }

            } catch (ClassCastException cce) {
                cce.printStackTrace();
            }
        }

    }

    private static final String BLANK_VALUE = "";

    private void updateParagraphContent(CTBookmark ctBookmark, List<Object> objectList, String textContent) {
    	
    	boolean bookmarkRangeStarted = false;
    	boolean textUpdated = false;
    	
    	for (Object object : objectList) {
			Object unwrappedObject = XmlUtils.unwrap(object);
			if (unwrappedObject instanceof CTBookmark) {
				CTBookmark bookmarkStart = (CTBookmark)unwrappedObject;
				if(bookmarkStart.getId() == ctBookmark.getId()) {
					System.out.println("BOOKMARK Name " + bookmarkStart.getName() + " ; BOOKMARK ID " + bookmarkStart.getId());
					bookmarkRangeStarted = true;
				}
			}
			else if(unwrappedObject instanceof CTMarkupRange && bookmarkRangeStarted) {
				CTMarkupRange bookmarkEnd = (CTMarkupRange)unwrappedObject;
				//System.out.println("BOOKMARK End :: " + bookmarkEnd.getId());
				if(ctBookmark.getId().intValue() == bookmarkEnd.getId().intValue()) {
					System.out.println("BOOKMARK Name " + ctBookmark.getName() + " text update completed");
					bookmarkRangeStarted = false;
				}
			}
			else if(bookmarkRangeStarted && unwrappedObject instanceof R) {
				System.out.println("Word Runtime Node matched");
				R runTime = (R)unwrappedObject;
				if(!textUpdated) {
					System.out.println("Replacing actual Text Content");
					updateRunObjectTextContent(runTime.getContent(),textContent);
					textUpdated = true;
				}
				else {
					System.out.println("As actual text content replaced, now replacing BLANK content");
					updateRunObjectTextContent(runTime.getContent(),BLANK_VALUE);
				}
			}
    	}
    }
    
    private void updateRunObjectTextContent(List<Object> objectList, String textContent) {
    	boolean textUpdateCompleted = false;
		for(Object object : objectList) {
			Object unwrappedObject = XmlUtils.unwrap(object);
			if(unwrappedObject instanceof Text) {
				Text text = (Text)unwrappedObject;
				if(!textUpdateCompleted) {
					text.setValue(textContent);
					textUpdateCompleted = true;
				}else {
					text.setValue(BLANK_VALUE);
				}
			}
		}
	}
}