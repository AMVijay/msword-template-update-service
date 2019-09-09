package vijay.poc.msword.update.service;

import java.util.ArrayList;
import java.util.List;

import javax.xml.bind.JAXBElement;

import org.docx4j.wml.ContentAccessor;

public class CommonService {

	public static List<Object> getAllElementFromObject(Object object, Class<?> toSearch) {
		List<Object> result = new ArrayList<Object>();
		if (object instanceof JAXBElement)
			object = ((JAXBElement<?>) object).getValue();

		if (object.getClass().equals(toSearch)) {
			result.add(object);
		} else if (object instanceof ContentAccessor) {
			List<?> children = ((ContentAccessor) object).getContent();
			for (Object child : children) {
				result.addAll(getAllElementFromObject(child, toSearch));
			}

		}
		return result;
	}

}
