package org.ileler.excel.validator;

import org.apache.commons.lang3.StringUtils;
import org.ileler.excel.util.LangCode;
import org.w3c.dom.Element;

import java.util.Map;

public class EmptyValidator extends Validator {
	
	public EmptyValidator(String name, String emsg, Element element) {
		super(name, emsg, element);
	}

	@Override
	public boolean check(Object value, LangCode langCode, Map<String, Object> session) {
		if (value == null || StringUtils.isEmpty((String)value) || StringUtils.isEmpty(((String)value).trim())) 	return false;
		return true;
	}

}