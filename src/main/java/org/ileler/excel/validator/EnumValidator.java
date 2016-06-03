package org.ileler.excel.validator;

import org.apache.commons.lang3.StringUtils;
import org.ileler.excel.util.LangCode;
import org.w3c.dom.Element;

import java.util.Map;

/**
 * Created by ileler@qq.com on 2016/5/12.
 */
public class EnumValidator extends Validator {
	
	private static final String ENUMVALUE = "evalue";
	
	private static final String ENUMSPLIT = ",";
	
	private String[] enuVal;

	public EnumValidator(String name, String emsg, Element element) {
		super(name, emsg, element);
		if (element != null) {
			String evalue = element.getAttribute(ENUMVALUE);
			enuVal = StringUtils.isEmpty(evalue) ? null : evalue.split(ENUMSPLIT);
		}
		if (enuVal == null || enuVal.length < 1) throw new IllegalArgumentException(ENUMVALUE + " is null.");
	}

	public String[] getEnuVal() {
		return enuVal;
	}

	public void setEnuVal(String[] enuVal) {
		this.enuVal = enuVal;
	}

	@Override
	public boolean check(Object value, LangCode langCode, Map<String, Object> session) {
		String[] enuVal = null;
		if (value == null || StringUtils.isEmpty((String)value) || StringUtils.isEmpty(((String)value).trim()) || (enuVal = getEnuVal()) == null || enuVal.length < 1) 	return true;
		for (int i = 0, j = enuVal.length; i < j; i++) {
			if (value.equals(enuVal[i])) 	return true;
		}
		return false;
	}

}