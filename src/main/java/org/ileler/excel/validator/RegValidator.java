package org.ileler.excel.validator;

import org.apache.commons.lang3.StringUtils;
import org.ileler.excel.util.LangCode;
import org.w3c.dom.Element;

import java.util.Map;
import java.util.regex.Pattern;

public class RegValidator extends Validator {
	
	private static final String PATTERN = "pattern";
	
	private String pattern;
	
	public RegValidator(String name, String emsg, Element element) {
		super(name, emsg, element);
		if (element != null) {
		    pattern = element.getAttribute(PATTERN);
		}
	}

	public String getPattern() {
        return pattern;
    }

    public void setPattern(String pattern) {
        this.pattern = pattern;
    }

    @Override
	public boolean check(Object arg, LangCode langCode, Map<String, Object> session) {
		if (arg == null || StringUtils.isBlank(arg.toString()) || StringUtils.isBlank(pattern)) 	return true;
		return Pattern.matches(pattern, arg.toString());
	}
    
//    public static void main(String[] args) {
//        System.out.println(Pattern.matches("^[^\\s].{0,99}$", "asfdasdfasde"));
//    }
	
}