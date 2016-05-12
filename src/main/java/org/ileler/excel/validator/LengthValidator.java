package org.ileler.excel.validator;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.ileler.excel.util.LangCode;
import org.w3c.dom.Element;

import java.util.Map;

public class LengthValidator extends Validator {
	
	private static final String GT = "gt";
	
	private static final String LT = "lt";
	
	private static final String EQ = "eq";
	
	private Integer gt;

	private Integer lt;
	
	private Integer eq;
	
	public LengthValidator(String name, String emsg, Element element) {
		super(name, emsg, element);
		if (element != null) {
		    String tmp = null;
			gt = (StringUtils.isBlank(tmp = element.getAttribute(GT)) || !NumberUtils.isDigits(tmp)) ? null : Integer.valueOf(tmp);
			lt = (StringUtils.isBlank(tmp = element.getAttribute(LT)) || !NumberUtils.isDigits(tmp)) ? null : Integer.valueOf(tmp);
			eq = (StringUtils.isBlank(tmp = element.getAttribute(EQ)) || !NumberUtils.isDigits(tmp)) ? null : Integer.valueOf(tmp);
		}
	}

	public Integer getGt() {
		return gt;
	}

	public void setGt(Integer gt) {
		this.gt = gt;
	}

	public Integer getLt() {
		return lt;
	}

	public void setLt(Integer lt) {
		this.lt = lt;
	}

	public Integer getEq() {
        return eq;
    }

    public void setEq(Integer eq) {
        this.eq = eq;
    }

	@Override
	public boolean check(Object arg, LangCode langCode, Map<String, Object> session) {
		if (arg == null || (gt == null && lt == null && eq == null)) 	return true;
		int length = arg.toString().length();
		if (gt != null && length < gt) return false;
		if (lt != null && length > lt) return false;
		if (eq != null && length != eq)return false;
		return true;
	}
	
}