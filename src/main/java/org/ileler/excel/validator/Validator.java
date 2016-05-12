package org.ileler.excel.validator;

import org.ileler.excel.util.LangCode;
import org.w3c.dom.Element;

import java.util.Map;

public abstract class Validator implements Cloneable {

	private String name;
	
	private String emsg;
	
	private Element element;
	
	public Validator(String name, String emsg, Element element) {
		super();
		this.name = name;
		this.emsg = emsg;
		this.element = element;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getEmsg() {
		return emsg;
	}

	public void setEmsg(String emsg) {
		this.emsg = emsg;
	}
	
	public Element getElement() {
		return element;
	}

	public void setElement(Element element) {
		this.element = element;
	}
	
	@Override
	public String toString() {
		return "{name:" + (name == null ? "" : name) + ",emsg:" + (emsg == null ? "" : emsg) + "}";
	}
	
	@Override
	public Validator clone() {
	    try {
            return (Validator) super.clone();
        } catch (CloneNotSupportedException e) {
            e.printStackTrace();
        }
	    return null;
	}

	
	public boolean check(Object value, LangCode langCode) {
	    return check(value, langCode, null);
	}
	
	/**
     * 规则校验具体实现方法
     * @param value     待校验值
     * @return
     */
	public abstract boolean check(Object value, LangCode langCode, Map<String, Object> session);
	
}
