package org.ileler.excel.validator;

import org.apache.commons.lang3.StringUtils;
import org.ileler.excel.util.LangCode;
import org.w3c.dom.Element;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

/**
 * Created by ileler@qq.com on 2016/5/12.
 */
public class DateValidator extends Validator {
	
	private static final String DATEPATTERN = "yyyy-MM-dd HH:mm:ss";
	
	private static final String PATTERN = "pattern";
	
	private static final String VALUE = "value";
	
	private static final String GT = "gt";
	
	private static final String LT = "lt";
	
	private static final String O = "o";
	
	private static final String E = "e";
	
	private static final String $ = "$";
	
	private String pattern;
	
	private String value;

	private String gt;

	private String lt;
	
	private boolean o;
	
	private boolean e;
	
	public DateValidator(String name, String emsg, Element element) {
		super(name, emsg, element);
		if (element != null) {
			pattern = element.getAttribute(PATTERN);
			value = element.getAttribute(VALUE);
			gt = element.getAttribute(GT);
			lt = element.getAttribute(LT);
			o = element.hasAttribute(O);
			e = element.hasAttribute(E);
		}
	}

	public String getPattern() {
		return StringUtils.isEmpty(pattern) ? DATEPATTERN : pattern;
	}

	public void setPattern(String pattern) {
		this.pattern = pattern;
	}

	public String getValue() {
		return value;
	}

	public void setValue(String value) {
		this.value = value;
	}

	public String getGt() {
		return gt;
	}

	public void setGt(String gt) {
		this.gt = gt;
	}

	public String getLt() {
		return lt;
	}

	public void setLt(String lt) {
		this.lt = lt;
	}

	public boolean getO() {
		return o;
	}

	public void setO(boolean o) {
		this.o = o;
	}

	public boolean getE() {
		return e;
	}

	public void setE(boolean e) {
		this.e = e;
	}

	@Override
	@SuppressWarnings("unchecked")
	public boolean check(Object arg, LangCode langCode, Map<String, Object> session) {
		if (arg == null || (!(arg instanceof String) && !(arg instanceof Map))) 	return true;
		String _value = null;
		if (arg instanceof String) {
		    _value = (String)arg;
		} else {
			if (StringUtils.isEmpty(value) || StringUtils.isEmpty(value.trim())) return true;
			Map<String, Object> map = (Map<String, Object>)arg;
			_value = value.startsWith($) ? map.get(value.substring(1)).toString() : value;
			if (!StringUtils.isEmpty(gt) && gt.startsWith($)) gt = map.get(gt.substring(1)).toString();
			if (!StringUtils.isEmpty(lt) && lt.startsWith($)) lt = map.get(lt.substring(1)).toString();
		}
		if (StringUtils.isEmpty(_value) || StringUtils.isEmpty(_value.trim())) return true;
		try {
			Date date = null;
			if ((date = getSimpleDateFormat(getPattern()).parse(_value)) != null) {	
				if (!o && !e) {
					return check1(date);
				} else if (!o && e) {
					return check3(date);
				} else if (o && e) {
					return check4(date);
				} else {
					return check2(date);
				}
			}
		} catch (ParseException e) {
			e.printStackTrace();
		}
		return false;
	}
	
	/**
	 * gt < date && date < lt
	 * @param date
	 * @return
	 * @throws ParseException
	 */
	private boolean check1(Date date) throws ParseException {
		if (!StringUtils.isEmpty(gt) && getSimpleDateFormat(getPattern()).parse(gt).compareTo(date) > -1) 	return false;
		if (!StringUtils.isEmpty(lt) && getSimpleDateFormat(getPattern()).parse(lt).compareTo(date) < 1) 	return false;
		return true;
	}

	/**
	 * gt < date || date < lt
	 * @param date
	 * @return
	 * @throws ParseException
	 */
	private boolean check2(Date date) throws ParseException {
		if ((!StringUtils.isEmpty(lt) && getSimpleDateFormat(getPattern()).parse(lt).compareTo(date) < 1)
				&& (!StringUtils.isEmpty(gt) && getSimpleDateFormat(getPattern()).parse(gt).compareTo(date) > -1)) 	return false;
		return true;
	}

	/**
	 * gte < date && date < lte
	 * @param date
	 * @return
	 * @throws ParseException
	 */
	private boolean check3(Date date) throws ParseException {
		if (!StringUtils.isEmpty(gt) && getSimpleDateFormat(getPattern()).parse(gt).compareTo(date) == 1) 	return false;
		if (!StringUtils.isEmpty(lt) && getSimpleDateFormat(getPattern()).parse(lt).compareTo(date) == -1) 	return false;
		return true;
	}

	/**
	 * gte < date || date < lte
	 * @param date
	 * @return
	 * @throws ParseException
	 */
	private boolean check4(Date date) throws ParseException {
		if ((!StringUtils.isEmpty(lt) && getSimpleDateFormat(getPattern()).parse(lt).compareTo(date) == -1)
				&& (!StringUtils.isEmpty(gt) && getSimpleDateFormat(getPattern()).parse(gt).compareTo(date) == 1)) 	return false;
		return true;
	}
	
	private SimpleDateFormat getSimpleDateFormat(String pattern) {
	    SimpleDateFormat sdf = new SimpleDateFormat(pattern);
	    sdf.setLenient(false);
	    return sdf;
	}

}