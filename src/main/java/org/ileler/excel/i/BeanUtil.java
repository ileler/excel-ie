package org.ileler.excel.i;

import org.apache.commons.lang3.StringUtils;
import org.ileler.excel.util.LangCode;
import org.ileler.excel.validator.Validator;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

public final class BeanUtil {
	
	private static final String VALIDATOR = "validator";
	
	private static final String NAME = "name";
	
	private static final String EMSG = "emsg";
	
	private static final String FIELD = "field";
	
	private static final String TITLE = "title";
	
	private static final String EFNAME = "efname";
	
	private static final String RFNAME = "rfname";
	
	private static final String DVALUE = "dvalue";
	
	private static final String ROWCHECK = "rowcheck";
	
	private static final String REFNAME = "efname";
	
	private static final String ESPLIT = "esplit";
	
	private static final String ETITLE = "etitle";
	
	private static final String STARTWITH = ".";
	
	private static ThreadLocal<DocumentBuilder> docBuildeIns = new ThreadLocal<DocumentBuilder>() {
		protected DocumentBuilder initialValue() {
			try {
				return DocumentBuilderFactory.newInstance().newDocumentBuilder();
			} catch (ParserConfigurationException e) {
				String msg = "DocumentBuilder 对象初始化失败！";
				throw new IllegalStateException(msg, e);
			}
		}
	};
	
	private static Validator[] getRuleBeans(NodeList ruleNodeList) {
		List<Validator> erbList = new ArrayList<Validator>();
		for(int a = 0, b = ruleNodeList.getLength(); a < b && ruleNodeList.item(a) instanceof Element; a++) {
			Element ruleElement = (Element) ruleNodeList.item(a);
			String name = null, emsg = null;
			if (StringUtils.isEmpty(name = ruleElement.getAttribute(NAME)) || StringUtils.isEmpty(emsg = ruleElement.getAttribute(EMSG))) 	continue;
			try {
				erbList.add((Validator) Class.forName((name.startsWith(STARTWITH) ? Validator.class.getPackage().getName() : "") + name).getConstructor(String.class, String.class, Element.class).newInstance(name, emsg, ruleElement));
			} catch (InstantiationException | IllegalAccessException
					| IllegalArgumentException
					| InvocationTargetException | NoSuchMethodException
					| SecurityException | ClassNotFoundException e) {
				e.printStackTrace();
			}
		}
		return erbList.toArray(new Validator[erbList.size()]);
	}
	
	public static EntityBean getEntityBean(InputStream src, LangCode langCode) {
		if (src == null) 	return null;
		try {
			Document doc = docBuildeIns.get().parse(src);
			NodeList fieldNodeList = doc.getElementsByTagName(FIELD);
			if (fieldNodeList == null || fieldNodeList.getLength() < 1) 	return null;
			List<FieldBean> efbList = new ArrayList<FieldBean>();
			for(int i = 0, j = fieldNodeList.getLength(); i < j && fieldNodeList.item(i) instanceof Element; i++) {
				Element fieldElement = (Element) fieldNodeList.item(i);
				String title = null,efName = null,rfName = fieldElement.getAttribute(RFNAME);
				if (StringUtils.isEmpty(title = fieldElement.getAttribute(TITLE)) || StringUtils.isEmpty(efName = fieldElement.getAttribute(EFNAME))) 	continue;
				efbList.add(new FieldBean(efbList.size(), title, efName, rfName, fieldElement.getAttribute(DVALUE), langCode, getRuleBeans(fieldElement.getElementsByTagName(VALIDATOR))));
			}
			if (efbList.size() < 1) 	return null;
			FieldBean[] fieldBeans = efbList.toArray(new FieldBean[efbList.size()]);
			NodeList rowCheckNodeList = doc.getElementsByTagName(ROWCHECK);
			if (rowCheckNodeList != null && rowCheckNodeList.getLength() > 0) {
				Element rowCheckElement = (Element) rowCheckNodeList.item(0);
				return new EntityBean(rowCheckElement.getAttribute(REFNAME), rowCheckElement.getAttribute(ESPLIT), rowCheckElement.getAttribute(ETITLE), fieldBeans, getRuleBeans(rowCheckElement.getElementsByTagName(VALIDATOR)));
			}
			return new EntityBean(null, null, null, fieldBeans, null);
		} catch (SAXException | IOException e) {
			e.printStackTrace();
		}
		return null;
	}
	
//	public static EntityBean getEntityBean(String src) {
//		if (StringUtils.isEmpty(src)) 	return null;
//		return getEntityBean(new ByteArrayInputStream(src.getBytes()));
//	}
	
	public static EntityBean getEntityBean(InputStream is, String[] titles, LangCode langCode) {
        if (is == null)   return null;
        if (langCode == null) langCode = LangCode.ZH_CN;
        EntityBean _entityBean = getEntityBean(is, langCode);
        FieldBean[] fieldBeans = _entityBean == null ? null : _entityBean.getFieldBeans();
        if (fieldBeans == null) {
             return null;
        }
        try {
            int i = 0, j = fieldBeans.length, a = titles.length;
            if (a >= j) {
                for (; i < j && i < titles.length; i++) {
                    FieldBean fieldBean = fieldBeans[i];
                    if (!fieldBean.getTitle().equalsIgnoreCase(titles[i])) {
                        break;
                    }
                }
                if (i >= j) {
                    return _entityBean;
                }
            }
        } catch (Exception e) {}
        return null;
    }
	
	public static EntityBean getEntityBean(InputStream[] iss, String[] titles, LangCode langCode) {
        if (iss == null || iss.length < 1)   return null;
        for (InputStream is : iss) {
            EntityBean _entityBean = getEntityBean(is, titles, langCode);
            FieldBean[] fieldBeans = _entityBean == null ? null : _entityBean.getFieldBeans();
            if (fieldBeans != null) {
                int i = 0, j = fieldBeans.length, a = titles.length;
                if (a < j) {
                    continue;
                }
                for (; i < j && i < titles.length; i++) {
                    FieldBean fieldBean = fieldBeans[i];
                    if (!fieldBean.getTitle().equalsIgnoreCase(titles[i])) {
                        break;
                    }
                }
                if (i < j) {
                    continue;
                } else {
                    return _entityBean;
                }
            }
        }
        return null;
    }
	
	public static InputStream getTemplateName(InputStream[] iss, String[] titles, LangCode langCode) {
        if (iss == null || iss.length < 1)   return null;
        for (InputStream is : iss) {
            EntityBean _entityBean = getEntityBean(is, titles, langCode);
            FieldBean[] fieldBeans = _entityBean == null ? null : _entityBean.getFieldBeans();
            if (fieldBeans != null) {
                int i = 0, j = fieldBeans.length, a = titles.length;
                if (a < j) {
                    continue;
                }
                for (; i < j && i < titles.length; i++) {
                    FieldBean fieldBean = fieldBeans[i];
                    if (!fieldBean.getTitle().equalsIgnoreCase(titles[i])) {
                        break;
                    }
                }
                if (i < j) {
                    continue;
                } else {
                    return is;
                }
            }
        }
        return null;
    }
	
}
