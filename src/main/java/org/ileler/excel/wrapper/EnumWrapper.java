package org.ileler.excel.wrapper;

import org.ileler.excel.util.LangCode;
import org.w3c.dom.Element;

import java.util.Map;

/**
 * Created by ileler@qq.com on 2016/5/12.
 */
public class EnumWrapper extends Wrapper {

    public EnumWrapper(Element element, Map<String, String> map) {
        super(element, map);
        if (map == null || map.size() < 1) {
            throw new IllegalArgumentException("value is null.");
        }
    }

    @Override
    public String value(Object value, Object rowValue, LangCode langCode) {
        if (value == null || map == null || map.size() < 1) {
            return "";
        }
        return map.get(value.toString() + "-" + langCode);
    }

}