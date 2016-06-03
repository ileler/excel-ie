package org.ileler.excel.wrapper;

import org.apache.commons.lang3.StringUtils;
import org.ileler.excel.util.LangCode;
import org.w3c.dom.Element;

import java.text.SimpleDateFormat;
import java.util.Map;

/**
 * Created by ileler@qq.com on 2016/5/12.
 */
public class DateWrapper extends Wrapper {

    private static final String PATTERN = "pattern";

    private String pattern;

    public DateWrapper(Element element, Map<String, String> map) {
        super(element, map);
        if (map != null && map.size() > 0) {
            pattern = map.get(PATTERN);
        }
        if (StringUtils.isBlank(pattern)) {
            throw new IllegalArgumentException(PATTERN + " is null.");
        }
    }

    @Override
    public String value(Object value, Object rowValue, LangCode langCode) {
        if (StringUtils.isBlank(pattern) || value == null || "".equals(value.toString().trim())) {
            return "";
        }
        try {
            return new SimpleDateFormat(pattern).format(value);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return value.toString();
    }

}