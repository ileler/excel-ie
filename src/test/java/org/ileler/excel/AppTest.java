package org.ileler.excel;

import org.ileler.excel.e.ExportUtil;
import org.ileler.excel.util.LangCode;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.Hashtable;
import java.util.Map;

/**
 * Created by zhangle on 2016/5/12.
 */
public class AppTest {

    public static void main(String[] args) throws Exception {
        Map<String, Object> data = new Hashtable<>(0);
        data.put("field1", "1");
        data.put("field2", "2");
        data.put("field3", "3");
        data.put("field4", "4");
        data.put("field5", "5");
        ExportUtil.exportExcel("test", new FileOutputStream("test.xls", true), new FileInputStream("test_zh_cn.xml"), LangCode.ZH_CN, Arrays.asList(new Object[]{data}));
    }

}
