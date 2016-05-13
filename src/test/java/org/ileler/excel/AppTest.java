package org.ileler.excel;

import org.ileler.excel.e.ExportUtil;
import org.ileler.excel.i.ImportUtil;
import org.ileler.excel.i.LeadinResult;
import org.ileler.excel.util.LangCode;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.Hashtable;
import java.util.Iterator;
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
        data.put("field4", "");
        data.put("field5", "25423652145236524125365241253654");
//        ExportUtil.exportExcel("test", new FileOutputStream("test.xls", true), AppTest.class.getResourceAsStream("/e/test_zh_cn.xml"), LangCode.ZH_CN, Arrays.asList(new Object[]{data}));
        ExportUtil.exportExcel("test", new FileOutputStream("test.xls", true), AppTest.class.getResourceAsStream("/e/test_zh_cn.xml"), LangCode.ZH_CN, Arrays.asList(new Object[]{data}));
        LeadinResult result = ImportUtil.resolve(new FileInputStream("test.xls"), AppTest.class.getResourceAsStream("/i/test_zh_cn.xml"), LangCode.ZH_CN, null);
        Map<Integer, Map<String, Object>> error =  result.getErrorList();
        Map<Integer, Map<String, Object>> right =  result.getRightList();
        System.out.println("error:");
        if (error != null && error.size() > 0) {
            Iterator<Map<String, Object>> iter = error.values().iterator();
            while(iter.hasNext()) {
                Map<String, Object> obj = iter.next();
                System.out.println(obj);
            }
        }
        System.out.println("right:");
        if (right != null && right.size() > 0) {
            Iterator<Map<String, Object>> iter = right.values().iterator();
            while(iter.hasNext()) {
                Map<String, Object> obj = iter.next();
                System.out.println(obj);
            }
        }
    }

}
