package org.ileler.excel;

import org.ileler.excel.e.ExportUtil;
import org.ileler.excel.util.LangCode;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Hashtable;
import java.util.List;
import java.util.Map;

/**
 * Created by zhangle on 2016/5/12.
 */
public class AppTest {

    public static void main(String[] args) throws FileNotFoundException {
        Map<String, Object> data = new Hashtable<>(0);
        data.put("", "");
        ExportUtil.exportExcel("test", new FileOutputStream("test.xls", true), new FileInputStream("test_zh_cn.xml"), LangCode.ZH_CN, );
    }

}
