package com.javasoso.util.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import org.apache.commons.lang3.RandomStringUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Assert;
import org.junit.Test;

/**
 * 工具测试
 *
 * @author jasonzhu
 * @date 2018/3/21
 */
public class UtilTest {

    @Test
    public void testBuildExcel() throws Exception {
        // 注意不要太多 当心内存撑爆
        int n = 100;
        List<TestModel> modelList = new ArrayList<>();
        for (int i = 0; i < n-1; i++) {
            TestModel t = new TestModel();
            t.setAge(i);
            t.setBirD(new Date());
            t.setRealName(
                RandomStringUtils.random(3, "事发后为女雷电口覅额叫我覅后果就是浪i哦好就离开发到哪里辜负你，新劳动法第三方登录个费"));
            t.setRemark(RandomStringUtils
                .random(7, "飞机豆浆粉定位饿就饿fsdfsdf房东i几个hirejofjdnbvkfnvewoijfoi1232oj5345k343k"));
            t.setAmount(new BigDecimal(1234.54));
            System.out.println(t.toString());
            modelList.add(t);
        }
        // 生成excel文件
        Workbook workbook = ExcelUtil.createWorkBook(modelList, TestModel.class, "测试", 0);
        File file = new File("./test.xls");
        OutputStream os = new FileOutputStream(file);
        workbook.write(os);
        os.flush();
        System.out.println("===========================");
        System.out.println("生成文件，model共：" + n);
        System.out.println("===========================");

        // 从第二行开始，读取到倒数第二行
        List<TestModel> result = ExcelUtil.getModelList(file, 0, TestModel.class, 1, -2);
        for (TestModel testModel : result) {
            System.out.println(testModel.toString());
        }
        System.out.println("===========================");
        System.out.println("读取文件，model共：" + result.size());
        System.out.println("===========================");
    }
}
