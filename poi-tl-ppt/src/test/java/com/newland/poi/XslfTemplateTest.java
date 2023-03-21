package com.newland.poi;

import com.deepoove.poi.data.ChartMultiSeriesRenderData;
import com.deepoove.poi.data.Charts;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * 测试ppt模版渲染
 *
 * @author xuzhh
 * @since 1.0 2023/3/17
 */
class XslfTemplateTest {

    @Test
    void testRenderTemplate() throws IOException {
        Map<String, Object> data = new HashMap<>();
        data.put("tableTitle", "单元格标题");
        data.put("tableCell", "单元格文本");
        data.put("shape", "填充形状");
        data.put("textbox", "【填充文本框】");
        data.put("artFont", "<填充艺术字>");
        data.put("footness", "填充页脚");
        data.put("comment", "填充批注");

        Map<String, String> notes = new HashMap<>();
        notes.put("index1", "#1");
        notes.put("index2", "#2");
        notes.put("index3", "#3");
        data.put("notes", notes);

        ChartMultiSeriesRenderData chartData = Charts.ofMultiSeries("语言学习",
                new String[]{"中文", "English", "日本語", "português", "中文", "English", "日本語", "português"}).addSeries("countries",
                new Double[]{15.0, 6.0, 18.0, 231.0, 150.0, 6.0, 118.0, 31.0}).addSeries("speakers",
                new Double[]{223.0, 119.0, 154.0, 142.0, 223.0, 119.0, 54.0, 42.0}).addSeries("youngs",
                new Double[]{323.0, 89.0, 54.0, 42.0, 223.0, 119.0, 54.0, 442.0}).create();
        data.put("chart", chartData);

        XslfTemplate template = XslfTemplate.compile("src/test/resources/template/test_struct.pptx").render(data);
        template.writeToFile("target/out_test_struct.pptx");
    }

}
