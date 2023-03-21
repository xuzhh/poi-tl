package com.newland.poi.resolver;

import com.deepoove.poi.config.Configure;
import com.deepoove.poi.resolver.DefaultElementTemplateFactory;
import com.deepoove.poi.template.ChartTemplate;
import com.deepoove.poi.template.ElementTemplate;
import com.newland.poi.template.CommentTemplate;
import com.newland.poi.template.TextRunTemplate;
import org.apache.poi.sl.usermodel.TextRun;
import org.apache.poi.xddf.usermodel.chart.XDDFChart;
import org.apache.poi.xslf.usermodel.XSLFComment;

import java.util.Set;

/**
 * 构造{@link ElementTemplate}实例的工厂类
 *
 * @author xuzhh
 * @since 1.0 2023/3/19
 */
public class XslfElementTemplateFactory extends DefaultElementTemplateFactory {

    public TextRunTemplate createTextRunTemplate(Configure config, String tag, TextRun run) {
        Set<Character> gramerChars = config.getGramerChars();
        Character symbol = EMPTY_CHAR;
        if (!"".equals(tag)) {
            char firstChar = tag.charAt(0);
            for (Character character : gramerChars) {
                if (character.equals(firstChar)) {
                    symbol = firstChar;
                    break;
                }
            }
        }

        String tagName = symbol.equals(EMPTY_CHAR) ? tag : tag.substring(1);
        TextRunTemplate template = new TextRunTemplate(tagName, symbol, run);
        template.setSource(config.getGramerPrefix() + tag + config.getGramerSuffix());
        return template;
    }

    public ChartTemplate createChartTemplate(Configure config, String tag, XDDFChart chart) {
        ChartTemplate template = new ChartTemplate(tag, chart, null);
        template.setSource(config.getGramerPrefix() + tag + config.getGramerSuffix());
        template.setSign(EMPTY_CHAR);
        return template;
    }

    public CommentTemplate createCommentTemplate(Configure config, String tag, XSLFComment comment) {
        CommentTemplate template = new CommentTemplate(tag, comment);
        template.setSource(config.getGramerPrefix() + tag + config.getGramerSuffix());
        template.setSign(EMPTY_CHAR);
        return template;
    }

}
