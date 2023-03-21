package com.newland.poi.template;

import com.deepoove.poi.render.processor.Visitor;
import com.deepoove.poi.template.ElementTemplate;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.sl.usermodel.TextRun;

import java.util.List;

/**
 * 文本处理
 *
 * @author xuzhh
 * @since 1.0 2023/3/19
 */
public class TextRunTemplate extends ElementTemplate {

    protected TextRun run;

    public TextRunTemplate(String tagName, Character sign, TextRun run) {
        this.tagName = tagName;
        this.sign = sign;
        this.run = run;
    }

    /**
     * @return the run
     */
    public TextRun getRun() {
        return run;
    }

    public Integer getRunPos() {
        TextParagraph<?, ?, ?> paragraph = run.getParagraph();
        List<?> textRuns = paragraph.getTextRuns();
        for (int i = 0; i < textRuns.size(); i++) {
            if (run == textRuns.get(i)) {
                return i;
            }
        }
        return null;
    }

    @Override
    public void accept(Visitor visitor) {
        visitor.visit(this);
    }

}
