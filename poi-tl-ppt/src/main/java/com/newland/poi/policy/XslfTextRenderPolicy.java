package com.newland.poi.policy;

import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.policy.TextRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.template.ElementTemplate;
import com.newland.poi.template.TextRunTemplate;
import org.apache.poi.sl.usermodel.TextRun;

/**
 * 文本替换
 *
 * @author xuzhh
 * @since 1.0 2023/3/21
 */
public class XslfTextRenderPolicy extends TextRenderPolicy {

    @Override
    public void doRender(RenderContext<TextRenderData> context) {
        ElementTemplate eleTemplate = context.getEleTemplate();
        if (eleTemplate instanceof TextRunTemplate) {
            // Helper.renderTextRun(((TextRunTemplate) eleTemplate).getRun(), context.getData());
            TextRun run = ((TextRunTemplate) eleTemplate).getRun();
            run.setText(context.getData().getText());
        }
    }

}
