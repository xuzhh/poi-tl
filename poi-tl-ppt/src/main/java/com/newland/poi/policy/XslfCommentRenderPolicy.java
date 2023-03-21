package com.newland.poi.policy;

import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.policy.TextRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.template.ElementTemplate;
import com.newland.poi.template.CommentTemplate;
import org.apache.poi.xslf.usermodel.XSLFComment;

/**
 * 批注文本处理
 *
 * @author xuzhh
 * @since 1.0 2023/3/21
 */
public class XslfCommentRenderPolicy extends TextRenderPolicy {

    @Override
    public void doRender(RenderContext<TextRenderData> context) {
        ElementTemplate eleTemplate = context.getEleTemplate();
        if (eleTemplate instanceof CommentTemplate) {
            XSLFComment comment = ((CommentTemplate) eleTemplate).getComment();
            TextRenderData data = context.getData();
            String placeholder = context.getTagSource();
            comment.setText(comment.getText().replace(placeholder, data.getText()));
        }
    }

}
