package com.newland.poi.template;

import com.deepoove.poi.config.Configure;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.render.processor.Visitor;
import com.deepoove.poi.template.ElementTemplate;
import org.apache.poi.xslf.usermodel.XSLFComment;

/**
 * 批注处理
 *
 * @author xuzhh
 * @since 1.0 2023/3/21
 */
public class CommentTemplate extends ElementTemplate {

    private XSLFComment comment;

    public CommentTemplate(String tagName, XSLFComment comment) {
        this.tagName = tagName;
        this.comment = comment;
    }

    public XSLFComment getComment() {
        return comment;
    }

    @Override
    public void accept(Visitor visitor) {
        visitor.visit(this);
    }

    @Override
    public RenderPolicy findPolicy(Configure config) {
        RenderPolicy policy = config.getCustomPolicy(tagName);
        return null == policy ? config.getTemplatePolicy(this.getClass()) : policy;
    }

}
