package com.newland.poi.xslf;

import com.deepoove.poi.exception.ReflectionException;
import com.deepoove.poi.util.ReflectionUtils;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextLineBreak;

import java.lang.reflect.Method;
import java.util.List;

/**
 * wrapper of {@link XSLFTextParagraph}
 *
 * @author xuzhh
 * @since 1.0 2023/3/19
 */
public class XslfTextParagraphWrapper {

    XSLFTextParagraph paragraph;

    public XslfTextParagraphWrapper(XSLFTextParagraph paragraph) {
        this.paragraph = paragraph;
    }

    public XSLFTextParagraph getParagraph() {
        return paragraph;
    }

    public XSLFTextRun addNewTextRun() {
        XSLFTextRun newRun = paragraph.addNewTextRun();
        return newRun;
    }

    public boolean removeTextRun(XSLFTextRun textRun) {
        return paragraph.removeTextRun(textRun);
    }

    public XSLFTextRun insertNewTextRun(int pos) {
        int lastPos = paragraph.getTextRuns().size() - 1;
        // 插入位置在列表尾部
        if (pos > lastPos) {
            return paragraph.addNewTextRun();
        } else if (pos >= 0) {
            XSLFTextRun newRun = newTextRun(paragraph.getXmlObject().insertNewR(pos));
            CTTextCharacterProperties textProps = newRun.getRPr(true);
            if (pos > 0) {
                // by default text run has the style of the previous one
                CTTextCharacterProperties prevRun = paragraph.getTextRuns().get(pos - 1).getRPr(true);
                textProps.set(prevRun);
            }
            getTextRuns().add(pos, newRun);
            return newRun;
        }
        return null;
    }

    public XSLFTextRun insertLineBreak(int pos) {
        int lastPos = paragraph.getTextRuns().size() - 1;
        if (pos >= lastPos) {
            paragraph.addLineBreak();
        } else if (pos >= 0) {
            XSLFTextRun newRun = newLineBreak(paragraph.getXmlObject().insertNewBr(pos));
            CTTextCharacterProperties brProps = newRun.getRPr(true);
            if (pos > 0) {
                // by default line break has the font size of the previous text run
                CTTextCharacterProperties prevRun = paragraph.getTextRuns().get(pos - 1).getRPr(true);
                brProps.set(prevRun);
                // don't copy hlink properties
                if (brProps.isSetHlinkClick()) {
                    brProps.unsetHlinkClick();
                }
                if (brProps.isSetHlinkMouseOver()) {
                    brProps.unsetHlinkMouseOver();
                }
            }
            getTextRuns().add(pos, newRun);
            return newRun;
        }
        return null;
    }

    public XSLFTextRun insertNewTextRunAfter(int pos) {
        return insertNewTextRun(pos + 1);
    }

    @SuppressWarnings("unchecked")
    private List<XSLFTextRun> getTextRuns() {
        return (List<XSLFTextRun>) ReflectionUtils.getValue("_runs", paragraph);
    }

    private XSLFTextRun newTextRun(XmlObject obj) {
        try {
            Method method = ReflectionUtils.findMethod(paragraph.getClass(), "newTextRun", XmlObject.class);
            return (XSLFTextRun) method.invoke(paragraph, obj);
        } catch (Exception e) {
            throw new ReflectionException("newTextRun", XSLFTextParagraph.class, e);
        }
    }

    private XSLFTextRun newLineBreak(CTTextLineBreak br) {
        try {
            Method method = ReflectionUtils.findMethod(paragraph.getClass(), "newTextRun", CTTextLineBreak.class);
            return (XSLFTextRun) method.invoke(paragraph, br);
        } catch (Exception e) {
            throw new ReflectionException("newTextRun", XSLFTextParagraph.class, e);
        }
    }

}
