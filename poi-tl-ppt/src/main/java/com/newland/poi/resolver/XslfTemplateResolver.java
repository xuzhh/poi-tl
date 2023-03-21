package com.newland.poi.resolver;

import com.deepoove.poi.config.Configure;
import com.deepoove.poi.resolver.AbstractResolver;
import com.deepoove.poi.resolver.ElementTemplateFactory;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.template.MetaTemplate;
import com.newland.poi.xslf.XslfChartWrapper;
import org.apache.commons.lang3.ClassUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.sl.usermodel.TextRun;
import org.apache.poi.xddf.usermodel.chart.XDDFChart;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFComment;
import org.apache.poi.xslf.usermodel.XSLFGraphicFrame;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFNotesMaster;
import org.apache.poi.xslf.usermodel.XSLFObjectShape;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.function.Function;
import java.util.regex.Matcher;

/**
 * POI PPT document resolver.
 *
 * @author xuzhh
 * @since 1.0 2023/3/17
 */
public class XslfTemplateResolver extends AbstractResolver<XMLSlideShow, POIXMLDocumentPart> {

    private static final Logger logger = LoggerFactory.getLogger(XslfTemplateResolver.class);

    private final XslfElementTemplateFactory elementTemplateFactory;

    private final Map<Class<? extends POIXMLDocumentPart>, Function<? super POIXMLDocumentPart, List<MetaTemplate>>> elementResolveMap;
    private final Map<Class<? extends XSLFShape>, BiConsumer<? super XSLFShape, List<MetaTemplate>>> shapeResolveMap;

    public XslfTemplateResolver(Configure config) {
        this(config, config.getElementTemplateFactory());
    }

    private XslfTemplateResolver(Configure config, ElementTemplateFactory elementTemplateFactory) {
        super(config);
        this.elementTemplateFactory = (XslfElementTemplateFactory) elementTemplateFactory;
        // 定义支持解析的element类型及对应解析方法
        this.elementResolveMap = new HashMap<>(3);
        this.elementResolveMap.put(XSLFSlide.class, this::resolveSlide);
        // 定义支持解析的shape类型及对应解析方法
        this.shapeResolveMap = new HashMap<>(10);
        this.shapeResolveMap.put(XSLFGroupShape.class, this::resolveGroupShape);
        this.shapeResolveMap.put(XSLFTextShape.class, this::resolveTextShape);
        this.shapeResolveMap.put(XSLFGraphicFrame.class, this::resolveChart);
        this.shapeResolveMap.put(XSLFTable.class, this::resolveTable);
        this.shapeResolveMap.put(XSLFPictureShape.class, this::resolvePictureShape);
        this.shapeResolveMap.put(XSLFObjectShape.class, this::resolveObjectShape);
    }

    @Override
    public List<MetaTemplate> resolveDocument(XMLSlideShow doc) {
        List<MetaTemplate> metaTemplates = new ArrayList<>();
        if (null == doc) {
            return metaTemplates;
        }

        logger.info("Resolve the document start...");
        // 幻灯片页
        metaTemplates.addAll(resolveDocumentParts(doc.getSlides()));
        // 幻灯片母版
        metaTemplates.addAll(resolveSlideMasters(doc.getSlideMasters()));
        // 备注母版
        XSLFNotesMaster notesMaster = doc.getNotesMaster();
        if (null != notesMaster) {
            metaTemplates.addAll(resolveNotesMaster(notesMaster));
        }
        logger.info("Resolve the document end, resolve and create {} MetaTemplates.", metaTemplates.size());
        return metaTemplates;
    }

    /**
     * 解析幻灯片母版
     *
     * @param masters 幻灯片母版
     * @return list of {@link MetaTemplate}
     */
    protected List<MetaTemplate> resolveSlideMasters(List<XSLFSlideMaster> masters) {
        // TODO 解析幻灯片母版
        return Collections.emptyList();
    }

    /**
     * 解析备注母版
     *
     * @param master 备注母版
     * @return list of {@link MetaTemplate}
     */
    protected List<MetaTemplate> resolveNotesMaster(XSLFNotesMaster master) {
        // TODO 解析备注母版
        return Collections.emptyList();
    }

    @Override
    public List<MetaTemplate> resolveDocumentParts(List<? extends POIXMLDocumentPart> elements) {
        List<MetaTemplate> metaTemplates = new ArrayList<>();
        if (null == elements) {
            return metaTemplates;
        }

        // current iterable templates state
        // Deque<BlockTemplate> stack = new LinkedList<BlockTemplate>();

        for (POIXMLDocumentPart e : elements) {
            if (null == e || !elementResolveMap.containsKey(e.getClass())) {
                continue;
            }
            metaTemplates.addAll(elementResolveMap.get(e.getClass()).apply(e));
        }

        // checkStack(stack);
        return metaTemplates;
    }

    /**
     * 解析幻灯片页
     *
     * @param slide 幻灯片
     * @return list of {@link MetaTemplate}
     */
    protected List<MetaTemplate> resolveSlide(POIXMLDocumentPart slide) {
        List<MetaTemplate> metaTemplates = new ArrayList<>();
        if (!(slide instanceof XSLFSlide)) {
            return metaTemplates;
        }

        XSLFSlide xs = (XSLFSlide) slide;
        // 解析形状
        xs.getShapes().forEach(s -> resolveShape(s, metaTemplates));
        // 解析批注
        xs.getComments().forEach(c -> resolveComment(c, metaTemplates));
        // 解析备注
        if (null != xs.getNotes()) {
            xs.getNotes().getTextParagraphs().stream().flatMap(Collection::stream).forEach(p -> resolveParagraph(p, metaTemplates));
        }

        return metaTemplates;
    }

    /**
     * 解析形状
     *
     * @param shape     形状
     * @param templates 待填充的模版集合，解析结果依次填充到列表中
     */
    protected void resolveShape(XSLFShape shape, List<MetaTemplate> templates) {
        if (null == shape) {
            return;
        }

        Class<? extends XSLFShape> shapeClass = shape.getClass();
        if (!shapeResolveMap.containsKey(shapeClass)) {
            for (Class<? extends XSLFShape> keyClass : shapeResolveMap.keySet()) {
                if (keyClass.isInstance(shape)) {
                    shapeClass = keyClass;
                    break;
                }
            }
        }

        if (shapeResolveMap.containsKey(shapeClass)) {
            shapeResolveMap.get(shapeClass).accept(shape, templates);
        }
    }

    /**
     * 解析形状组
     *
     * @param group     形状组
     * @param templates 待填充的模版集合，解析结果依次填充到列表中
     */
    protected void resolveGroupShape(XSLFShape group, List<MetaTemplate> templates) {
        if (group instanceof XSLFGroupShape) {
            ((XSLFGroupShape) group).iterator().forEachRemaining(s -> resolveShape(s, templates));
        }
    }

    /**
     * 解析文本框
     *
     * @param text      文本框
     * @param templates 待填充的模版集合，解析结果依次填充到列表中
     */
    protected void resolveTextShape(XSLFShape text, List<MetaTemplate> templates) {
        if (text instanceof XSLFTextShape) {
            ((XSLFTextShape) text).getTextParagraphs().forEach(p -> resolveParagraph(p, templates));
        }
    }

    /**
     * 解析表格
     *
     * @param table     表格
     * @param templates 待填充的模版集合，解析结果依次填充到列表中
     */
    protected void resolveTable(XSLFShape table, List<MetaTemplate> templates) {
        if (table instanceof XSLFTable) {
            List<XSLFTableRow> rows = ((XSLFTable) table).getRows();
            if (null == rows || rows.isEmpty()) {
                return;
            }

            for (XSLFTableRow row : rows) {
                List<XSLFTableCell> cells = row.getCells();
                if (null == cells || cells.isEmpty()) {
                    continue;
                }
                cells.forEach(c -> resolveTextShape(c, templates));
            }
        }
    }

    /**
     * 解析图表
     *
     * @param shape     图表
     * @param templates 待填充的模版集合，解析结果依次填充到列表中
     */
    protected void resolveChart(XSLFShape shape, List<MetaTemplate> templates) {
        if (shape instanceof XSLFGraphicFrame && ((XSLFGraphicFrame) shape).getChart() != null) {
            XslfChartWrapper chartWrapper = new XslfChartWrapper(((XSLFGraphicFrame) shape).getChart(), shape);
            ElementTemplate template = parseTemplateFactory(chartWrapper.getTitle(), chartWrapper.getChart());
            if (template == null) {
                template = parseTemplateFactory(chartWrapper.getDesc(), chartWrapper.getChart());
            }
            if (template != null) {
                templates.add(template);
            }
        }
    }

    /**
     * 解析图片
     *
     * @param picture   图片
     * @param templates 待填充的模版集合，解析结果依次填充到列表中
     */
    protected void resolvePictureShape(XSLFShape picture, List<MetaTemplate> templates) {
        // TODO 解析图片
    }

    /**
     * 解析内嵌对象
     *
     * @param object    对象
     * @param templates 待填充的模版集合，解析结果依次填充到列表中
     */
    protected void resolveObjectShape(XSLFShape object, List<MetaTemplate> templates) {
        // TODO 解析内嵌对象
    }

    /**
     * 解析批注
     *
     * @param comment   批注
     * @param templates 待填充的模版集合，解析结果依次填充到列表中
     */
    protected void resolveComment(XSLFComment comment, List<MetaTemplate> templates) {
        if (null == comment) {
            return;
        }

        Matcher matcher = templatePattern.matcher(comment.getText());
        while (matcher.find()) {
            String placeholder = matcher.group();
            ElementTemplate template = parseTemplateFactory(placeholder, comment);
            if (template != null) {
                templates.add(template);
            }
        }
    }

    /**
     * 解析文本段
     *
     * @param paragraph 文本段
     * @param templates 待填充的模版集合，解析结果依次填充到列表中
     */
    protected void resolveParagraph(XSLFTextParagraph paragraph, List<MetaTemplate> templates) {
        List<XSLFTextRun> runs = new TextParagraphRefactor(paragraph, templatePattern).refactorRun();
        // 生成template
        for (XSLFTextRun run : runs) {
            if (StringUtils.isBlank(run.getRawText())) {
                continue;
            }
            ElementTemplate template = parseTemplateFactory(run.getRawText(), run);
            if (template != null) {
                templates.add(template);
            }
        }
    }

    protected ElementTemplate parseTemplateFactory(String text, Object obj) {
        if (null == text) {
            return null;
        }
        ElementTemplate elementTemplate = null;
        if (templatePattern.matcher(text).matches()) {
            String shortClassName = ClassUtils.getShortClassName(obj.getClass());
            String tag = gramerPattern.matcher(text).replaceAll("").trim();

            if (obj instanceof XSLFTextRun) {
                elementTemplate = elementTemplateFactory.createTextRunTemplate(config, tag, (TextRun) obj);
            } else if (obj instanceof XSLFComment) {
                elementTemplate = elementTemplateFactory.createCommentTemplate(config, tag, (XSLFComment) obj);
            } else if (obj instanceof XSLFChart) {
                elementTemplate = elementTemplateFactory.createChartTemplate(config, tag, (XDDFChart) obj);
            }

            if (null != elementTemplate) {
                logger.debug("Resolve where text: {}, and create {} for {}", text,
                        ClassUtils.getShortClassName(elementTemplate.getClass()), shortClassName);
            }
        }
        return elementTemplate;
    }

}
