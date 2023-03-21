package com.newland.poi;

import com.deepoove.poi.PoiTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.ConfigureBuilder;
import com.deepoove.poi.config.GramerSymbol;
import com.deepoove.poi.exception.ResolverException;
import com.deepoove.poi.render.DefaultRender;
import com.deepoove.poi.render.Render;
import com.deepoove.poi.template.MetaTemplate;
import com.deepoove.poi.util.PoitlIOUtils;
import com.newland.poi.policy.NoopRenderPolicy;
import com.newland.poi.policy.XslfCommentRenderPolicy;
import com.newland.poi.policy.XslfTextRenderPolicy;
import com.newland.poi.resolver.XslfElementTemplateFactory;
import com.newland.poi.resolver.XslfTemplateResolver;
import com.newland.poi.template.CommentTemplate;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * The facade of ppt(pptx) template
 * <p>
 * It works by expanding tags in a template using values provided in a Map or Object.
 * </p>
 *
 * @author xuzhh
 * @since 1.0 2023/3/15
 */
public class XslfTemplate implements PoiTemplate<XMLSlideShow> {

    private static final Logger logger = LoggerFactory.getLogger(XslfTemplate.class);

    private XMLSlideShow doc;
    private Configure config;
    private XslfTemplateResolver resolver;
    private Render renderer;
    private List<MetaTemplate> eleTemplates;

    private XslfTemplate() {
        ConfigureBuilder builder = Configure.builder();
        builder.setElementTemplateFactory(new XslfElementTemplateFactory());

        // 添加或覆盖渲染策略
        builder.addPlugin(GramerSymbol.TEXT.getSymbol(), new XslfTextRenderPolicy());
        builder.addPlugin(GramerSymbol.TEXT_ALIAS.getSymbol(), new XslfTextRenderPolicy());
        builder.addPlugin(CommentTemplate.class, new XslfCommentRenderPolicy());
        // 暂不支持的解析类型先置空
        NoopRenderPolicy noopRenderPolicy = new NoopRenderPolicy();
        builder.addPlugin(GramerSymbol.TABLE.getSymbol(), noopRenderPolicy);
        builder.addPlugin(GramerSymbol.DOCX_TEMPLATE.getSymbol(), noopRenderPolicy);
        builder.addPlugin(GramerSymbol.IMAGE.getSymbol(), noopRenderPolicy);
        builder.addPlugin(GramerSymbol.NUMBERING.getSymbol(), noopRenderPolicy);
        builder.addPlugin(GramerSymbol.ITERABLE_START.getSymbol(), noopRenderPolicy);
        builder.addPlugin(GramerSymbol.BLOCK_END.getSymbol(), noopRenderPolicy);

        config = builder.build();
    }

    /**
     * Compile template from absolute file path
     *
     * @param absolutePath template path
     * @return {@link XslfTemplate}
     */
    public static XslfTemplate compile(String absolutePath) {
        return compile(new File(absolutePath));
    }

    /**
     * Compile template from file
     *
     * @param templateFile template file
     * @return {@link XslfTemplate}
     */
    public static XslfTemplate compile(File templateFile) {
        return compile(templateFile, null);
    }

    /**
     * Compile template from template input stream
     *
     * @param inputStream template input
     * @return {@link XslfTemplate}
     */
    public static XslfTemplate compile(InputStream inputStream) {
        return compile(inputStream, null);
    }

    /**
     * Compile template from document
     *
     * @param doc template document
     * @return {@link XslfTemplate}
     */
    public static XslfTemplate compile(XMLSlideShow doc) {
        return compile(doc, null);
    }

    /**
     * Compile template from absolute file path with configure
     *
     * @param absolutePath absolute template file path
     * @param config       config
     * @return {@link XslfTemplate}
     */
    public static XslfTemplate compile(String absolutePath, Configure config) {
        return compile(new File(absolutePath), config);
    }

    /**
     * Compile template from file with configure
     *
     * @param templateFile template file
     * @param config       config
     * @return {@link XslfTemplate}
     */
    public static XslfTemplate compile(File templateFile, Configure config) {
        try {
            return compile(new FileInputStream(templateFile), config);
        } catch (FileNotFoundException e) {
            throw new ResolverException("Cannot find the file [" + templateFile.getPath() + "]", e);
        }
    }

    /**
     * Compile template from document with configure
     *
     * @param doc    template document
     * @param config config
     * @return {@link XslfTemplate}
     */
    public static XslfTemplate compile(XMLSlideShow doc, Configure config) {
        try {
            return compile(PoitlIOUtils.docToInputStream(doc), config);
        } catch (IOException e) {
            throw new ResolverException("Cannot compile document", e);
        }
    }

    /**
     * Compile template from template input stream with configure
     *
     * @param inputStream template input
     * @param config      config
     * @return {@link XslfTemplate}
     */
    public static XslfTemplate compile(InputStream inputStream, Configure config) {
        try {
            XslfTemplate template = new XslfTemplate();
            if (config != null) {
                template.config = config;
            }
            template.doc = new XMLSlideShow(inputStream);
            template.resolver = new XslfTemplateResolver(template.config);
            template.renderer = new DefaultRender();
            template.eleTemplates = template.resolver.resolveDocument(template.doc);
            return template;
        } catch (POIXMLException e) {
            logger.error("Poi-tl-ppt currently only supports .pptx format");
            throw new ResolverException("Compile template failed", e);
        } catch (IOException e) {
            throw new ResolverException("Compile template failed", e);
        }
    }

    /**
     * Render the template by data model
     *
     * @param model render data
     * @return {@link XslfTemplate}
     */
    @Override
    public XslfTemplate render(Object model) {
        this.renderer.render(this, model);
        return this;
    }

    /**
     * write to output stream, don't forget invoke {@link XslfTemplate#close()}, {@link OutputStream#close()} finally
     *
     * @param out eg.ServletOutputStream
     * @throws IOException error occurs when writing
     */
    @Override
    public void write(OutputStream out) throws IOException {
        this.doc.write(out);
    }

    /**
     * reload the template
     *
     * @param doc load new template document
     */
    @Override
    public void reload(XMLSlideShow doc) {
        PoitlIOUtils.closeLoggerQuietly(this.doc);
        this.doc = doc;
        this.eleTemplates = this.resolver.resolveDocument(doc);
    }

    @Override
    public void close() throws IOException {
        this.doc.close();
    }

    @Override
    public List<MetaTemplate> getElementTemplates() {
        return eleTemplates;
    }

    @Override
    public XMLSlideShow getDocument() {
        return this.doc;
    }

    @Override
    public Configure getConfig() {
        return config;
    }

    @Override
    public XslfTemplateResolver getResolver() {
        return resolver;
    }

}
