package com.deepoove.poi;

import com.deepoove.poi.config.Configure;
import com.deepoove.poi.resolver.ExResolver;
import com.deepoove.poi.resolver.Resolver;
import com.deepoove.poi.template.MetaTemplate;
import com.deepoove.poi.util.PoitlIOUtils;
import org.apache.poi.ooxml.POIXMLDocument;

import java.io.Closeable;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

/**
 * 接口定义解析模版，可根据解析的文档类型(docx/pptx/xlsx)提供对应的具体实现
 *
 * @author xuzhh
 * @since 1.0 2023/3/20
 */
public interface PoiTemplate<D extends POIXMLDocument> extends Closeable {

    /**
     * Get configuration
     *
     * @return {@link Configure}
     */
    Configure getConfig();

    /**
     * Get document
     *
     * @return instance of {@link POIXMLDocument}
     */
    D getDocument();

    /**
     * Get all tags in the document
     *
     * @return {@link MetaTemplate} list
     */
    List<MetaTemplate> getElementTemplates();

    /**
     * Get Resolver
     *
     * @return {@link Resolver}
     */
    ExResolver<D, ?> getResolver();

    /**
     * Render the template by data model
     *
     * @param model render data
     * @return instance of current {@link PoiTemplate}
     */
    PoiTemplate<D> render(Object model);

    /**
     * Render the template by data model and write to OutputStream, don't forget invoke {@link Closeable#close()},
     * {@link OutputStream#close()}
     *
     * @param model render data
     * @param out   output
     * @return instance of current {@link PoiTemplate}
     * @throws IOException error occurs when rendering
     */
    default PoiTemplate<D> render(Object model, OutputStream out) throws IOException {
        this.render(model);
        this.write(out);
        return this;
    }

    /**
     * write to output stream, don't forget invoke {@link Closeable#close()}, {@link OutputStream#close()} finally
     *
     * @param out eg.ServletOutputStream
     * @throws IOException error occurs when writing
     */
    void write(OutputStream out) throws IOException;

    /**
     * write to and close output stream
     *
     * @param out eg.ServletOutputStream
     * @throws IOException error occurs
     */
    default void writeAndClose(OutputStream out) throws IOException {
        try {
            this.write(out);
            out.flush();
        } finally {
            PoitlIOUtils.closeQuietly(out);
            this.close();
        }
    }

    /**
     * write to file, this method will close all the stream
     *
     * @param path output path
     * @throws IOException error occurs
     */
    default void writeToFile(String path) throws IOException {
        this.writeAndClose(new FileOutputStream(path));
    }

    /**
     * reload the template
     *
     * @param doc load new template document
     */
    void reload(D doc);

}
