package com.deepoove.poi.resolver;

import com.deepoove.poi.template.MetaTemplate;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.ooxml.POIXMLDocumentPart;

import java.util.List;

/**
 * Resolver document and elements.
 *
 * @author xuzhh
 * @since 1.0 2023/3/17
 */
public interface ExResolver<D extends POIXMLDocument, P extends POIXMLDocumentPart> {

    /**
     * resolve document
     *
     * @param doc document
     * @return list of {@link MetaTemplate}
     */
    List<MetaTemplate> resolveDocument(D doc);

    /**
     * resolve document elements
     *
     * @param elements document elements
     * @return list of {@link MetaTemplate}
     */
    List<MetaTemplate> resolveDocumentParts(List<? extends P> elements);

}
