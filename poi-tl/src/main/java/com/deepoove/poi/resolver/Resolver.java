/*
 * Copyright 2014-2021 Sayi
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.deepoove.poi.resolver;

import com.deepoove.poi.template.MetaTemplate;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.util.Collections;
import java.util.List;

/**
 * Resolver document and part
 *
 * @author Sayi
 */
public interface Resolver extends ExResolver<XWPFDocument, POIXMLDocumentPart> {

    /**
     * resolve document
     *
     * @param doc document
     * @return list of {@link MetaTemplate}
     */
    @Override
    List<MetaTemplate> resolveDocument(XWPFDocument doc);

    /**
     * resolve document elements
     *
     * @param elements document elements
     * @return list of {@link MetaTemplate}
     */
    @Override
    default List<MetaTemplate> resolveDocumentParts(List<? extends POIXMLDocumentPart> elements) {
        return Collections.emptyList();
    }

    /**
     * resolve body elements
     *
     * @param bodyElements
     * @return
     */
    List<MetaTemplate> resolveBodyElements(List<IBodyElement> bodyElements);

    /**
     * resolve runs at same paragraph
     *
     * @param runs
     * @return
     */
    List<MetaTemplate> resolveXWPFRuns(List<XWPFRun> runs);

}
