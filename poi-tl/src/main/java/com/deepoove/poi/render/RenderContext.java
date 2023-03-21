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
package com.deepoove.poi.render;

import com.deepoove.poi.PoiTemplate;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.template.run.RunTemplate;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * Template context
 *
 * @author Sayi
 */
public class RenderContext<T> {

    private final ElementTemplate eleTemplate;
    private final T data;
    private final PoiTemplate<?> template;
    private final WhereDelegate where;

    public RenderContext(ElementTemplate eleTemplate, T data, PoiTemplate<?> template) {
        this.eleTemplate = eleTemplate;
        this.data = data;
        this.template = template;
        if (eleTemplate instanceof RunTemplate) {
            where = new WhereDelegate(((RunTemplate) this.eleTemplate).getRun());
        } else {
            where = null;
        }
    }

    public ElementTemplate getEleTemplate() {
        return eleTemplate;
    }

    public T getThing() {
        return data;
    }

    public T getData() {
        return data;
    }

    public PoiTemplate<?> getTemplate() {
        return template;
    }

    public NiceXWPFDocument getXWPFDocument() {
        return template instanceof XWPFTemplate ? ((XWPFTemplate) template).getXWPFDocument() : null;
    }

    public WhereDelegate getWhereDelegate() {
        return where;
    }

    public XWPFRun getWhere() {
        return getRun();
    }

    public XWPFRun getRun() {
        return ((RunTemplate) eleTemplate).getRun();
    }

    public IBody getContainer() {
        // XWPFTableCell、XWPFDocument、XWPFHeaderFooter、XWPFAbstractFootnoteEndnote
        return ((XWPFParagraph) getRun().getParent()).getBody();
    }

    public Configure getConfig() {
        return getTemplate().getConfig();
    }

    public String getTagSource() {
        return getEleTemplate().getSource();
    }

}
