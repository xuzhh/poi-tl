package com.newland.poi.policy;

import com.deepoove.poi.PoiTemplate;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.template.ElementTemplate;

/**
 * 空操作处理
 *
 * @author xuzhh
 * @since 1.0 2023/3/21
 */
public class NoopRenderPolicy implements RenderPolicy {

    @Override
    public void render(ElementTemplate eleTemplate, Object data, PoiTemplate<?> template) {
        // no-op
    }

}
