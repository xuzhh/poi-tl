package com.newland.poi.xslf;

import com.deepoove.poi.exception.ReflectionException;
import com.deepoove.poi.util.ReflectionUtils;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;

import java.lang.reflect.Method;
import java.util.Objects;

/**
 * wrapper of {@link XSLFChart}
 *
 * @author xuzhh
 * @since 1.0 2023/3/20
 */
public class XslfChartWrapper {

    XSLFShape container;
    XSLFChart chart;

    public XslfChartWrapper(XSLFChart chart, XSLFShape container) {
        Objects.requireNonNull(chart, "chart cannot be null !");
        Objects.requireNonNull(container, "container cannot be null !");
        this.chart = chart;
        this.container = container;
    }

    public XSLFChart getChart() {
        return chart;
    }

    public XSLFShape getContainer() {
        return container;
    }

    public String getTitle() {
        return getCNvPr(container).getTitle();
    }

    public String getDesc() {
        return getCNvPr(container).getDescr();
    }

    private CTNonVisualDrawingProps getCNvPr(XSLFShape shape) {
        try {
            Method method = ReflectionUtils.findMethod(shape.getClass(), "getCNvPr");
            return (CTNonVisualDrawingProps) method.invoke(shape);
        } catch (Exception e) {
            throw new ReflectionException("getCNvPr", XSLFShape.class, e);
        }
    }

}
