package org.ddr.poi.html.util;

import com.steadystate.css.dom.CSSStyleDeclarationImpl;
import com.steadystate.css.dom.Property;
import org.w3c.dom.DOMException;

import java.util.Collections;
import java.util.List;

/**
 * 空样式
 *
 * @author Draco
 * @since 2022-10-21
 */
public class EmptyCSSStyle extends CSSStyleDeclarationImpl {
    @Override
    public void setProperties(List<Property> properties) {
        throw new UnsupportedOperationException();
    }

    @Override
    public List<Property> getProperties() {
        return Collections.unmodifiableList(super.getProperties());
    }

    @Override
    public void setProperty(String propertyName, String value, String priority) throws DOMException {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setCssText(String cssText) throws DOMException {
        throw new UnsupportedOperationException();
    }
}
