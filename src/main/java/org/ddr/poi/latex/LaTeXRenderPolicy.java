package org.ddr.poi.latex;

import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import uk.ac.ed.ph.snuggletex.SnuggleSession;

/**
 * LaTeX字符串渲染策略
 *
 * @author Draco
 * @since 2021-04-14
 */
public class LaTeXRenderPolicy extends AbstractRenderPolicy<String> {
    private SnuggleSession session;

    @Override
    protected boolean validate(String data) {
        if (StringUtils.isBlank(data)) {
            return false;
        }

        // https://www2.ph.ed.ac.uk/snuggletex/documentation/overview-and-features.html
        session = LaTeXUtils.createSession();
        return LaTeXUtils.parse(session, data);
    }

    @Override
    public void doRender(RenderContext<String> context) throws Exception {
        XWPFParagraph paragraph = (XWPFParagraph) context.getRun().getParent();
        CTR ctr = context.getRun().getCTR();
        LaTeXUtils.renderTo(paragraph, ctr, session);
    }

    @Override
    protected void afterRender(RenderContext<String> context) {
        clearPlaceholder(context, false);
    }
}
