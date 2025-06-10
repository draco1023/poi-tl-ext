/*
 * Copyright 2016 - 2021 Draco, https://github.com/draco1023
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

package org.ddr.poi.math;

import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;

/**
 * MathML字符串渲染策略
 *
 * @author Draco
 * @since 2021-04-10 22:51
 */
public class MathMLRenderPolicy extends AbstractRenderPolicy<String> {
    private final MathRenderConfig config;
    private Element math;

    public MathMLRenderPolicy() {
        this(new MathRenderConfig());
    }

    public MathMLRenderPolicy(MathRenderConfig config) {
        this.config = config;
    }

    public MathRenderConfig getConfig() {
        return config;
    }

    @Override
    protected boolean validate(String data) {
        if (StringUtils.isBlank(data)) {
            return false;
        }
        math = Jsoup.parseBodyFragment(data).selectFirst("math");
        return math != null;
    }

    @Override
    public void doRender(RenderContext<String> context) throws Exception {
        if (!math.hasAttr("xmlns")) {
            math.attr("xmlns", "http://www.w3.org/1998/Math/MathML");
        }
        String mathml = math.outerHtml();
        MathMLUtils.renderTo((XWPFParagraph) context.getRun().getParent(), context.getRun().getCTR(), mathml, config);
    }

    @Override
    protected void afterRender(RenderContext<String> context) {
        clearPlaceholder(context, false);
    }
}
