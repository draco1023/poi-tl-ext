package org.ddr.poi.latex;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

class LaTeXRenderPolicyTest {

    @Test
    void doRender() throws IOException {
        LaTeXRenderPolicy laTeXRenderPolicy = new LaTeXRenderPolicy();
        Configure configure = Configure.builder()
                .bind("math1", laTeXRenderPolicy)
                .bind("math2", laTeXRenderPolicy)
                .bind("math3", laTeXRenderPolicy)
                .bind("math4", laTeXRenderPolicy)
                .build();
        Map<String, Object> data = new HashMap<>(4);
        // https://www2.ph.ed.ac.uk/snuggletex/documentation/math-mode.html
        data.put("math1", "$$ x+2=3 $$");
        data.put("math2", "\\[ \\sum_{i=1}^{\\infty} \\frac{1}{n^s} \n" +
                "= \\prod_p \\frac{1}{1 - p^{-s}} \\tag{1.1}\\]");
        data.put("math3", "Product $\\prod_{i=a}^{b} f(i) \\tag{\\textcircled{1}}$ inside text. $\\mathbb{N} \\mathbf{N} \\mathbb{Z} \\mathbf{Z} \\mathbb{D} \\mathbf{D} \\mathbb{Q} \\mathbf{Q} \\mathbb{R} \\mathbf{R} \\mathbb{C} \\mathbf{C}$");
        data.put("math4", "$\\lim_{x\\to\\infty} f(x)$");
        try (InputStream inputStream = LaTeXRenderPolicyTest.class.getResourceAsStream("/math.docx")) {
            XWPFTemplate.compile(inputStream, configure).render(data).writeToFile("latex_out.docx");
        }
    }
}