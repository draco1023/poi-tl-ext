package org.ddr.poi.html.tag;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.ddr.image.ImageType;
import org.ddr.poi.html.HtmlConstants;
import org.ddr.poi.html.HtmlRenderContext;
import org.ddr.poi.util.ByteArrayCopyStream;
import org.jsoup.nodes.Element;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;

/**
 * svg标签渲染器
 *
 * @author Draco
 * @since 2022-04-13
 */
public class SvgRenderer extends ImageRenderer {
    private static final Logger log = LoggerFactory.getLogger(SvgRenderer.class);

    private static final String[] TAGS = {HtmlConstants.TAG_SVG};

    /**
     * 开始渲染
     *
     * @param element HTML元素
     * @param context 渲染上下文
     * @return 是否继续渲染子元素
     */
    @Override
    public boolean renderStart(Element element, HtmlRenderContext context) {
        if (!element.hasAttr("xmlns")) {
            element.attr("xmlns", "http://www.w3.org/2000/svg");
        }
        String svg = element.outerHtml().replace(" />", "/>");
        byte[] bytes = svg.getBytes(StandardCharsets.UTF_8);
        BufferedImage image;
        try (InputStream svgStream = new ByteArrayInputStream(bytes)) {
            image = ImageIO.read(svgStream);

            ImageType type = typeOf(image);

            int size = image.getData().getDataBuffer().getSize();
            ByteArrayCopyStream outputStream = new ByteArrayCopyStream(size);
            ImageIO.write(image, type.getExtension(), outputStream);

            InputStream imageStream = outputStream.toInput();
            addPicture(element, context, imageStream, type.getType(), image.getWidth(), image.getHeight(), bytes);
        } catch (IOException | InvalidFormatException e) {
            log.warn("Failed to render svg as image: {}", svg, e);
        } finally {
            // 释放资源
            image = null;
        }
        return false;
    }

    /**
     * @return 支持的HTML标签
     */
    @Override
    public String[] supportedTags() {
        return TAGS;
    }

}
