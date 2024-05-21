package org.ddr.image.heif;

import javax.imageio.ImageReadParam;

public class HeicOnlineParam extends ImageReadParam {

    private Target target = Target.JPG;
    private Integer quality = 85;
    private Integer width;
    private Integer height;

    private Integer dpi;

    public Target getTarget() {
        return target;
    }

    public void setTarget(Target target) {
        this.target = target;
    }

    public Integer getQuality() {
        return quality;
    }

    public void setQuality(Integer quality) {
        this.quality = quality;
    }

    public Integer getWidth() {
        return width;
    }

    public void setWidth(Integer width) {
        this.width = width;
    }

    public Integer getHeight() {
        return height;
    }

    public void setHeight(Integer height) {
        this.height = height;
    }

    public Integer getDpi() {
        return dpi;
    }

    public void setDpi(Integer dpi) {
        this.dpi = dpi;
    }

    public String withId(String fileId) {
        StringBuilder sb = new StringBuilder();
        sb.append("fileid=").append(fileId)
                .append("&target=").append(target.name())
                .append("&width=").append(width == null ? "" : width)
                .append("&height=").append(height == null ? "" : height)
                .append("&dpi=").append(dpi == null ? "" : dpi);
        if (target.getQualityName() != null) {
            sb.append("&").append(target.getQualityName()).append("=").append(quality == null ? "" : quality);
        }
        return sb.toString();
    }

    public enum Target {
        JPG("jpg_quality"),
        PNG(null),
        BMP(null);

        private final String qualityName;

        Target(String qualityName) {
            this.qualityName = qualityName;
        }

        public String getQualityName() {
            return qualityName;
        }
    }
}
