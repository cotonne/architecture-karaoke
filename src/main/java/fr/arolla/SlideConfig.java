package fr.arolla;

import java.nio.file.Path;

public class SlideConfig {
    private final String title;
    private final String leftText;
    private final String rightText;
    private final Path img;

    public SlideConfig(String title, String leftText, String rightText, Path image) {
        this.title = title;
        this.leftText = leftText;
        this.rightText = rightText;
        this.img = image;
    }

    public String getTitle() {
        return title;
    }

    public String getLeftText() {
        return leftText;
    }

    public String getRightText() {
        return rightText;
    }

    public Path getImg() {
        return img;
    }
}
