package fr.arolla;

import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Collections;
import java.util.List;
import java.util.Random;
import java.util.stream.Stream;

import static java.util.stream.Collectors.joining;
import static java.util.stream.Collectors.toList;

public class EntryPoint {

    private static final String WORDS_S_TXT = "words/%s.txt";
    private static final Random RANDOM = new Random();
    private static final List<Path> IMAGES = listImages();

    private static List<Path> listImages() {
        try {
            return Files.list(Paths.get("imgs/")).collect(toList());
        } catch (IOException e) {
            e.printStackTrace();
            return Collections.emptyList();
        }
    }

    public static void main(String[] args) throws IOException {
        // From http://www.baeldung.com/apache-poi-slideshow
        XMLSlideShow ppt = createNewPresentation();

        XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);
        // Slide : Titre
        titleSlide(ppt, defaultMaster);

        SlideConfig la_reunion = new SlideConfig("La réunion", "", "L'heure est grave!\nIl est temps de renouveller notre architecture!", oneImage());
        createTwoSideSlide(ppt, defaultMaster, la_reunion);

        SlideConfig la_situation_actuelle = new SlideConfig("La situation actuelle", "Voici les principaux composants de notre architecture:\n" + join("old_communication", "old", "shinny_language", "random"), "", oneImage());
        createTwoSideSlide(ppt, defaultMaster, la_situation_actuelle);

        // Notre problème
        SlideConfig le_probleme = new SlideConfig("Le problème!", "", "Parce que nous rencontrons le problème suivant: " + join("problem", "problem"), oneImage());
        createTwoSideSlide(ppt, defaultMaster, le_probleme);

        // Notre solution
        SlideConfig la_solution = new SlideConfig("La solution!", "Le meilleur moyen pour résoudre ça:\n" + join("shinny_language", "shinny_infra", "old"), "", oneImage());
        createTwoSideSlide(ppt, defaultMaster, la_solution);

        SlideConfig le_pouvoir = new SlideConfig("Une solution puissante !", "", "Nous mélangerons:\n" + join("random", "shinny_infra", "old"), oneImage());
        createTwoSideSlide(ppt, defaultMaster, le_pouvoir);

        anyQuestion(ppt, defaultMaster);

        savePresentation(ppt);
    }

    private static String join(String... contexts) {
        return Stream.of(contexts)
                .map(EntryPoint::getOneWordFrom)
                .collect(joining("\n"));
    }

    private static Path oneImage() {
        return IMAGES.get(RANDOM.nextInt(IMAGES.size()));
    }

    private static String getOneWordFrom(String source) {
        try {
            List<String> strings;
            strings = Files.readAllLines(Paths.get(String.format(WORDS_S_TXT, source)));
            int index = RANDOM.nextInt(strings.size());
            return strings.get(index);
        } catch (IOException e) {
            e.printStackTrace();
            return "";
        }
    }

    private static void createTwoSideSlide(XMLSlideShow ppt, XSLFSlideMaster defaultMaster, SlideConfig slideConfig) throws IOException {
        XSLFSlideLayout layout = defaultMaster.getLayout(SlideLayout.TWO_OBJ);
        XSLFSlide slide = ppt.createSlide(layout);

        XSLFTextShape titleShape = slide.getPlaceholder(0);
        titleShape.setText(slideConfig.getTitle());
        XSLFTextShape left = slide.getPlaceholder(1);
        left.setText(slideConfig.getLeftText());

        XSLFTextShape right = slide.getPlaceholder(2);
        right.setText(slideConfig.getRightText());

        byte[] pictureData = IOUtils.toByteArray(new FileInputStream(slideConfig.getImg().toString()));
        XSLFPictureData pd = ppt.addPicture(pictureData, PictureData.PictureType.PNG);
        XSLFPictureShape picture = slide.createPicture(pd);
        int x;
        int y;
        if ("".equals(slideConfig.getLeftText())) {
            x = 20;
            y = 125;
        } else {
            x = 350;
            y = 125;
        }
        picture.setAnchor(new Rectangle(x, y, 300, 375));

    }

    private static void titleSlide(XMLSlideShow ppt, XSLFSlideMaster defaultMaster) {
        XSLFSlideLayout layout
                = defaultMaster.getLayout(SlideLayout.TITLE);
        XSLFSlide slide = ppt.createSlide(layout);
        XSLFTextShape titleShape = slide.getPlaceholder(0);
        titleShape.setText("Meeting\nde la dernière chance");
        XSLFTextShape contentShape = slide.getPlaceholder(1);
        contentShape.setText("Jam de code - 02/08/2018");
    }

    private static XMLSlideShow createNewPresentation() {
        // Create a new presentation
        XMLSlideShow ppt = new XMLSlideShow();
        ppt.createSlide();
        return ppt;
    }

    private static void anyQuestion(XMLSlideShow ppt, XSLFSlideMaster defaultMaster) throws IOException {
        XSLFSlideLayout layout
                = defaultMaster.getLayout(SlideLayout.TITLE_ONLY);
        XSLFSlide slide = ppt.createSlide(layout);
        XSLFTextShape titleShape = slide.getPlaceholder(0);
        titleShape.setText("Merci!\nDes questions?");
        byte[] pictureData = IOUtils.toByteArray(new FileInputStream(oneImage().toString()));
        XSLFPictureData pd = ppt.addPicture(pictureData, PictureData.PictureType.PNG);
        XSLFPictureShape picture = slide.createPicture(pd);
        picture.setAnchor(new Rectangle(20, 125, 600, 375));
    }

    private static void savePresentation(XMLSlideShow ppt) throws IOException {
        FileOutputStream out = new FileOutputStream("powerpoint.pptx");
        ppt.write(out);
        out.close();
    }
}
