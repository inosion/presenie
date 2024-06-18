package org.apache.poi.xslf;

import java.awt.geom.Rectangle2D;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;

import org.apache.poi.sl.usermodel.PictureData.PictureType;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.junit.Test;

public class TestXSLFBugs2 {
    private static final String SRC_PPT_FN = "source.pptx";
    private static final String DST_PPT_FN = "dest.pptx";
    private static final String CV_PIC_FN = "closeview.jpeg";
    private static final String LV_PIC_FN = "longview.jpeg";

    private static final String LO_NAME = "CloseLong";
    private static final String CV_PH_NAME = "CloseView";
    private static final String LV_PH_NAME = "LongView";
    private static final String LOC_PH_NAME = "Location";

    @Test
    public void bug60552() throws IOException {
        String srcPPTFN = SRC_PPT_FN;
        String cvPicFN = CV_PIC_FN;
        String lvPicFN = LV_PIC_FN;
        String locTxt = "This is where the site is";

        try (FileInputStream spptIS = new FileInputStream(srcPPTFN);
             FileInputStream cvPicIS = new FileInputStream(cvPicFN);
             FileInputStream lvPicIS = new FileInputStream(lvPicFN)) {
            byte[] cvPic = IOUtils.toByteArray(cvPicIS);
            byte[] lvPic = IOUtils.toByteArray(lvPicIS);

            Set<XSLFSlideLayout> visitedLOs = new HashSet<>();
            XMLSlideShow srcPPT = new XMLSlideShow(spptIS);
            XSLFSlide srcSlide = getSrcSlide(srcPPT);

            XMLSlideShow destPPT = new XMLSlideShow();

            XSLFSlide slide1 = createSlide(destPPT, srcSlide, visitedLOs);
            setPic(destPPT, slide1, cvPic, CV_PH_NAME);
            setPic(destPPT, slide1, lvPic, LV_PH_NAME);
            setText(slide1, LOC_PH_NAME, locTxt);

            XSLFSlide slide2 = createSlide(destPPT, srcSlide, visitedLOs);
            setPic(destPPT, slide2, cvPic, CV_PH_NAME);
            setPic(destPPT, slide2, lvPic, LV_PH_NAME);
            setText(slide2, LOC_PH_NAME, locTxt);

            FileOutputStream fos = new FileOutputStream(DST_PPT_FN);
            destPPT.write(fos);
            fos.close();
        }
    }

    private static XSLFSlide getSrcSlide(XMLSlideShow ppt) {
        XSLFSlide slide = null;
        for (XSLFSlide srcSlide : ppt.getSlides()) {
            if(LO_NAME.equalsIgnoreCase(srcSlide.getSlideLayout().getName())) {
                slide = srcSlide;
                break;
            }
        }
        return slide;
    }


    private static XSLFShape getShape(XSLFSlide slide, String shapeName) {
        XSLFShape shape = null;
        Iterator<XSLFShape> sit = slide.getSlideLayout().getShapes().iterator();
        while (sit.hasNext()) {
            XSLFShape lshape = sit.next();
            if (lshape != null && lshape.getShapeName().equalsIgnoreCase(shapeName)) {
                shape = lshape;
                break;
            }
        }
        return shape;
    }

    private static XSLFShape getPlaceholder(XSLFSlide slide, String txtShapeName) {
        for (XSLFTextShape ph : slide.getPlaceholders()) {
            String name = ph.getShapeName();
            if (name.equalsIgnoreCase(txtShapeName)) {
                return ph;
            }
        }
        return null;
    }

    private static void setText(XSLFSlide slide, String txtShapeName, String txtToSet) {
        XSLFShape shape = getPlaceholder(slide, txtShapeName);
        if (shape != null && shape instanceof XSLFTextShape) {
            XSLFTextShape tshape = (XSLFTextShape) shape;
            tshape.setText(txtToSet);
        }
    }

    private static void setPic(XMLSlideShow ppt, XSLFSlide slide, byte[] pic, String shapeName) {
        XSLFShape shape = getShape(slide, shapeName);
        Rectangle2D anchor = shape.getAnchor();
        slide.removeShape(shape);

        XSLFPictureData pd = ppt.addPicture(pic, PictureType.JPEG);
        XSLFPictureShape pics = slide.createPicture(pd);
        pics.setAnchor(anchor);
    }

    public static XSLFSlide createSlide(XMLSlideShow ppt, XSLFSlide srcSlide, Set<XSLFSlideLayout> visitedLOs) {
        XSLFSlide slide = ppt.createSlide();
        XSLFSlideLayout srcLayout = srcSlide.getSlideLayout();
        if (!visitedLOs.contains(srcLayout)) {
            visitedLOs.add(srcLayout);
            XSLFSlideLayout dstLayout = slide.getSlideLayout();
            dstLayout.importContent(srcLayout);
        }
        slide.importContent(srcSlide);
        return slide;
    }
}
