package com.exportapp.ppt.services;

import com.exportapp.ppt.entity.Collection;
import com.exportapp.ppt.jpa.CollectionRepo;
import org.apache.poi.sl.extractor.SlideShowExtractor;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.xslf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

@Service
public class Helper {

    @Autowired
    private CollectionRepo collectionRepo;

    public void extractAllImages(MultipartFile ppt) {

        XMLSlideShow slideshow = null;

        try {
            slideshow = new XMLSlideShow(ppt.getInputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }

        int picId = 0;

        for (XSLFPictureData pictureDatum : slideshow.getPictureData()) {
            byte[] picData = pictureDatum.getData();
            PictureData.PictureType pictureType = pictureDatum.getType();
            String ext;
            switch (pictureType) {
                case JPEG:
                    ext = ".jpg";
                    break;
                case PNG:
                    ext = ".png";
                    break;
                case WMF:
                    ext = ".wmf";
                    break;
                case EMF:
                    ext = ".emf";
                    break;
                case PICT:
                    ext = ".pict";
                    break;
                default:
                    continue;
            }

            FileOutputStream out;
            try {
                out = new FileOutputStream((picId++) + ext);
                out.write(picData);
                out.close();
            } catch (IOException e) {
                e.getLocalizedMessage();
            }
        }
    }

    public void extractAllImagesSlideWise(MultipartFile ppt) {

        XMLSlideShow slideshow = null;

        try {
            slideshow = new XMLSlideShow(ppt.getInputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }

        for (XSLFSlide slide : slideshow.getSlides()) {
            for (XSLFShape shape : slide.getShapes()) {
                if (shape instanceof XSLFPictureShape) {
                    XSLFPictureShape xslfPictureData = (XSLFPictureShape) shape;
                    byte[] data = xslfPictureData.getPictureData().getData();
                    PictureData.PictureType type = xslfPictureData.getPictureData().getType();
                    String ext;
                    switch (type) {
                        case JPEG:
                            ext = ".jpg";
                            break;
                        case PNG:
                            ext = ".png";
                            break;
                        case WMF:
                            ext = ".wmf";
                            break;
                        case EMF:
                            ext = ".emf";
                            break;
                        case PICT:
                            ext = ".pict";
                            break;
                        default:
                            continue;
                    }
                    FileOutputStream out;
                    try {
                        File folder = new File(String.format("%d", slide.getSlideNumber()));
                        folder.mkdir();
                        File pic = new File(String.format("%d/%d%s", slide.getSlideNumber(), shape.getShapeId(), ext));
                        out = new FileOutputStream(pic);
                        out.write(data);
                        out.close();
                    } catch (IOException e) {
                        e.getLocalizedMessage();
                    }
                }
            }
        }
    }

    public void extractMostDominantImageFromSlide(MultipartFile ppt) {
        XMLSlideShow slideshow = null;

        try {
            slideshow = new XMLSlideShow(ppt.getInputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }

        for (XSLFSlide slide : slideshow.getSlides()) {
            List<XSLFPictureShape> xslfPictureShapeList = new ArrayList<>();
            for (XSLFShape shape : slide.getShapes()) {
                if (shape instanceof XSLFPictureShape) {
                    XSLFPictureShape xslfPictureData = (XSLFPictureShape) shape;
                    xslfPictureShapeList.add(xslfPictureData);
                }
            }

            if (xslfPictureShapeList.size() == 0) continue;

            XSLFPictureShape maxImage = xslfPictureShapeList.get(0);
            double area = Double.MIN_VALUE;

            for (XSLFPictureShape currentPic : xslfPictureShapeList) {
                double currentArea = currentPic.getAnchor().getWidth() * currentPic.getAnchor().getHeight();
                if (currentArea > area) {
                    area = currentArea;
                    maxImage = currentPic;
                }
            }

            byte[] data = maxImage.getPictureData().getData();
            PictureData.PictureType type = maxImage.getPictureData().getType();
            String ext;
            switch (type) {
                case JPEG:
                    ext = ".jpg";
                    break;
                case PNG:
                    ext = ".png";
                    break;
                case WMF:
                    ext = ".wmf";
                    break;
                case EMF:
                    ext = ".emf";
                    break;
                case PICT:
                    ext = ".pict";
                    break;
                default:
                    continue;
            }

            FileOutputStream out;

            try {
                File folder = new File(String.format("%d", slide.getSlideNumber()));
                folder.mkdir();
                File pic = new File(String.format("%d/%s%s", slide.getSlideNumber(), "pic", ext));
                out = new FileOutputStream(pic);
                out.write(data);
                out.close();
            } catch (IOException e) {
                e.getLocalizedMessage();
            }
        }
    }

    public String extractAllTextFromSlideshow(MultipartFile ppt) {
        XMLSlideShow slideshow = null;

        try {
            slideshow = new XMLSlideShow(ppt.getInputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }

        SlideShowExtractor<XSLFShape, XSLFTextParagraph> slideShowExtractor
                = new SlideShowExtractor<XSLFShape, XSLFTextParagraph>(slideshow);

        return slideShowExtractor.getText();
    }

    public void extractTextSlideWise(MultipartFile ppt) {
        XMLSlideShow slideshow = null;

        try {
            slideshow = new XMLSlideShow(ppt.getInputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }

        for (XSLFSlide slide : slideshow.getSlides()) {
            for (XSLFShape shape : slide.getShapes()) {
                if (shape instanceof XSLFTextShape) {
                    System.out.println("Slide number: " + slide.getSlideNumber());
                    System.out.println();
                    System.out.println(((XSLFTextShape) shape).getText());
                    System.out.println();
                }
            }
        }
    }
}

/*
    public String fetchText(MultipartFile ppt) {

        XMLSlideShow slideshow = null;

        try {
            slideshow = new XMLSlideShow(ppt.getInputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }

        SlideShowExtractor<XSLFShape, XSLFTextParagraph> slideShowExtractor
                = new SlideShowExtractor<XSLFShape, XSLFTextParagraph>(slideshow);


        for (XSLFSlide slide : slideshow.getSlides()) {

            List<XSLFPictureShape> xslfPictureShapeList = new ArrayList<>();
            for (XSLFShape shape : slide.getShapes()) {
                //for images
                if (shape instanceof XSLFPictureShape) {
                    XSLFPictureShape xslfPictureData = (XSLFPictureShape) shape;
                    xslfPictureShapeList.add(xslfPictureData);
                }
                //for text
                if (shape instanceof XSLFTextShape) {
                    XSLFTextShape textShape = (XSLFTextShape) shape;
                    String text = textShape.getText();
                    if (!text.contains(":")) break;
                    String[] array = text.split(":");

                    Collection collection = new Collection();
                    try {
                        collection.setStyleNo(array[1].trim());
                        collection.setDescription(array[2].trim());
                        collection.setComp(array[3].trim());
                        collection.setSize(array[4].trim());
                        collection.setWeight(array[5].trim());
                    } catch (ArrayIndexOutOfBoundsException e) {
                        e.getLocalizedMessage();
                    } finally {
                        //collectionRepo.save(collection);
                    }
                }
            }

            //storing dominant image
            if (xslfPictureShapeList.size() == 0) continue;
            XSLFPictureShape maxImage = xslfPictureShapeList.get(0);
            double area = Double.MIN_VALUE;
            for (XSLFPictureShape currentPic : xslfPictureShapeList) {
                double currentArea = currentPic.getAnchor().getWidth() * currentPic.getAnchor().getHeight();
                if (currentArea > area) {
                    area = currentArea;
                    maxImage = currentPic;
                }
            }

            byte[] data = maxImage.getPictureData().getData();
            PictureData.PictureType type = maxImage.getPictureData().getType();
            String ext;
            switch (type) {
                case JPEG:
                    ext = ".jpg";
                    break;
                case PNG:
                    ext = ".png";
                    break;
                case WMF:
                    ext = ".wmf";
                    break;
                case EMF:
                    ext = ".emf";
                    break;
                case PICT:
                    ext = ".pict";
                    break;
                default:
                    continue;
            }
            FileOutputStream out;
            try {
                File folder = new File(String.format("%d", slide.getSlideNumber()));
                folder.mkdir();
                File pic = new File(String.format("%d/%s%s", slide.getSlideNumber(), "pic", ext));
                out = new FileOutputStream(pic);
                out.write(data);
                out.close();
            } catch (IOException e) {
                e.getLocalizedMessage();
            }

        }

        //extractImage(slideshow);

        //saveCollection(slideShowExtractor.getText());
        return slideShowExtractor.getText();
    }

    private void extractImage(XMLSlideShow ppt) {
        List<XSLFPictureData> pdata = ppt.getPictureData();
        for (int i = 0; i < pdata.size(); i++) {
            XSLFPictureData pict = pdata.get(i);

            byte[] data = pict.getData();

            PictureData.PictureType type = pict.getType();
            String ext;
            switch (type) {
                case JPEG:
                    ext = ".jpg";
                    break;
                case PNG:
                    ext = ".png";
                    break;
                case WMF:
                    ext = ".wmf";
                    break;
                case EMF:
                    ext = ".emf";
                    break;
                case PICT:
                    ext = ".pict";
                    break;
                default:
                    continue;
            }
            FileOutputStream out;
            try {
                out = new FileOutputStream("pict_" + (i + 1) + ext);
                out.write(data);
                out.close();
            } catch (IOException e) {
                e.getLocalizedMessage();
            }
            //TODO: Cloud storage and URI in Database0
        }
    }
}
*/