package com.exportapp.ppt.services;

import com.exportapp.ppt.entity.Collection;
import com.exportapp.ppt.jpa.CollectionRepo;
import org.apache.poi.sl.extractor.SlideShowExtractor;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.Shape;
import org.apache.poi.sl.usermodel.Slide;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.xslf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.List;

@Service
public class Helper {

    @Autowired
    private CollectionRepo collectionRepo;

    public String fetchText(MultipartFile ppt){

        XMLSlideShow slideshow = null;

        try {
            slideshow = new XMLSlideShow(ppt.getInputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }

        SlideShowExtractor<XSLFShape,XSLFTextParagraph> slideShowExtractor
                = new SlideShowExtractor<XSLFShape,XSLFTextParagraph>(slideshow);

        for (XSLFSlide slide: slideshow.getSlides()){
            for (XSLFShape shape:slide.getShapes()){

                //for images
                if (shape instanceof XSLFPictureShape) {
                    XSLFPictureShape xslfPictureData = (XSLFPictureShape) shape;
                    XSLFPictureData pict = xslfPictureData.getPictureData();
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

                //for text
                if (shape instanceof XSLFTextShape) {
                    XSLFTextShape textShape = (XSLFTextShape)shape;
                    String text = textShape.getText();
                    if(!text.contains(":")) break;
                    String[] array = text.split(":");

                    Collection collection = new Collection();
                    try{
                        collection.setStyleNo(array[1].trim());
                        collection.setDescription(array[2].trim());
                        collection.setComp(array[3].trim());
                        collection.setSize(array[4].trim());
                        collection.setWeight(array[5].trim());
                    }catch (ArrayIndexOutOfBoundsException e){
                        e.getLocalizedMessage();
                    }finally {
                        //collectionRepo.save(collection);
                    }
                }
            }
        }

        //extractImage(slideshow);

        //saveCollection(slideShowExtractor.getText());
        return slideShowExtractor.getText();
    }

    private void extractImage(XMLSlideShow ppt){
        List<XSLFPictureData> pdata = ppt.getPictureData();
        for (int i = 0; i < pdata.size(); i++){
            XSLFPictureData pict = pdata.get(i);

            byte[] data = pict.getData();

            PictureData.PictureType type = pict.getType();
            String ext;
            switch (type){
                case JPEG: ext=".jpg"; break;
                case PNG: ext=".png"; break;
                case WMF: ext=".wmf"; break;
                case EMF: ext=".emf"; break;
                case PICT: ext=".pict"; break;
                default: continue;
            }
            FileOutputStream out;
            try {
                out = new FileOutputStream("pict_"+ (i+1) + ext);
                out.write(data);
                out.close();
            } catch (IOException e) {
                e.getLocalizedMessage();
            }
            //TODO: Cloud storage and URI in Database
        }
    }
}
