package com.exportapp.ppt.services;

import com.exportapp.ppt.dto.PPTImportBody;
import com.exportapp.ppt.entity.Collection;
import com.exportapp.ppt.entity.ExcelSchema;
import com.exportapp.ppt.entity.Item;
import com.exportapp.ppt.entity.PptItem;
import com.exportapp.ppt.jpa.CollectionRepo;
import com.exportapp.ppt.jpa.ExcelSchemaRepo;
import com.exportapp.ppt.jpa.ItemRepo;
import com.exportapp.ppt.jpa.PptItemRepo;
import com.sun.istack.NotNull;
import org.apache.poi.sl.extractor.SlideShowExtractor;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Service
public class Helper {

    @Autowired
    private CollectionRepo collectionRepo;

    @Autowired
    private ItemRepo itemRepo;

    @Autowired
    private PptItemRepo pptItemRepo;

    public void extractAllImages(@NotNull MultipartFile ppt) {

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

    public void extractAllImagesSlideWise(@NotNull MultipartFile ppt) {

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

    public void extractMostDominantImageFromSlide(@NotNull MultipartFile ppt) {
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

    public String extractAllTextFromSlideshow(@NotNull MultipartFile ppt) {
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

    public void extractTextSlideWise(@NotNull MultipartFile ppt) {
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

    public List<Item> excelExtraction(MultipartFile xl) {
        try {
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(xl.getInputStream());
            int totalSheets = xssfWorkbook.getNumberOfSheets();

            DataFormatter dataFormatter = new DataFormatter();

            for(int sheetNum=0; sheetNum<totalSheets; sheetNum++){
                XSSFSheet sheet = xssfWorkbook.getSheetAt(sheetNum);
                int rowNum = 0;
                Iterator<Row> rows = sheet.iterator();
                while(rows.hasNext()){
                    Row currentRow = rows.next();
                    if(rowNum == 0){
                        rowNum++;
                        continue;
                    }
                    Iterator<Cell> cells = currentRow.iterator();

                    Item item = new Item();
                    while (cells.hasNext()){
                        Cell currentCell = cells.next();
                        int columnIndex = currentCell.getColumnIndex();
                        switch (columnIndex){
                            case 2:
                                item.setSeason(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 3:
                                item.setStyleNo(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 4:
                                item.setPatternDescription(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 5:
                                item.setLicenseBrand(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 6:
                                item.setDescription(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 7:
                                item.setFactoryName(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 8:
                                item.setFactoryID(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 9:
                                item.setCoo(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 10:
                                item.setPort(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 11:
                                item.setHts(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 12:
                                item.setFiberContent(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 13:
                                item.setMaterialComposition(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 14:
                                item.setConstruction(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 15:
                                item.setColor(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 16:
                                item.setItemSize(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 17:
                                item.setItemWeight(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 18:
                                item.setInnerQty(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 19:
                                item.setInnerPackDimensions(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 20:
                                item.setPackSizeQty(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 21:
                                item.setCtnToFill(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 22:
                                item.setPrice(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 23:
                                item.setTerms(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 24:
                                item.setDutyPercent(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 25:
                                item.setLeadTime(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 26:
                                item.setMasterCartonDimensions(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 27:
                                item.setGrossWeight(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 28:
                                item.setLandedCost(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 29:
                                item.setPackagingType(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 30:
                                item.setPackagingDimension(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 31:
                                item.setPackagingCost(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 32:
                                item.setCareInstructions(dataFormatter.formatCellValue(currentCell));
                                break;
                            case 33:
                                item.setCallouts(dataFormatter.formatCellValue(currentCell));
                                break;
                            default:
                                break;
                        }
                    }
                    itemRepo.save(item);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return itemRepo.findAll();

    }

    public void pptImport(MultipartFile ppt, PPTImportBody pptImportBody){

        String[] patterns = pptImportBody.getPatterns();
        List<String> regexPatterns = new ArrayList<>(patterns.length);

        for(String pattern:patterns) {
            StringBuilder regexPattern = new StringBuilder();
            for (Character character : pattern.toCharArray()) {
                if (Character.isAlphabetic(character)) {
                    regexPattern.append("\\w");
                } else if (Character.isDigit(character)) {
                    regexPattern.append("\\d");
                } else if (Character.isWhitespace(character)) {
                    regexPattern.append("\\s");
                } else if (!Character.isLetterOrDigit(character)) {
                    regexPattern.append("[").append(character).append("]");
                }
            }
            regexPatterns.add(regexPattern.toString());
        }

            XMLSlideShow slideshow = null;

            try {
                slideshow = new XMLSlideShow(ppt.getInputStream());
            } catch (IOException e) {
                e.printStackTrace();
            }

            for (XSLFSlide slide : slideshow.getSlides()) {
                PptItem pptItem = new PptItem();
                List<XSLFPictureShape> xslfPictureShapeList = new ArrayList<>();
                String styleNo = null;
                for (XSLFShape shape : slide.getShapes()) {
                    if (shape instanceof XSLFPictureShape) {
                        XSLFPictureShape xslfPictureData = (XSLFPictureShape) shape;
                        xslfPictureShapeList.add(xslfPictureData);
                    }

                    if(shape instanceof XSLFTextShape){
                        String line = ((XSLFTextShape) shape).getText();
                        for(int itr=0;itr<patterns.length;itr++){
                            Pattern p = Pattern.compile(regexPatterns.get(itr).toString());
                            Matcher matcher = p.matcher(line);
                            while (matcher.find()){
                                styleNo = matcher.group();
                                break;
                            }
                        }
                    }
                }
                if(styleNo == null){
                    continue;
                }else {
                    pptItem.setStyleNo(styleNo);
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
                    File pic = new File(String.format("%d/%s%s", slide.getSlideNumber(), styleNo, ext));
                    out = new FileOutputStream(pic);
                    out.write(data);
                    out.close();
                } catch (IOException e) {
                    e.getLocalizedMessage();
                }finally {
                    pptItem.setImageURI(String.format("%d/%s%s", slide.getSlideNumber(), styleNo, ext));
                }
                pptItemRepo.save(pptItem);
            }
    }
}