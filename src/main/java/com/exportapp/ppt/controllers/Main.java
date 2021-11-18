package com.exportapp.ppt.controllers;

import com.exportapp.ppt.dto.PPTImportBody;
import com.exportapp.ppt.entity.ExcelSchema;
import com.exportapp.ppt.entity.Item;
import com.exportapp.ppt.services.Helper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

@RestController
public class Main {

    @Autowired
    Helper helper;

    @PostMapping("/1")
    public void extractAllImages(@RequestParam("file") MultipartFile file){
        helper.extractAllImages(file);
    }

    @PostMapping("/2")
    public void extractAllImagesSlideWise(@RequestParam("file") MultipartFile file){
        helper.extractAllImagesSlideWise(file);
    }

    @PostMapping("/3")
    public void extractMostDominantImageFromSlide(@RequestParam("file") MultipartFile file){
        helper.extractMostDominantImageFromSlide(file);
    }

    @PostMapping("/4")
    public String extractAllTextFromSlideshow(@RequestParam("file") MultipartFile file){
        return helper.extractAllTextFromSlideshow(file);
    }

    @PostMapping("/5")
    public void extractTextSlideWise(@RequestParam("file") MultipartFile file){
        helper.extractTextSlideWise(file);
    }

    @PostMapping("/6")
    public List<Item> excelExtraction(@RequestParam("file") MultipartFile file){
        return helper.excelExtraction(file);
    }

    @PostMapping("/v1")
    public void pptImport(@RequestParam("file") MultipartFile file, @RequestPart("json_body") PPTImportBody pptImportBody){
        System.out.println(pptImportBody.getPatterns().length);
        System.out.println(pptImportBody.getCollectionId());
        helper.pptImport(file, pptImportBody);
    }
}
