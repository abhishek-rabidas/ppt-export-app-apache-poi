package com.exportapp.ppt.controllers;

import com.exportapp.ppt.services.Helper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

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
}
