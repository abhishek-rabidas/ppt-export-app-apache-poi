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

    @PostMapping("/fetchtext")
    public String fetchText(@RequestParam("file") MultipartFile file){
        return helper.fetchText(file);
    }
}