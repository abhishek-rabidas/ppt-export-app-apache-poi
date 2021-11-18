package com.exportapp.ppt.dto;

import lombok.Getter;

@Getter
public class PPTImportBody {
    private String[] patterns;
    private int collectionId;
}
