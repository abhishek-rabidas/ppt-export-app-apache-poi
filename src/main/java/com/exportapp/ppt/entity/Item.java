package com.exportapp.ppt.entity;

import lombok.Getter;
import lombok.Setter;
import org.springframework.data.jpa.domain.AbstractPersistable;

import javax.persistence.Entity;
import javax.persistence.ManyToOne;

@Getter
@Setter
@Entity
public class Item extends AbstractPersistable<Long> {
    @ManyToOne
    private Collection collection;
    private String season;
    private String styleNo;
    private String patternDescription;
    private String licenseBrand;
    private String description;
    private String factoryName;
    private String FactoryID;
    private String coo;
    private String port;
    private String hts;
    private String fiberContent;
    private String materialComposition;
    private String construction;
    private String color;
    private String itemSize;
    private String itemWeight;
    private String innerQty;
    private String innerPackDimensions;
    private String packSizeQty;
    private String ctnToFill;
    private String price;
    private String terms;
    private String dutyPercent;
    private String leadTime;
    private String masterCartonDimensions;
    private String grossWeight;
    private String landedCost;
    private String packagingType;
    private String packagingDimension;
    private String packagingCost;
    private String careInstructions;
    private String callouts;
}