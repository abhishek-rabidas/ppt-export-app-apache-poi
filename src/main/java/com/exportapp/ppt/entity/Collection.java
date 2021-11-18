package com.exportapp.ppt.entity;

import lombok.Getter;
import lombok.Setter;
import org.springframework.data.jpa.domain.AbstractPersistable;

import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.Id;

@Entity
@Getter
@Setter
public class Collection extends AbstractPersistable<Long> {
    private String collectionName;
    private String styleNo;
    private String description;
    private String comp;
    private String size;
    private String weight;
}
