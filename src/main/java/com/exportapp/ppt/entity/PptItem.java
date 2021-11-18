package com.exportapp.ppt.entity;

import lombok.Getter;
import lombok.Setter;
import org.springframework.data.jpa.domain.AbstractPersistable;

import javax.persistence.Entity;
import javax.persistence.ManyToOne;

@Entity
@Getter
@Setter
public class PptItem extends AbstractPersistable<Long> {
    private String imageURI;
    private String styleNo;
    @ManyToOne
    private Collection collection;
}
