package com.exportapp.ppt.entity;

import lombok.Getter;
import lombok.Setter;
import org.springframework.data.jpa.domain.AbstractPersistable;

import javax.persistence.Entity;

@Entity
@Getter
@Setter
public class ExcelSchema extends AbstractPersistable<Long> {
    private String question;
    private int code;
    private String answer;
}
