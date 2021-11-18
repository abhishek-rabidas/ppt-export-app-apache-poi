package com.exportapp.ppt.jpa;

import com.exportapp.ppt.entity.ExcelSchema;
import org.springframework.data.jpa.repository.JpaRepository;

public interface ExcelSchemaRepo extends JpaRepository<ExcelSchema, Long> {
}
