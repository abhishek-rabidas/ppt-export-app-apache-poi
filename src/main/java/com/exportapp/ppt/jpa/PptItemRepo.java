package com.exportapp.ppt.jpa;

import com.exportapp.ppt.entity.PptItem;
import org.springframework.data.jpa.repository.JpaRepository;

public interface PptItemRepo extends JpaRepository<PptItem, Long> {
}
