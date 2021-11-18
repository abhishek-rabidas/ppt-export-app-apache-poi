package com.exportapp.ppt.jpa;

import com.exportapp.ppt.entity.Item;
import org.springframework.data.jpa.repository.JpaRepository;

public interface ItemRepo extends JpaRepository<Item, Long> {
}
