package com.exportapp.ppt.jpa;

import com.exportapp.ppt.entity.Collection;
import org.springframework.data.jpa.repository.JpaRepository;

public interface CollectionRepo extends JpaRepository<Collection, Long> {
}
