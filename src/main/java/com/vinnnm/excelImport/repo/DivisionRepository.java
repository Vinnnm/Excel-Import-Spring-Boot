package com.vinnnm.excelImport.repo;

import com.vinnnm.excelImport.models.Division;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.Optional;
@Repository
public interface DivisionRepository extends JpaRepository<Division, Long> {
    Optional<Division> findByName(String name);
}
