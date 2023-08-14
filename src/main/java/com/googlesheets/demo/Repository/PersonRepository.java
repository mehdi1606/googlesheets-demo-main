package com.googlesheets.demo.Repository;


import com.googlesheets.demo.Entity.Person;
import org.apache.commons.codec.language.bm.Lang;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.List;


@Repository
public interface PersonRepository extends JpaRepository<Person, Long> {
    List<Person> findAllByOrderByIdAsc();

}
