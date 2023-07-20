package com.googlesheets.demo.Entity;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.springframework.boot.autoconfigure.domain.EntityScan;

import javax.persistence.*;

@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
@Entity
public class Person {
    @Id
    @Column(name = "id")
    private  Double  id;

    @Column(name = "lastName")
    private String lastName;

    @Column(name = "firstName")
    private String firstName;

    @Column(name = "gender")
    private String gender;

    @Column(name = "implementationZone")
    private String implementationZone;

    @Column(name = "title")
    private String title;

    @Column(name = "localeState")
    private String localeState;

    @Column(name = "age")
    private int age;

    @Column(name = "educationLevel")
    private String educationLevel;

    @Column(name = "degreeSpecialty")
    private String degreeSpecialty;

    @Column(name = "currentStatus")
    private String currentStatus;

    @Column(name = "currentLegalStatus")
    private String currentLegalStatus;

    @Column(name = "projectDescription")
    private String projectDescription;

    @Column(name = "projectState")
    private String projectState;

    @Column(name = "hasReceivedFunding")
    private boolean hasReceivedFunding;

    @Column(name = "currentHR")
    private double currentHR;

    @Column(name = "projectedHR")
    private double projectedHR;

    @Column(name = "region")
    private String region;


}
