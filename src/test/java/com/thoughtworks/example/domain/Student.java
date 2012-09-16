package com.thoughtworks.example.domain;

import java.util.Date;

import com.thoughtworks.excelparser.annotations.ExcelField;
import com.thoughtworks.excelparser.annotations.ExcelObject;
import com.thoughtworks.excelparser.annotations.ParseType;

@ExcelObject(parseType = ParseType.ROW, start = 6, end = 8)
public class Student {

	@ExcelField(position = 2)
	private Long roleNumber;
	
	@ExcelField(position = 3)
	private String name;
	
	@ExcelField(position = 4)
	private Date dateOfBirth;
	
	@ExcelField(position = 5)
	private String fatherName;
	
	@ExcelField(position = 6)
	private String motherName;
	
	@ExcelField(position = 7)
	private String address;
	
	@ExcelField(position = 8)
	private Double totalScore;

	public Long getRoleNumber() {
    	return roleNumber;
    }

	public String getName() {
    	return name;
    }

	public Date getDateOfBirth() {
    	return dateOfBirth;
    }

	public String getFatherName() {
    	return fatherName;
    }

	public String getMotherName() {
    	return motherName;
    }

	public String getAddress() {
    	return address;
    }

	public Double getTotalScore() {
    	return totalScore;
    }

	
}
