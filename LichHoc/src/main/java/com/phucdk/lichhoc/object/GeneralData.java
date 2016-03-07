/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.phucdk.lichhoc.object;

import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 *
 * @author Administrator
 */
public class GeneralData {

    private List<Teacher> listTeachers;
    private List<LectureSchedule> listLectureSchedules;
    private Date startDateOfWeek;
    private Map<Teacher, List<String>> mapTeacherTimes;

    private List<Teacher> listBusyTeachers;
    private List<BusySchedule> listBusySchedules;

    public List<Teacher> getListTeachers() {
        return listTeachers;
    }

    public void setListTeachers(List<Teacher> listTeachers) {
        this.listTeachers = listTeachers;
    }

    public List<LectureSchedule> getListLectureSchedules() {
        return listLectureSchedules;
    }

    public void setListLectureSchedules(List<LectureSchedule> listLectureSchedules) {
        this.listLectureSchedules = listLectureSchedules;
    }

    public GeneralData() {
        this.listTeachers = new ArrayList<>();
        this.listLectureSchedules = new ArrayList<>();
        this.listBusyTeachers = new ArrayList<>();
        this.listBusySchedules = new ArrayList<>();
        this.mapTeacherTimes = new HashMap<>();
    }

    public Date getStartDateOfWeek() {
        return startDateOfWeek;
    }

    public void setStartDateOfWeek(Date startDateOfWeek) {
        this.startDateOfWeek = startDateOfWeek;
    }

    public List<Teacher> getListBusyTeachers() {
        return listBusyTeachers;
    }

    public void setListBusyTeachers(List<Teacher> listBusyTeachers) {
        this.listBusyTeachers = listBusyTeachers;
    }

    public List<BusySchedule> getListBusySchedules() {
        return listBusySchedules;
    }

    public void setListBusySchedules(List<BusySchedule> listBusySchedules) {
        this.listBusySchedules = listBusySchedules;
    }

    public Map<Teacher, List<String>> getMapTeacherTimes() {
        return mapTeacherTimes;
    }

    public void setMapTeacherTimes(Map<Teacher, List<String>> mapTeacherTimes) {
        this.mapTeacherTimes = mapTeacherTimes;
    }
    
}
