/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.phucdk.lichhoc.object;

import java.util.ArrayList;
import java.util.List;

/**
 *
 * @author Administrator
 */
public class GeneralData {

    private List<Teacher> listTeachers;
    private List<LectureSchedule> listLectureSchedules;

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
    }

}
