/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.phucdk.lichhoc.util;

/**
 *
 * @author Administrator
 */
public interface Constants {

    public interface COLUMN {

        public static final int TIME_COLUMN = 2;
        public static final int MONDAY_COLUMN = 9;
        public static final int SUNDAY_COLUMN = 15;
        public static final int CAMPUS_COLUMN = 4;
        public static final int SCHOOL_CLASS_COLUMN = 6;

        public interface INVIDUAL {

            public static final int SUNDAY_COLUMN = 7;
        }
    }

    public interface ROW {

        public static final int START_ROW = 1;
    }

    public interface BUSY_SCHEDULE {

        public interface COLUMN {

            public static final int MONDAY_COLUMN = 1;
            public static final int SUNDAY_COLUMN = 8;
        }

        public interface ROW {

            public static final int START_ROW_FRANCE = 6;
            public static final int START_ROW_VI = 5;
        }
        
        public interface COLOR{
            public static final int NO_FILL = 64;
        }
    }

}
