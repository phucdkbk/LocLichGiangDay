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
public class StringUtils {
    
    
    public static boolean isEmpty(String str){
        if(str==null){
            return true;
        } else {
            if("".equals(str.trim())){
                return true;
            }
        }
        return false;
    }
}
