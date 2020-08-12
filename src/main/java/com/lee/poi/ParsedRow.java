package com.lee.poi;

import lombok.Data;

import java.util.HashMap;

@Data
public class ParsedRow {
    private Long rowNum;
    private HashMap data = new HashMap();
}
