package com.my.excelinbox.excel;

import org.jetbrains.annotations.NotNull;

import java.util.HashMap;
import java.util.Map;
import java.util.Set;

/**
 * @author fengran
 */
public class SheetHeader {
    private Map<Integer, String> headerName;
    private Map<String, Integer> headerIndex;

    public SheetHeader(@NotNull Map<Integer, String> header) {
        this.headerName = header;

        headerIndex = new HashMap<>();
        header.keySet().forEach(index -> {
            if (index < 0) {
                throw new IndexOutOfBoundsException("column index in sheet must be greater than 0");
            }
            headerIndex.put(header.get(index), index);
        });

        if (headerIndex.keySet().size() != headerName.keySet().size()) {
            throw new IllegalStateException("column name in sheet must be unique");
        }
    }

    public String getColumnName(int index) {
        return headerName.get(index);
    }

    public int getColumnIndex(String columnName) {
        return headerIndex.get(columnName);
    }

    public Set<Integer> getColumnIndexs() {
        return headerName.keySet();
    }

    public int size(){
        return headerName.size();
    }
}
