package com.mygs.trackppt.pojo;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;
import java.util.Map;

/**
 * 生成折线图所需数据
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class LineChartData {
    // 折线图名称
    private String title;

    private Map<String, List<Double>> angleList;
}
