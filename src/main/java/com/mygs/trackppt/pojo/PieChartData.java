package com.mygs.trackppt.pojo;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Map;

/**
 * 生成饼图所需数据
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class PieChartData {
    // 饼图名称
    private String title;

    private Map<String, Double> amountList;
}
