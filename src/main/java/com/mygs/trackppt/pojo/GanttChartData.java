package com.mygs.trackppt.pojo;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * 生成gantt图所需数据
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class GanttChartData {
    // 甘特图名称
    private String title;

    private List<TrackingDevice> deviceList;
}
