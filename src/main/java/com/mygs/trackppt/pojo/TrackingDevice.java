package com.mygs.trackppt.pojo;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;


@NoArgsConstructor
@AllArgsConstructor
@Data
public class TrackingDevice {
    //设备名称
    private String deviceName;

    //开始时间
    private Double relativeStartTime;

    //结束时间
    private Double relativeEndTime;
}
