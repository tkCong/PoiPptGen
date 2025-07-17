package com.mygs.trackppt.constant;

import java.awt.*;

/**
 * 图表数据配置常量类
 * 用于控制图表生成过程中的参数范围、预设值和样式
 */
public class ChartData {

    /** ================= 折线图配置 ================= */

    /** 最大折线条数（1 ~ 10） */
    public static final int MAX_LINE_COUNT = 10;

    /** 每条折线的最大数据点数量（0 ~ N） */
    public static final int MAX_LINE_LENGTH = 10;

    /** 折线图数据的最小值（含） */
    public static final int MIN_VALUE = -2;

    /** 折线图数据的最大值（含） */
    public static final int MAX_VALUE = 2;

    /** 预设折线图颜色方案 */
    public static final Color[] DEVICE_COLORS = {
            new Color(52, 152, 219),   // 蓝色
            new Color(46, 204, 113),   // 绿色
            new Color(241, 196, 15),   // 黄色
            new Color(231, 76, 60),    // 红色
            new Color(155, 89, 182),   // 紫色
            new Color(26, 188, 156),   // 青色
            new Color(230, 126, 34),   // 橙色
            new Color(149, 165, 166)   // 灰色
    };

    /** 折线图/系统监控图例名称（计算机性能指标） */
    public static final String[] AEROSPACE_TRACKING_TERMS = {
            "CPU使用率", "内存占用", "网络延迟", "磁盘读写速率", "线程数", "缓存命中率", "数据吞吐量", "连接数",
            "负载均衡比", "IO等待时间", "CPU上下文切换", "内存碎片率", "GC频率", "JVM堆使用率", "响应时间",
            "系统负载", "磁盘使用率", "TCP连接数", "UDP丢包率", "系统调用频率", "文件句柄数", "线程池活跃数",
            "数据库响应时间", "HTTP请求数", "接口成功率", "服务可用率", "CPU温度", "网络带宽利用率",
            "平均事务耗时", "事务并发数", "内存页交换率", "磁盘IOPS", "缓存大小", "连接建立时间", "DNS解析时间",
            "API错误率", "请求排队长度", "消息队列积压", "心跳丢失次数", "SSL握手时长", "对象创建速率",
            "类加载数量", "JVM非堆内存使用", "资源回收速率", "服务启动时长", "页面加载时间", "WebSocket连接数",
            "数据包重传率", "处理器中断速率"
    };

    /** ================= 饼图配置 ================= */

    /** 饼图最小数据项数（2 ~ 10） */
    public static final int PIE_MIN_PIE_ITEMS = 2;

    /** 饼图最大数据项数（2 ~ 10） */
    public static final int PIE_MAX_PIE_ITEMS = 10;

    /** 饼图数据最小值（含） */
    public static final int PIE_MIN_VALUE = 10;

    /** 饼图数据最大值（含） */
    public static final int PIE_MAX_VALUE = 100;

    // 禁止实例化
    private ChartData() {
        throw new UnsupportedOperationException("This is a constants class and cannot be instantiated.");
    }
}
