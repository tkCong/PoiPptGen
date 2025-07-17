package com.mygs.trackppt.utils;

import com.mygs.trackppt.constant.ChartData;
import com.mygs.trackppt.pojo.GanttChartData;
import com.mygs.trackppt.pojo.TrackingDevice;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.xslf.usermodel.*;
import org.springframework.util.ResourceUtils;

import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.List;
import java.util.concurrent.ThreadLocalRandom;

/**
 * 工具类：用于生成甘特图 PPT 演示文稿
 *
 * @author z
 * @since 1.0.0
 */
@Slf4j
public class GanttChartPptUtil {

    /**
     * 生成包含甘特图的 PPT 文件
     *
     * @param templateFilePath PPT 模板文件路径
     * @param outputFilePath   输出文件路径
     */
    public static void generatePPTChart(String templateFilePath, String outputFilePath) {
        try {
            InputStream fis = ResourceUtils.getURL(templateFilePath).openStream();
            XMLSlideShow ppt = new XMLSlideShow(fis);
            fis.close();

            List<TrackingDevice> trackingDevices = generateTrackingDevices(10);
            GanttChartData ganttChartData = new GanttChartData("示例甘特图", trackingDevices);

            generateGanttChart(ppt, ganttChartData, 1);

            try (FileOutputStream out = new FileOutputStream(outputFilePath)) {
                ppt.write(out);
                out.flush();
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 在指定幻灯片上生成甘特图
     *
     * @param ppt            PPT 文档对象
     * @param ganttChartData 甘特图数据对象
     * @param page           页码（从1开始）
     * @throws Exception 异常处理
     */
    public static void generateGanttChart(XMLSlideShow ppt, GanttChartData ganttChartData, int page) throws Exception {
        int slideTotalWidth = 1280;
        int slideWidth = 1080;
        int horizontalOffset = (slideTotalWidth - slideWidth) / 2;
        int leftMargin = 80 + horizontalOffset;
        int rightMargin = 40 + horizontalOffset;
        int initialSlideHeight = 270;
        int topMargin = 50;
        int bottomMargin = 0;

        int chartWidth = slideWidth - leftMargin - rightMargin;
        int chartHeight = initialSlideHeight - topMargin - bottomMargin;

        double maxEndTime = 0;
        for (TrackingDevice d : ganttChartData.getDeviceList()) {
            maxEndTime = Math.max(maxEndTime, d.getRelativeEndTime());
        }
        if (maxEndTime == 0) maxEndTime = 60;

        double pixelsPerSecond = (double) chartWidth / maxEndTime;

        Map<String, Integer> deviceYMap = new LinkedHashMap<>();
        Map<String, Color> deviceColorMap = new HashMap<>();
        int colorIndex = 0;
        for (TrackingDevice d : ganttChartData.getDeviceList()) {
            if (!deviceYMap.containsKey(d.getDeviceName())) {
                deviceYMap.put(d.getDeviceName(), deviceYMap.size());
                deviceColorMap.put(d.getDeviceName(), ChartData.DEVICE_COLORS[colorIndex % ChartData.DEVICE_COLORS.length]);
                colorIndex++;
            }
        }

        int deviceCount = Math.max(deviceYMap.size(), 1);
        int axisAreaHeight = 30;
        int availableHeight = chartHeight - axisAreaHeight;
        int rowHeight = availableHeight / deviceCount;
        int barPadding = 4;
        int barHeight = Math.min(4, rowHeight - barPadding);

        XSLFSlide slide = ppt.getSlides().get(page - 1);

        // 添加标题
        if (ganttChartData.getTitle() != null && !ganttChartData.getTitle().isEmpty()) {
            XSLFTextShape title = slide.createTextBox();
            title.setAnchor(new Rectangle(0, 20, slideWidth, 50));
            title.setText(ganttChartData.getTitle());
            title.setFillColor(null);
            title.setLineColor(null);
            XSLFTextParagraph para = title.getTextParagraphs().get(0);
            para.setTextAlign(TextParagraph.TextAlign.CENTER);
            XSLFTextRun run = para.getTextRuns().get(0);
            run.setFontSize(18.0);
            run.setBold(true);
            run.setFontColor(Color.BLACK);
        }

        // 绘制Y轴
        int yAxisX = leftMargin;
        int yAxisTop = topMargin;
        int yAxisBottom = topMargin + chartHeight - axisAreaHeight;

        XSLFAutoShape yAxis = slide.createAutoShape();
        yAxis.setShapeType(ShapeType.LINE);
        yAxis.setAnchor(new Rectangle(yAxisX, yAxisTop, 0, yAxisBottom - yAxisTop));
        yAxis.setLineColor(Color.BLACK);
        yAxis.setLineWidth(2.0);

        // 绘制X轴
        int xAxisY = yAxisBottom;
        int xAxisLeft = leftMargin;
        int xAxisRight = leftMargin + chartWidth;

        XSLFAutoShape xAxis = slide.createAutoShape();
        xAxis.setShapeType(ShapeType.LINE);
        xAxis.setAnchor(new Rectangle(xAxisLeft, xAxisY, xAxisRight - xAxisLeft, 0));
        xAxis.setLineColor(Color.BLACK);
        xAxis.setLineWidth(2.0);

        // 绘制设备标签
        for (Map.Entry<String, Integer> entry : deviceYMap.entrySet()) {
            String device = entry.getKey();
            int deviceIndex = entry.getValue();
            int y = topMargin + deviceIndex * rowHeight + (rowHeight - 30) / 2;

            XSLFTextShape deviceLabel = slide.createTextBox();
            deviceLabel.setAnchor(new Rectangle(10, y, leftMargin - 20, 30));
            deviceLabel.setText(device);
            deviceLabel.setFillColor(null);
            deviceLabel.setLineColor(null);
            XSLFTextParagraph para = deviceLabel.getTextParagraphs().get(0);
            para.setTextAlign(TextParagraph.TextAlign.RIGHT);
            XSLFTextRun run = para.getTextRuns().get(0);
            run.setFontSize(device.length() > 9 ? 14.0 - device.length() + 9 : 14.0);
            run.setFontColor(Color.BLACK);
        }

        // 绘制X轴刻度
        int tickInterval = calculateTickInterval((int) maxEndTime);
        for (int seconds = 0; seconds <= maxEndTime; seconds += tickInterval) {
            int x = leftMargin + (int) (seconds * pixelsPerSecond);

            XSLFAutoShape tick = slide.createAutoShape();
            tick.setShapeType(ShapeType.LINE);
            tick.setAnchor(new Rectangle(x, xAxisY, 0, 10));
            tick.setLineColor(Color.BLACK);
            tick.setLineWidth(1.0);

            XSLFTextShape tickLabel = slide.createTextBox();
            tickLabel.setAnchor(new Rectangle(x - 20, xAxisY + 15, 40, 25));
            tickLabel.setText(String.valueOf(seconds));
            tickLabel.setFillColor(null);
            tickLabel.setLineColor(null);
            XSLFTextParagraph para = tickLabel.getTextParagraphs().get(0);
            para.setTextAlign(TextParagraph.TextAlign.CENTER);
            XSLFTextRun run = para.getTextRuns().get(0);
            run.setFontSize(12.0);
            run.setFontColor(Color.BLACK);
        }

        // 绘制任务条
        for (TrackingDevice d : ganttChartData.getDeviceList()) {
            int deviceIndex = deviceYMap.get(d.getDeviceName());
            int y = topMargin + deviceIndex * rowHeight + (rowHeight - barHeight) / 2;
            double duration = d.getRelativeEndTime() - d.getRelativeStartTime();
            int x = leftMargin + (int) (d.getRelativeStartTime() * pixelsPerSecond);
            int width = Math.max(10, (int) (duration * pixelsPerSecond));
            Color barColor = deviceColorMap.get(d.getDeviceName());
            createRoundedRectangle(slide, x, y, width, barHeight, barColor);
        }
    }

    /**
     * 创建带有圆角的矩形任务条
     */
    private static void createRoundedRectangle(XSLFSlide slide, int x, int y, int width, int height, Color color) {
        if (width < height) {
            XSLFAutoShape circle = slide.createAutoShape();
            circle.setShapeType(ShapeType.ELLIPSE);
            circle.setAnchor(new Rectangle(x, y, height, height));
            circle.setFillColor(color);
            circle.setLineColor(color);
            return;
        }

        int radius = height / 2;

        XSLFAutoShape mainRect = slide.createAutoShape();
        mainRect.setShapeType(ShapeType.RECT);
        mainRect.setAnchor(new Rectangle(x + radius, y, width - 2 * radius, height));
        mainRect.setFillColor(color);
        mainRect.setLineColor(color);

        XSLFAutoShape leftCircle = slide.createAutoShape();
        leftCircle.setShapeType(ShapeType.ELLIPSE);
        leftCircle.setAnchor(new Rectangle(x, y, height, height));
        leftCircle.setFillColor(color);
        leftCircle.setLineColor(color);

        XSLFAutoShape rightCircle = slide.createAutoShape();
        rightCircle.setShapeType(ShapeType.ELLIPSE);
        rightCircle.setAnchor(new Rectangle(x + width - height, y, height, height));
        rightCircle.setFillColor(color);
        rightCircle.setLineColor(color);
    }

    /**
     * 计算刻度间隔值
     */
    private static int calculateTickInterval(int maxTime) {
        if (maxTime == 0) return 10;
        int[] intervals = {1, 2, 5, 10, 15, 20, 30, 60, 120, 300, 600, 1200, 1800, 3600};
        int targetTicks = 10;
        for (int interval : intervals) {
            if (maxTime / interval <= targetTicks) {
                return interval;
            }
        }
        int fallback = maxTime / targetTicks;
        int magnitude = (int) Math.pow(10, Math.floor(Math.log10(fallback)));
        return ((fallback / magnitude) + 1) * magnitude;
    }

    /**
     * 随机生成跟踪设备列表
     */
    public static List<TrackingDevice> generateTrackingDevices(int count) {
        List<TrackingDevice> result = new ArrayList<>();
        Set<String> used = new HashSet<>();
        Random rand = new Random();

        for (int i = 0; i < count; i++) {
            String name;
            do {
                name = ChartData.AEROSPACE_TRACKING_TERMS[rand.nextInt(ChartData.AEROSPACE_TRACKING_TERMS.length)];
            } while (used.contains(name));
            used.add(name);

            double start = ThreadLocalRandom.current().nextDouble(0, 60);
            double duration = ThreadLocalRandom.current().nextDouble(5, 30);
            double end = Math.min(start + duration, 100);

            result.add(new TrackingDevice(name, start, end));
        }

        return result;
    }
}
