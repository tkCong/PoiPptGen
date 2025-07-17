package com.mygs.trackppt.utils;

import com.mygs.trackppt.constant.ChartData;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.ResourceUtils;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Arrays;
import java.util.List;
import java.util.Random;
import java.util.stream.Collectors;

/**
 * PPT散点图生成工具类
 * 用于生成包含设备跟踪方位俯仰角曲线的PPT图表
 *
 * @author z
 * @version 1.0
 * @since 2025
 */
@Slf4j
public class LineChartPptUtil {

    private static final Logger logger = LoggerFactory.getLogger(LineChartPptUtil.class);

    /**
     * 生成PPT图表文件
     *
     * @param templateFilePath 模板文件路径
     * @param outputFilePath   输出文件路径
     * @param pageNumber       要修改的幻灯片页码 (从1开始)
     * @param chartTitle       图表标题
     * @return 是否生成成功
     */
    public static boolean generatePPTChart(String templateFilePath, String outputFilePath,
                                           Integer pageNumber, String chartTitle,List<List<Double>> dataList) {
        try {
            // 从资源路径加载PPT模板文件
            InputStream fis = ResourceUtils.getURL(templateFilePath).openStream();
            // 创建XMLSlideShow对象，表示一个PPT演示文稿
            XMLSlideShow ppt = new XMLSlideShow(fis);
            // 关闭输入流
            fis.close();

            // 调用makePPT方法生成PPT内容
            makePPT(pageNumber, ppt, chartTitle,dataList);

            // 修改完 ppt 后保存
            try (FileOutputStream out = new FileOutputStream(outputFilePath)) {
                // 将修改后的PPT写入输出流
                ppt.write(out);
                // 刷新输出流
                out.flush();
            }

            logger.info("PPT生成成功！文件路径：{}", outputFilePath);
            return true;

        } catch (Exception e) {
            logger.error("PPT生成失败", e);
            return false;
        }
    }

    /**
     * 生成PPT的主方法
     *
     * @param page       要修改的幻灯片页码 (从1开始)
     * @param ppt        XMLSlideShow对象
     * @param chartTitle 图表标题
     */
    private static void makePPT(Integer page, XMLSlideShow ppt, String chartTitle,List<List<Double>> dataList) {
        // 获取指定页码的幻灯片
        XSLFSlide slide = ppt.getSlides().get(page - 1);
        // 填充图表数据到PPT
        fillChartToPPT(dataList, slide, chartTitle);
    }

    /**
     * 生成PPT的主方法（使用自定义数据）
     *
     * @param page       要修改的幻灯片页码 (从1开始)
     * @param ppt        XMLSlideShow对象
     * @param chartTitle 图表标题
     * @param customData 自定义数据
     */
    private void makePPTWithData(Integer page, XMLSlideShow ppt, String chartTitle, double[][] customData) {
        // 获取指定页码的幻灯片
        XSLFSlide slide = ppt.getSlides().get(page - 1);
        // 将二维数组转换为List<List<Double>>，以便于处理
        List<List<Double>> dataList = Arrays.stream(customData)
                .map(row -> Arrays.stream(row).boxed().collect(Collectors.toList()))
                .collect(Collectors.toList());

        // 填充图表数据到PPT
        fillChartToPPT(dataList, slide, chartTitle);
    }

    /**
     * 生成随机折线图数据
     *
     * @return 随机生成的二维数组，每个内部数组代表一条折线的数据点
     */
    public static double[][] generateRandomLineData() {
        Random rand = new Random();

        // 固定生成5条折线 (实际值，未完全遵循MAX_LINE_COUNT)
        int lineCount = rand.nextInt(ChartData.MAX_LINE_COUNT) + 1;
        //int lineCount = 8;
        double[][] result = new double[lineCount][];

        // 遍历生成每条折线的数据
        for (int i = 0; i < lineCount; i++) {
            // 每条折线的长度随机，0到MAX_LINE_LENGTH之间
            int length = rand.nextInt(ChartData.MAX_LINE_LENGTH + 1) + 2;
            double[] line = new double[length];
            // 填充折线数据点，取值范围在MIN_VALUE和MAX_VALUE之间
            for (int j = 0; j < length; j++) {
                line[j] = ChartData.MIN_VALUE + (ChartData.MAX_VALUE - ChartData.MIN_VALUE) * rand.nextDouble();
            }
            result[i] = line;
        }

        return result;
    }

    /**
     * 填充图表数据（增强版）
     * 该方法遍历幻灯片上的所有形状，找到图表，然后用传入的数据填充图表，并设置图表标题和坐标轴格式。
     *
     * @param list       列表 - 二维数组，每个内部列表代表一个数据系列
     * @param slide      幻灯片
     * @param chartTitle 图表标题
     */
    private static void fillChartToPPT(List<List<Double>> list, XSLFSlide slide, String chartTitle) {
        logger.info("开始填充图表数据...");

        // 遍历幻灯片上的所有形状
        for (XSLFShape shape : slide.getShapes()) {
            logger.debug("遍历形状: {}", shape.getClass().getSimpleName());

            // 检查形状是否是图形框架 (图表通常嵌入在图形框架中)
            if (shape instanceof XSLFGraphicFrame) {
                logger.info("找到图表框架");
                XSLFGraphicFrame graphicFrame = (XSLFGraphicFrame) shape;
                // 获取图形框架中的图表对象
                XSLFChart chart = graphicFrame.getChart();

                // 如果找到了图表对象
                if (chart != null) {
                    logger.info("获取到图表对象");
                    try {
                        // 获取图表中的Excel工作簿，图表数据存储在嵌入的Excel中
                        XSSFWorkbook workbook = chart.getWorkbook();
                        // 获取工作簿的第一个工作表
                        XSSFSheet sheet = workbook.getSheetAt(0);

                        logger.info("获取到工作簿和工作表");

                        // 检查输入数据是否为空
                        if (list == null || list.isEmpty()) {
                            logger.warn("警告：输入数据为空");
                            return;
                        }

                        // 找到所有数据系列中的最大行数，即最长的数据系列长度
                        int maxRows = 0;
                        for (List<Double> series : list) {
                            if (series.size() > maxRows) {
                                maxRows = series.size();
                            }
                        }

                        logger.info("数据系列数量: {}, 最大行数: {}", list.size(), maxRows);

                        // 清空现有数据并重新创建
                        // 移除现有数据的标题行（假设第一行是标题行）
                        sheet.removeRow(sheet.getRow(0));

                        // 创建新的标题行
                        XSSFRow headerRow = sheet.createRow(0);
                        // 创建第一个单元格作为X轴标题
                        XSSFCell xCell = headerRow.createCell(0);
                        xCell.setCellValue("X 值");

                        // 为每个数据系列创建列名 (从第二列开始)
                        Random rand = new Random();
                        for (int i = 0; i < list.size(); i++) {
                            XSSFCell cell = headerRow.createCell(i + 1);
                            cell.setCellValue(ChartData.AEROSPACE_TRACKING_TERMS[rand.nextInt(ChartData.AEROSPACE_TRACKING_TERMS.length)]);
                        }

                        logger.info("创建了标题行，包含{}列", list.size() + 1);

                        // 填充数据行
                        for (int row = 0; row < maxRows; row++) {
                            // 创建数据行，从第二行开始 (因为第一行是标题)
                            XSSFRow dataRow = sheet.createRow(row + 1);

                            // 第一列：X值 (从0开始，例如0, 1, 2...)
                            XSSFCell xValueCell = dataRow.createCell(0);
                            xValueCell.setCellValue(row);

                            // 从第二列开始填充每个数据系列的数据
                            for (int col = 0; col < list.size(); col++) {
                                XSSFCell cell = dataRow.createCell(col + 1);
                                // 如果当前行索引小于当前数据系列的长度，则填充实际数据
                                if (row < list.get(col).size()) {
                                    cell.setCellValue(list.get(col).get(row));
                                } else {
                                    // 否则，填充0.0以补充缺失数据，确保图表正确绘制
                                    cell.setCellValue(0.0);
                                }
                            }
                        }

                        logger.info("填充了{}行数据", maxRows);

                        // 强制Excel工作簿重新计算公式，确保图表数据更新
                        sheet.setForceFormulaRecalculation(true);
                        workbook.setForceFormulaRecalculation(true);

                        // 获取图表数据 (通常一个图表只有一个XDDFChartData对象)
                        List<XDDFChartData> chartDataList = chart.getChartSeries();
                        if (chartDataList.isEmpty()) {
                            logger.warn("警告：图表中没有数据系列");
                            return;
                        }

                        // 获取第一个图表数据对象
                        XDDFChartData xddfChartData = chartDataList.get(0);
                        logger.info("图表类型: {}", xddfChartData.getClass().getSimpleName());

                        // 设置数据源范围
                        // X值数据源：从Excel工作表的第2行到第maxRows+1行，第1列 (索引0)
                        XDDFDataSource<Double> xValues = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                                new CellRangeAddress(1, maxRows, 0, 0));

                        logger.info("设置X值范围: 行({},{}), 列(0,0)", 1, maxRows);

                        // 获取现有系列数量
                        int existingSeriesCount = xddfChartData.getSeriesCount();
                        logger.info("现有系列数量: {}", existingSeriesCount);
                        logger.info("需要的系列数量: {}", list.size());

                        // 清除所有现有系列，以便重新添加
                        while (xddfChartData.getSeriesCount() > 0) {
                            xddfChartData.removeSeries(0);
                        }
                        logger.info("清除了所有现有系列");

                        // 重新添加所有数据系列
                        for (int i = 0; i < list.size(); i++) {
                            // 每个系列的Y值数据源：从Excel工作表的第2行到第maxRows+1行，第i+2列 (索引i+1)
                            XDDFNumericalDataSource yValues = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                                    new CellRangeAddress(1, maxRows, i + 1, i + 1));

                            // 添加新系列
                            XDDFChartData.Series newSeries = xddfChartData.addSeries(xValues, yValues);
                            // 设置系列标题
                            rand = new Random();

                            newSeries.setTitle(ChartData.AEROSPACE_TRACKING_TERMS[rand.nextInt(ChartData.AEROSPACE_TRACKING_TERMS.length)], null);

                            logger.info("添加系列{}, Y值范围: 行({},{}), 列({},{})",
                                    i + 1, 1, maxRows, i + 1, i + 1);
                        }

                        // 设置图表标题
                        if (chartTitle != null && !chartTitle.trim().isEmpty()) {
                            try {
                                chart.setTitleText(chartTitle);
                                logger.info("设置图表标题: {}", chartTitle);
                            } catch (Exception titleException) {
                                logger.warn("设置图表标题时出错: {}", titleException.getMessage());
                            }
                        }

                        // 设置坐标轴格式为°
                        /*try {
                            // 通过直接操作图表XML来设置坐标轴格式
                            setAxisFormatAlternative(chart);
                        } catch (Exception axisException) {
                            logger.warn("设置坐标轴格式时出错: {}", axisException.getMessage());
                        }*/

                        // 重新绘图，使更改生效
                        chart.plot(xddfChartData);
                        logger.info("重新绘制图表");
                    } catch (Exception e) {
                        logger.error("填充图表数据时出错", e);
                    }
                }
            }
        }
    }

    /**
     * 【备选方案】通过直接操作图表XML来设置坐标轴格式
     * 此方法直接访问POI底层XML对象，可以进行更精细的控制，但通常不如XDDFAPI直观。
     *
     * @param chart XSLFChart对象
     */
    private static void setAxisFormatAlternative(XSLFChart chart) {
        logger.info("开始设置坐标轴格式（备选方案）...");

        try {
            // 获取图表的底层XML文档对象
            org.openxmlformats.schemas.drawingml.x2006.chart.CTChart ctChart = chart.getCTChart();

            // 获取绘图区域
            org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea plotArea = ctChart.getPlotArea();

            // 查找值轴（Y轴） - 假设Y轴是第二个值轴 (索引1)
            if (plotArea.getValAxArray() != null && plotArea.getValAxArray().length > 0) {
                org.openxmlformats.schemas.drawingml.x2006.chart.CTValAx valAx = plotArea.getValAxArray(1);

                // 设置数字格式
                if (valAx.getNumFmt() == null) {
                    valAx.addNewNumFmt(); // 如果没有数字格式，则添加一个新的
                }
                valAx.getNumFmt().setFormatCode("0.0\"°\""); // 设置格式代码，例如"0.0°"
                valAx.getNumFmt().setSourceLinked(false); // 不链接到源数据格式

                logger.info("通过直接XML操作设置Y轴格式成功");
            }

        } catch (Exception e) {
            logger.warn("备选方案设置失败: {}", e.getMessage());
        }
    }
}