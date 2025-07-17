package com.mygs.trackppt.utils;

import com.mygs.trackppt.constant.ChartData;
import com.mygs.trackppt.pojo.PieChartData;
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
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;

/**
 * PPT饼图生成工具类
 *
 * @author z
 * @date
 */
public class PieChartPptUtil {
    private static final Logger logger = LoggerFactory.getLogger(PieChartPptUtil.class);
    /**
     * 生成PPT饼图文件
     *
     * @param templateFilePath 模板文件路径
     * @param pieChartData     饼图数据
     * @param pageNumber       要修改的幻灯片页码 (从1开始)
     * @return 是否生成成功
     */
    public static boolean generatePieChartPPT(String templateFilePath, String outputFilePath, PieChartData pieChartData, int pageNumber) {
        try {
            // 从资源路径加载PPT模板文件
            InputStream fis = ResourceUtils.getURL(templateFilePath).openStream();
            // 创建XMLSlideShow对象，表示一个PPT演示文稿
            XMLSlideShow ppt = new XMLSlideShow(fis);
            // 关闭输入流
            fis.close();
            // 调用makePPT方法生成PPT内容
            makePPT(pageNumber, pieChartData, ppt);

            // 修改完 ppt 后保存
            try (FileOutputStream out = new FileOutputStream(outputFilePath)) {
                // 将修改后的PPT写入输出流
                ppt.write(out);
                // 刷新输出流
                out.flush();
            } finally {
                ppt.close();
            }

            return true;

        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
    }

    /**
     * 生成PPT的主方法
     *
     * @param page         要修改的幻灯片页码 (从1开始)
     * @param pieChartData 饼图数据
     * @param ppt          XMLSlideShow对象
     */
    public static void makePPT(Integer page, PieChartData pieChartData, XMLSlideShow ppt) {
        // 获取指定页码的幻灯片
        XSLFSlide slide = ppt.getSlides().get(page - 1);

        // 填充图表数据到PPT
        fillPieChartToPPT(pieChartData.getAmountList(), slide, pieChartData.getTitle());
    }

    /**
     * 生成随机饼图数据
     *
     * @return 随机生成的饼图数据，键为类别名称，值为数值
     */
    public static Map<String, Double> generateRandomPieData() {
        Random rand = new Random();
        Map<String, Double> pieData = new LinkedHashMap<>();
        int itemCount = ChartData.PIE_MIN_PIE_ITEMS + rand.nextInt(ChartData.PIE_MAX_PIE_ITEMS - ChartData.PIE_MIN_PIE_ITEMS + 1);
        // 生成数据项
        for (int i = 0; i < itemCount; i++) {
            String category = ChartData.AEROSPACE_TRACKING_TERMS[rand.nextInt(ChartData.AEROSPACE_TRACKING_TERMS.length)];
            double value = ChartData.PIE_MIN_VALUE + (ChartData.PIE_MAX_VALUE - ChartData.PIE_MIN_VALUE) * rand.nextDouble();
            pieData.put(category, value);
        }

        return pieData;
    }

    /**
     * 填充饼图数据到PPT
     *
     * @param pieData    饼图数据
     * @param slide      幻灯片
     * @param chartTitle 图表标题
     */
    private static void fillPieChartToPPT(Map<String, Double> pieData, XSLFSlide slide, String chartTitle) {
        logger.info("开始填充饼图数据...");

        // 遍历幻灯片上的所有形状
        for (XSLFShape shape : slide.getShapes()) {
            logger.info("遍历形状: " + shape.getClass().getSimpleName());

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
                        if (pieData == null || pieData.isEmpty()) {
                            logger.warn("警告：输入数据为空");
                            return;
                        }

                        // 清空现有数据 - 避免直接删除表格对象（可能导致死循环）
                        // 先清空所有行
                        int lastRowNum = sheet.getLastRowNum();
                        for (int i = lastRowNum; i >= 0; i--) {
                            XSSFRow row = sheet.getRow(i);
                            if (row != null) {
                                sheet.removeRow(row);
                            }
                        }

                        logger.info("清空了现有数据");

                        // 创建标题行（第一行）
                        // 根据Excel表格要求和修复信息，第一列名称不能为空
                        XSSFRow headerRow = sheet.createRow(0);

                        // A1单元格：设置为" "（符合Excel表格列名要求）
                        XSSFCell a1Cell = headerRow.createCell(0);
                        a1Cell.setCellValue(" ");

                        // B1单元格设置饼图名称
                        XSSFCell b1Cell = headerRow.createCell(1);
                        b1Cell.setCellValue(chartTitle);

                        logger.info("创建了标题行: A1= , B1=" + chartTitle);

                        // 填充数据行：从第二行开始
                        int rowIndex = 1;
                        for (Map.Entry<String, Double> entry : pieData.entrySet()) {
                            XSSFRow dataRow = sheet.createRow(rowIndex);

                            // A列：类别名称
                            XSSFCell categoryCell = dataRow.createCell(0);
                            categoryCell.setCellValue(entry.getKey());

                            // B列：数值
                            XSSFCell valueCell = dataRow.createCell(1);
                            valueCell.setCellValue(entry.getValue());

                            logger.info("填充数据行 " + rowIndex + ": " + entry.getKey() + " = " + entry.getValue());
                            rowIndex++;
                        }

                        logger.info("填充了" + pieData.size() + "行数据");

                        // 强制Excel工作簿重新计算公式，确保图表数据更新
                        sheet.setForceFormulaRecalculation(true);
                        workbook.setForceFormulaRecalculation(true);

                        // 获取图表数据
                        List<XDDFChartData> chartDataList = chart.getChartSeries();
                        if (chartDataList.isEmpty()) {
                            logger.warn("警告：图表中没有数据系列");
                            return;
                        }

                        // 获取第一个图表数据对象
                        XDDFChartData xddfChartData = chartDataList.get(0);
                        logger.info("图表类型: " + xddfChartData.getClass().getSimpleName());

                        // 设置饼图数据源
                        // 类别数据源：A列（从第2行开始，即rowIndex=1）
                        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(sheet,
                                new CellRangeAddress(1, pieData.size(), 0, 0));

                        // 数值数据源：B列（从第2行开始，即rowIndex=1）
                        XDDFNumericalDataSource values = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                                new CellRangeAddress(1, pieData.size(), 1, 1));

                        logger.info("设置数据源范围: 类别(1," + pieData.size() + ",0,0), 数值(1," + pieData.size() + ",1,1)");

                        // 清除所有现有系列
                        while (xddfChartData.getSeriesCount() > 0) {
                            xddfChartData.removeSeries(0);
                        }
                        logger.info("清除了所有现有系列");

                        // 添加饼图数据系列
                        XDDFChartData.Series pieSeries = xddfChartData.addSeries(categories, values);
                        pieSeries.setTitle(chartTitle, null);

                        logger.info("添加了饼图数据系列");

                        // 设置图表标题
                        if (chartTitle != null && !chartTitle.trim().isEmpty()) {
                            try {
                                chart.setTitleText(chartTitle);
                                logger.info("设置图表标题: " + chartTitle);
                            } catch (Exception titleException) {
                                logger.error("设置图表标题时出错: " + titleException.getMessage());
                            }
                        }

                        // 重新绘图，使更改生效
                        chart.plot(xddfChartData);
                        logger.info("重新绘制图表");

                    } catch (Exception e) {
                        logger.error("填充饼图数据时出错: " + e.getMessage());
                        e.printStackTrace();
                    }
                }
            }
        }
    }
}