import com.mygs.trackppt.pojo.PieChartData;
import com.mygs.trackppt.utils.GanttChartPptUtil;
import com.mygs.trackppt.utils.LineChartPptUtil;
import com.mygs.trackppt.utils.PieChartPptUtil;
import org.junit.jupiter.api.Test;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 随机图表数据测试类，用于生成折线图、饼图、甘特图的PPT演示文件
 */
public class RanDataTest {

    // 时间戳后缀（用于唯一命名输出文件）
    private static final String TIME_STAMP = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));

    // 输出文件名
    private static final String LINE_FILE_NAME = String.format("line-chart-example_%s.pptx", TIME_STAMP);
    private static final String PIE_FILE_NAME = String.format("pie-chart-example_%s.pptx", TIME_STAMP);
    private static final String GANTT_FILE_NAME = String.format("gantt-chart-example_%s.pptx", TIME_STAMP);

    // 输出路径
    private static final String LINE_OUTPUT_PATH = "src/main/resources/output/" + LINE_FILE_NAME;
    private static final String PIE_OUTPUT_PATH = "src/main/resources/output/" + PIE_FILE_NAME;
    private static final String GANTT_OUTPUT_PATH = "src/main/resources/output/" + GANTT_FILE_NAME;

    // 模板文件路径
    private static final String LINE_TEMPLATE_FILE_PATH = "src/main/resources/templates/line_template.pptx";
    private static final String PIE_TEMPLATE_FILE_PATH = "src/main/resources/templates/pie_template.pptx";
    private static final String GANTT_TEMPLATE_FILE_PATH = "src/main/resources/templates/gantt_template.pptx";

    /**
     * 测试：生成甘特图PPT
     */
    @Test
    public void testGanttChartGeneration() {
        GanttChartPptUtil.generatePPTChart(GANTT_TEMPLATE_FILE_PATH, GANTT_OUTPUT_PATH);
    }

    /**
     * 测试：生成折线图PPT
     */
    @Test
    public void testLineChartGeneration() {
        // 生成二维随机数据
        double[][] data = LineChartPptUtil.generateRandomLineData();

        // 转换为 List<List<Double>> 格式，便于传入图表生成方法
        List<List<Double>> dataList = Arrays.stream(data)
                .map(row -> Arrays.stream(row).boxed().collect(Collectors.toList()))
                .collect(Collectors.toList());

        // 生成折线图PPT
        LineChartPptUtil.generatePPTChart(LINE_TEMPLATE_FILE_PATH, LINE_OUTPUT_PATH, 1, "示例折线图标题", dataList);
    }

    /**
     * 测试：生成饼图PPT
     */
    @Test
    public void testPieChartGeneration() {
        // 生成模拟饼图数据
        Map<String, Double> pieData = PieChartPptUtil.generateRandomPieData();

        // 对数据按值降序排序
        List<Map.Entry<String, Double>> sortedEntries = new ArrayList<>(pieData.entrySet());
        sortedEntries.sort((e1, e2) -> Double.compare(e2.getValue(), e1.getValue()));

        // 构建有序Map用于生成图表
        Map<String, Double> sortedPieData = new LinkedHashMap<>();
        for (Map.Entry<String, Double> entry : sortedEntries) {
            sortedPieData.put(entry.getKey(), entry.getValue());
        }

        // 封装饼图数据对象
        PieChartData pieChartData = new PieChartData("示例饼图标题", sortedPieData);

        // 生成饼图PPT
        PieChartPptUtil.generatePieChartPPT(PIE_TEMPLATE_FILE_PATH, PIE_OUTPUT_PATH, pieChartData, 1);
    }
}
