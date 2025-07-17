# PoiPptGen

一个基于 Apache POI 的 PPT 模板生成工具，支持通过模板自动生成甘特图、折线图、饼图等多种图表，方便快捷地制作带有动态数据的PPT。

## 功能介绍

  * 支持基于 PPT 模板生成包含多种图表的演示文稿。
  * 支持**甘特图 (Gantt Chart)**、**折线图 (Line Chart)**、**饼图 (Pie Chart)**。
  * 支持自定义数据填充，图表标题和数据动态更新。
  * 内置日志记录，方便调试和排查问题。

## 技术栈

  * **Java 8+**
  * **Apache POI** (POI-OOXML 和 XDDF)
  * **SLF4J** 日志框架
  * **Spring Core** (资源文件加载)

## 运行示例

参考test目录下的示例

## 模板文件说明

  * `gantt_template.pptx`：甘特图 PPT 模板
  * `line_template.pptx`：折线图 PPT 模板
  * `pie_template.pptx`：饼图 PPT 模板

请在模板中预先插入对应的图表占位符，程序会自动根据数据填充图表。

## 日志说明

项目使用 SLF4J 记录日志，方便调试及错误排查，默认会打印生成进度及错误信息。

**作者**：zzw
**版本**：1.0
**日期**：2025年
