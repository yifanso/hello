package com.onlyoffice.integration.controllers;

import cn.hutool.core.io.resource.ResourceUtil;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.*;
import com.deepoove.poi.data.style.ParagraphStyle;
import com.deepoove.poi.data.style.RowStyle;
import com.deepoove.poi.data.style.Style;
import com.onlyoffice.integration.dto.SheetMigration;
import com.onlyoffice.integration.excel.Grid;
import com.onlyoffice.integration.excel.UserCell;
import com.onlyoffice.integration.utils.RestResponse;
import com.onlyoffice.integration.utils.SheetToPicture;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.list.TreeList;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.Color;
import java.awt.image.BufferedImage;
import java.io.*;
import java.text.DecimalFormat;
import java.util.*;
import java.util.List;

import static com.onlyoffice.integration.convert.DrawFromExcel.isInMerged;
import static java.lang.Thread.sleep;
import static org.apache.poi.ss.usermodel.CellType.STRING;


@RestController
@CrossOrigin("*")
@Slf4j
public class DrawTable {

    @Autowired
    private SheetToPicture sheetToPicture;

    @Value("${file.save.path}")
    private String path;

    @PostMapping("/cross/table3")
    public RestResponse sheet2Word(@RequestBody SheetMigration migration) throws Exception {

        RestResponse response = new RestResponse<>();
        // 获取 Word 模板所在路径
        String docPath = "/Users/songyifan/Desktop/Java.Spring.Example/Java Spring Example/documents/127.0.0.1/";
        String filepath = docPath + migration.getDocxFileName();
        // 通过 XWPFTemplate 编译文件并渲染数据到模板中
        XWPFTemplate template = XWPFTemplate.compile(filepath).render(
                new HashMap<String, Object>() {
                    {
                        FileInputStream inputStream = new FileInputStream(new File("/Users/songyifan/Desktop/Java.Spring.Example/Java Spring Example/documents/127.0.0.1/" + migration.getXlsFileName()));
                        Workbook workbook = new XSSFWorkbook(inputStream);  // 读取XLSX格式的文件
                        Sheet sheet = workbook.getSheetAt(0);
                        List<Float> withList = new ArrayList();

                        //RowRenderData tableHead = Rows.of(cell).center().bgColor("3672e5").create();
                        log.info("width0: {}", sheet.getColumnWidthInPixels(0));
                        // 表格数据初始化
                        List<RowRenderData> renderDataList = new TreeList<>();
                        for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                            CellRenderData[] rowData = new CellRenderData[sheet.getRow(0).getPhysicalNumberOfCells()];
                            float high = 0;
                            withList.clear();
                            for (int j = 0; j < sheet.getRow(0).getPhysicalNumberOfCells(); j++) {
                                //if (sheet.getRow(i) != null && sheet.getRow(i).getCell(j) != null) {
                                    // 设置文本

                                    withList.add(j, sheet.getColumnWidthInPixels(j));
                                    String strCell;
                                    Cell cell = sheet.getRow(i).getCell(j);
                                    if(sheet.getRow(i+1) != null) {
                                        high = sheet.getRow(i + 1).getHeightInPoints() - sheet.getRow(i).getHeightInPoints();
                                    }
                                cell.setCellType(STRING);
                                CellStyle cs = cell.getCellStyle();
                                    CellType cellType = cell.getCellType();
                                    switch (cellType) {
                                        case STRING:
                                            strCell = cell.getStringCellValue();
                                            break;
                                        case NUMERIC:
                                            //判断单元格内容  是否是 数值较大的数字 （即是否用了E表示）
                                            String str = String.valueOf(cell.getNumericCellValue());
                                            if (str.contains("E")) {
                                                String LeftEStr = str.toString().split("E")[0];
                                                strCell = LeftEStr.split("\\.")[0] + LeftEStr.split("\\.")[1];
                                            } else {
                                                strCell = String.valueOf(cell.getNumericCellValue());
                                            }
                                            break;
                                        case BLANK:
                                            strCell = "";
                                            break;
                                        case FORMULA:
                                            try {
                                                strCell = String.valueOf(cell.getNumericCellValue());
                                            } catch (IllegalStateException e) {
                                                strCell = String.valueOf(cell.getRichStringCellValue());
                                            }
                                            break;
                                        default:
                                            strCell = "";
                                            break;
                                    }
                                    //String value = sheet.getRow(i).getCell(j).getStringCellValue();
                                    XSSFColor bgColor = (XSSFColor) sheet.getRow(i).getCell(j).getCellStyle().getFillForegroundColorColor();
                                    if (bgColor != null) {
                                        log.info("color: {}", bgColor.getARGBHex());
                                    }
                                    //String color = bgColor.getARGBHex();
                                    int height = sheet.getRow(i).getHeight();
                                    if (strCell != null) {
                                        CellRenderData cell1 = new CellRenderData();
                                        ParagraphRenderData graph = new ParagraphRenderData();
                                        graph.addText(strCell);
                                        cell1.addParagraph(graph);
                                        com.deepoove.poi.data.style.CellStyle style = new com.deepoove.poi.data.style.CellStyle();
                                        if (bgColor != null) {
                                            style.setBackgroundColor(bgColor.getARGBHex());
                                        }
                                        Font font1 = workbook.getFontAt(cs.getFontIndex());
                                        Style textStyle;
                                        if(font1.getBold()) {
                                             textStyle = Style.builder()
                                                    .buildFontSize(font1.getFontHeightInPoints())
                                                    .buildFontFamily(font1.getFontName())
                                                    .buildBold()
                                                    .build();
                                        } else {
                                             textStyle = Style.builder()
                                                    .buildFontSize(font1.getFontHeightInPoints())
                                                    .buildFontFamily(font1.getFontName())
                                                    .build();
                                        }
                                        ParagraphStyle paragraphStyle = ParagraphStyle.builder()
                                                .withDefaultTextStyle(textStyle)
                                                .build();
                                        style.setDefaultParagraphStyle(paragraphStyle);
                                        cell1.setCellStyle(style);
                                        rowData[j] = cell1;
                                    }
                                }
                            //}
                            RowRenderData row = Rows.of(rowData).rowHeight(high * 2.54/72).center().create();
                            renderDataList.add(row);
                        }
                        // 表格行构建
                        RowRenderData[] tableRows = new RowRenderData[sheet.getPhysicalNumberOfRows()];
                        CellRenderData cellRenderData = new CellRenderData();
                        // 添加数据行
                        for (int i = 0; i < renderDataList.size(); i++) {
                            tableRows[i] = renderDataList.get(i);
                        }

                        Map<MergeCellRule.Grid, MergeCellRule.Grid> map = new HashMap<>();
                        MergeCellRule.MergeCellRuleBuilder builder1 = MergeCellRule.builder();

                        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
                            System.out.println("Merged region start row: " + mergedRegion.getFirstRow());
                            System.out.println("Merged region end row: " + mergedRegion.getLastRow());
                            System.out.println("Merged region start column: " + mergedRegion.getFirstColumn());
                            System.out.println("Merged region end column: " + mergedRegion.getLastColumn());
                            builder1.map(MergeCellRule.Grid.of(mergedRegion.getFirstRow(), mergedRegion.getFirstColumn()), MergeCellRule.Grid.of(mergedRegion.getLastRow(), mergedRegion.getLastColumn()));

                            //map.put(MergeCellRule.Grid.of(mergedRegion.getFirstRow(), mergedRegion.getFirstColumn()), MergeCellRule.Grid.of(mergedRegion.getFirstColumn(), mergedRegion.getLastColumn()));

                        }
                        MergeCellRule rule = builder1.build();
                        //MergeCellRule rule1 = MergeCellRule.builder().map(MergeCellRule.Grid.of(0, 0), MergeCellRule.Grid.of(3, 0)).map(MergeCellRule.Grid.of(1, 1), MergeCellRule.Grid.of(3, 0)).build();
                        com.deepoove.poi.data.style.BorderStyle borderStyle1 = new com.deepoove.poi.data.style.BorderStyle();

                        borderStyle1.setColor("A6A6A6");
                        borderStyle1.setSize(2);
                        borderStyle1.setType(XWPFTable.XWPFBorderType.SINGLE);
                        int width1 = withList.stream().mapToInt(Float::intValue).sum();
                        double[] doubleArray = new double[withList.size()];
                        for(int i = 0; i < withList.size(); i++) {
                            doubleArray[i] = withList.get(i) * 0.026 ;
                        }

                        log.info("width1 {}", width1);
                        put(migration.getLable(), Tables.of(tableRows).width(width1*0.026, doubleArray).mergeRule(rule).center().border(borderStyle1).create());


                    }
                }
        );
        try {
            // 将完成数据渲染的文档写出
            template.writeAndClose(new FileOutputStream(filepath));
        } catch (IOException e) {
            e.printStackTrace();
        }

        response.setCode(0);
        response.setMsg("success");
        return response;

    }

    private void pictureReplace() throws Exception {
        // 获取 Word 模板所在路径
        String filepath = "/Users/songyifan/Desktop/Java.Spring.Example/Java Spring Example/documents/172.20.10.11/11.docx";
        // 通过 XWPFTemplate 编译文件并渲染数据到模板中
        XWPFTemplate template = XWPFTemplate.compile(filepath).render(
                new HashMap<String, Object>() {
                    {
                        InputStream stream = ResourceUtil.getStream("/Users/songyifan/Desktop/Java.Spring.Example/Java Spring Example/documents/172.20.10.11/logo.jpg");
                        put("companyLogoUrl", Pictures.ofStream(stream, PictureType.PNG).create());
                    }
                }
        );
    }


    private void getMerge(String path) throws Exception {
        InputStream inputStream = new FileInputStream(path);
        // 创建workbook、sheet对象、存储内容集合
        Workbook workbook;
        Sheet sheet;
        List<List<String>> excelList;
        // 用于判断单元格背景颜色设置
        int bgColorFlag = 1;
        //if (path.split("\\.")[1].equals("xlsx")) {
        // 上传的excel是2010以及更高版本
        // 初始化workbook 对象
        workbook = new XSSFWorkbook(inputStream);
        Sheet sheet1 = workbook.getSheetAt(0);
        for (int i = 0; i < sheet1.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = sheet1.getMergedRegion(i);
            System.out.println("Merged region start row: " + mergedRegion.getFirstRow());
            System.out.println("Merged region end row: " + mergedRegion.getLastRow());
            System.out.println("Merged region start column: " + mergedRegion.getFirstColumn());
            System.out.println("Merged region end column: " + mergedRegion.getLastColumn());
        }
    }

    public void convert(String path, String name) {
        // 设置基本参数
        int sheetNum = 0;// 读取表格第几页 ->0才是第一页
        int imageWidth = 0;// 图片宽度
        int imageHeight = 0;// 图片高度
        // 创建字节输入流
        InputStream inputStream;
        try {
            inputStream = new FileInputStream(path);
            // 创建workbook、sheet对象、存储内容集合
            Workbook workbook;
            Sheet sheet;
            List<List<String>> excelList = new ArrayList<>();
            // 用于判断单元格背景颜色设置
            int bgColorFlag = 1;
            //if (path.split("\\.")[1].equals("xlsx")) {
            // 上传的excel是2010以及更高版本
            // 初始化workbook 对象
            workbook = new XSSFWorkbook(inputStream);
            // 读取xlsx文件内容
//            } else {
//                // 上传的excel是2007以及更低版本
//                // 初始化workbook 对象
//                workbook = new HSSFWorkbook(inputStream);
//                // 读取xlsx文件内容
//                excelList = readXls(path, sheetNum);
//                // excel版本过低，无法正确显示原单元格背景颜色，故flag设置为0，使用白色作为背景颜色
//                bgColorFlag = 0;
//            }
            // 初始化sheet对象
            sheet = workbook.getSheetAt(sheetNum);
            // 获取整个sheet中合并单元格组合的集合
            List<CellRangeAddress> rangeAddress = sheet.getMergedRegions();
            // 根据读取数据，动态获得表边界行列
            int totalRow = excelList.size() + 1;
            int totalCol = excelList.get(0).size();
            // 创建单元格数组，用于遍历单元格
            UserCell[][] cells = new UserCell[totalRow + 1][totalCol + 1];
            int[] rowPixPos = new int[totalRow + 1];// 存放行边界
            rowPixPos[0] = 0;
            int[] colPixPos = new int[totalCol + 1];// 存放列边界
            colPixPos[0] = 0;
            // 开始遍历单元格
            for (int i = 0; i < totalRow - 1; i++) {
                for (int j = 0; j < totalCol; j++) {
                    cells[i][j] = new UserCell();
                    cells[i][j].setCell(sheet.getRow(i).getCell(j));
                    cells[i][j].setRow(i);
                    cells[i][j].setCol(j);
                    boolean ifShow = !(sheet.isColumnHidden(j) || sheet.getRow(i)
                            .getZeroHeight());
                    cells[i][j].setShow(ifShow);
                    // 计算所求区域宽度
                    float widthPix = !ifShow ? 0
                            : (sheet.getColumnWidthInPixels(j)); // 如果该单元格是隐藏的，则置宽度为0
                    if (i == 0) {
                        imageWidth += widthPix;
                    }
                    colPixPos[j + 1] = (int) (widthPix * 1.15 + colPixPos[j]);
                }
                // 计算所求区域高度
                boolean ifShow = (i >= 0); // 行序列在指定区域中间
                ifShow = ifShow && !sheet.getRow(i).getZeroHeight(); // 行序列不能隐藏
                float heightPoint = !ifShow ? 0 : (sheet.getRow(i).getHeightInPoints()); // 如果该单元格是隐藏的，则置高度为0
                imageHeight += heightPoint;
                rowPixPos[i + 1] = (int) (heightPoint * 96 / 72) + rowPixPos[i];
            }
            imageHeight = imageHeight * 96 / 72;
            imageWidth = imageWidth * 115 / 100;
            //-------------- 设置单元格属性 ----------------
            List<Grid> grids = new ArrayList<Grid>();
            for (int i = 0; i < totalRow - 1; i++) {
                for (int j = 0; j < totalCol; j++) {
                    Grid grid = new Grid();
                    // 设置坐标和宽高
                    grid.setX(colPixPos[j]);
                    grid.setY(rowPixPos[i]);
                    grid.setWidth(colPixPos[j + 1] - colPixPos[j]);
                    grid.setHeight(rowPixPos[i + 1] - rowPixPos[i]);
                    grid.setRow(cells[i][j].getRow());
                    grid.setCol(cells[i][j].getCol());
                    grid.setShow(cells[i][j].isShow());
                    // 判断是否为合并单元格
                    int[] isInMergedStatus = isInMerged(grid.getRow(),
                            grid.getCol(), rangeAddress);
                    if (isInMergedStatus[0] == 0 && isInMergedStatus[1] == 0) {
                        // 此单元格是合并单元格，并且不是第一个单元格，需要跳过本次循环，不进行绘制
                        continue;
                    } else if (isInMergedStatus[0] != -1
                            && isInMergedStatus[1] != -1) {
                        // 此单元格是合并单元格，并且属于第一个单元格，则需要调整网格大小
                        int lastRowPos = isInMergedStatus[0] > totalRow - 1 ? totalRow - 1 : isInMergedStatus[0];
                        int lastColPos = isInMergedStatus[1] > totalCol - 1 ? totalCol - 1 : isInMergedStatus[1];
                        grid.setWidth(colPixPos[lastColPos + 1] - colPixPos[j]);
                        grid.setHeight(rowPixPos[lastRowPos + 1] - rowPixPos[i]);
                    }
                    // 单元格背景颜色
                    Cell cell = cells[i][j].getCell();
                    if (cell != null) {
                        CellStyle cs = cell.getCellStyle();
                        grid.setBgColor(cs.getFillForegroundColorColor());
                        // 设置字体
                        org.apache.poi.ss.usermodel.Font font = workbook.getFontAt(cs.getFontIndex());
                        grid.setFont(font);
                        // 设置前景色
                        grid.setFtColor(cs.getFillBackgroundColorColor());
                        // 设置文本
                        String strCell;
                        CellType cellType = cell.getCellType();
                        switch (cellType) {
                            case STRING:
                                strCell = cell.getStringCellValue();
                                break;
                            case NUMERIC:
                                //判断单元格内容  是否是 数值较大的数字 （即是否用了E表示）
                                String str = String.valueOf(cell.getNumericCellValue());
                                if (str.contains("E")) {
                                    String LeftEStr = str.toString().split("E")[0];
                                    strCell = LeftEStr.split("\\.")[0] + LeftEStr.split("\\.")[1];
                                } else {
                                    strCell = String.valueOf(cell.getNumericCellValue());
                                }
                                break;
                            case BLANK:
                                strCell = "";
                                break;
                            case FORMULA:
                                try {
                                    strCell = String.valueOf(cell.getNumericCellValue());
                                } catch (IllegalStateException e) {
                                    strCell = String.valueOf(cell.getRichStringCellValue());
                                }
                                break;
                            default:
                                strCell = "";
                                break;
                        }
                        if (cell.getCellStyle().getDataFormatString()
                                .contains("0.00%")) {
                            try {
                                double dbCell = Double.valueOf(strCell);
                                strCell = new DecimalFormat("0.00").format(dbCell * 100) + "%";
                            } catch (NumberFormatException e) {
                            }
                        }
                        grid.setText(strCell.matches("\\w*\\.0") ? strCell
                                .substring(0, strCell.length() - 2) : strCell);
                    }
                    grids.add(grid);
                }
            }

            BufferedImage image = new BufferedImage(imageWidth, imageHeight, BufferedImage.TYPE_INT_RGB);
            Graphics2D g2d = image.createGraphics();

            g2d.setColor(java.awt.Color.white);
            g2d.fillRect(0, 0, imageWidth, imageHeight);
            // 平滑字体
            g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            g2d.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
            g2d.setRenderingHint(RenderingHints.KEY_STROKE_CONTROL, RenderingHints.VALUE_STROKE_NORMALIZE);
            g2d.setRenderingHint(RenderingHints.KEY_TEXT_LCD_CONTRAST, 140);
            g2d.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);
            g2d.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
            // 绘制表格
            for (Grid g : grids) {
                if (!g.isShow()) {
                    continue;
                }
                // 绘制背景色
                if (bgColorFlag == 1) {
                    // Excel2010以及更高-->使用原单元格背景色
                    g2d.setColor(g.getBgColor() == null ? java.awt.Color.white : g.getBgColor());
                } else {
                    // Excel2007以及更低-->使用白色作为背景色
                    g2d.setColor(java.awt.Color.white);
                }
                g2d.fillRect(g.getX(), g.getY(), g.getWidth(), g.getHeight());
                // 绘制边框
                g2d.setColor(Color.black);
                g2d.setStroke(new BasicStroke(1));
                g2d.drawRect(g.getX(), g.getY(), g.getWidth(), g.getHeight());
                // 绘制文字,居中显示
                g2d.setColor(g.getFtColor());
                java.awt.Font font = g.getFont();
                if (font == null) {
                    continue;
                }
                FontMetrics fm = g2d.getFontMetrics(font);
                int strWidth = fm.stringWidth(g.getText());// 获取将要绘制的文字宽度
                g2d.setFont(font);
                g2d.drawString(
                        g.getText(),
                        g.getX() + (g.getWidth() - strWidth) / 2,
                        g.getY() + (g.getHeight() - font.getSize()) / 2 + font.getSize());
            }
            // 表格最后一行有可能不显示，手动画上一行
            g2d.drawLine(0, imageHeight - 1, imageWidth - 4, imageHeight - 1);
            g2d.dispose();
            ImageIO.write(image, "png", new File(name));
            workbook.close();
        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("----Output to PNG file Success!----");
    }

}
