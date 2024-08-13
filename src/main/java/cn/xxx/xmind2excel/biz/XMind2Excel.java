package cn.xxx.xmind2excel.biz;


import cn.xxx.xmind2excel.util.ExcelUtil;
import cn.xxx.xmind2excel.util.FileExtension;
import cn.xxx.xmind2excel.util.FileUtil;
import org.apache.poi.ss.usermodel.*;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * @author xiongchenghui
 * @date 2020-08-12
 * &Desc XMind to Excel tool
 */
public class XMind2Excel {
    private static Logger logger = LoggerFactory.getLogger(XMind2Excel.class);

    /** XMind 原始文件 */
    static File xMindFile;

    /** 输出的测试用例Excel文件 */
    static String excelFilePath;

    /** 用例集合 */
    static List<TestCasePO> testCases = new ArrayList<>();

    /** 用例集合 */
    public static TestCaseInfo testCaseInfo = new TestCaseInfo();

    /** 导出Excel Style */
    static CellStyle cellStyle = null;

    public static void setXMindFile(File xMindFile) {
        XMind2Excel.xMindFile = xMindFile;
    }

    public static void setExcelFilePath(String excelFilePath) {
        XMind2Excel.excelFilePath = excelFilePath;
    }

    /***
     * &Desc: xMind 转换 Excel文件
     * @param
     * @return void
     */
    public static void xMind2Excel() {
        // 将xMind转换为zip文件
        File xMindZipFile = FileUtil.transferXMind2Zip(xMindFile);
        // 获取xMind的zip文件路径
        String zipPath = xMindZipFile.getAbsolutePath();
        // 获取ZIP文件解压目录descDir \\转译符
        String descDir = zipPath.replaceAll("\\" + FileExtension.ZIP, "");

        // 解压ZIP文件 到descDir目录
        try {
            FileUtil.unZipFiles(zipPath, descDir);
        } catch (IOException e) {
            e.printStackTrace();
            logger.error(e.getMessage());
        }

        // 获取xMind content.xml文件
        File xmlFile = new File(descDir + File.separator + "content.xml");
        // 创建一个SAXReader对象
        SAXReader sax = new SAXReader();

        try{
            // 获取document对象,如果文档无节点，则会抛出Exception提前结束
            Document document = sax.read(xmlFile);
            // 获取xMind根节点
            Element root = document.getRootElement();
            // 从根节点开始遍历解析所有节点, 生成测试用例list
            testCases.clear();
            logger.info("******************解析用例********************");
            parseXMind(root);
            logger.info("=====================>用例解析完毕");
        }catch (DocumentException de){
            de.printStackTrace();
            logger.error(de.getMessage());
        }catch (FileNotFoundException nfe){
            nfe.printStackTrace();
            logger.error(nfe.getMessage());
        }

        //删除临时文件夹
        logger.info("******************清理临时文件********************");
        File temp = new File(xMindZipFile.getParent());
        FileUtil.deleteDir(temp);

        // list用例写入Excel 统计用例总数、步骤数、验证点数
        logger.info("******************用例写入Excel，统计用例基本信息********************");
        testCaseWrite2Excel();

        logger.info("******************用例转换完毕********************");

    }

    /***
     * &Desc: test case从list中写入Excel
     * @param
     * @return void
     */
    private static void testCaseWrite2Excel(){
        ExcelUtil.createExcel(excelFilePath);
        // 读取Excel文件;
        Workbook workbook = ExcelUtil.readExcel(excelFilePath);

        // 创建一个样式
        setCellStyle(workbook);
        // 获取解析用例的表格
        Sheet caseSheet = workbook.getSheetAt(0);
        // 创建表头
        setSheetColumnHeader(caseSheet, cellStyle);

        // 逐条写入用例 并统计测试步骤验证点数量
        int steps = 0;
        int checkPointers = 0;
        int testCaseRowNum = 1;
        Row testcaseRow = caseSheet.createRow(testCaseRowNum);
        for (TestCasePO po: testCases) {
            steps += po.getActions().size();
            checkPointers += po.getResults().size();
            insertTestCase2Excel(testcaseRow, po);
            if(testCaseRowNum < testCases.size()){
                testCaseRowNum ++;
                testcaseRow = caseSheet.createRow(testCaseRowNum);
            }
        }
        // 统计测试用例、测试步骤、验证点数量
        testCaseInfo.setTestCaseNo(testCases.size());
        testCaseInfo.setTestCaseSteps(steps);
        testCaseInfo.setTestCaseCheckPointers(checkPointers);

        // 关闭文件流
        OutputStream stream = null;
        try {
            stream = new FileOutputStream(excelFilePath);
            workbook.write(stream);
        } catch (IOException e) {
            e.printStackTrace();
            logger.error(e.getMessage());
        } finally {
            try {
                stream.close();
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

    /***
     * &Desc: 设置单元格格式
     * @param workbook 工作簿
     * @return void
     */
    private static void setCellStyle(Workbook workbook){
        /** 创建一个样式 */
        cellStyle = workbook.createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
    }

    /***
     * &Desc: 设置用例列头
     * @param sheet 工作sheet
     * @param cellStyle 单元格格式
     * @return void
     */
    private static void setSheetColumnHeader(Sheet sheet, CellStyle cellStyle){
        // 创建表头
        Row testcaseTitle = sheet.createRow(0);

        Cell testCatalogCell = testcaseTitle.createCell(TestCaseTemplate.TESTCASECATALOG);
        testCatalogCell.setCellValue("用例目录");
        testCatalogCell.setCellStyle(cellStyle);

        Cell testNameCell = testcaseTitle.createCell(TestCaseTemplate.TESTCASENAME);
        testNameCell.setCellValue("用例名称");
        testNameCell.setCellStyle(cellStyle);

        Cell predicationCell = testcaseTitle.createCell(TestCaseTemplate.PREDICATION);
        predicationCell.setCellValue("前置条件");
        predicationCell.setCellStyle(cellStyle);

        Cell stepCell = testcaseTitle.createCell(TestCaseTemplate.ACTIONS);
        stepCell.setCellValue("用例步骤");
        stepCell.setCellStyle(cellStyle);

        Cell expectCell = testcaseTitle.createCell(TestCaseTemplate.RESULTS);
        expectCell.setCellValue("预期结果");
        expectCell.setCellStyle(cellStyle);

        Cell testCaseTypeCell = testcaseTitle.createCell(TestCaseTemplate.TESTCASETYPE);
        testCaseTypeCell.setCellValue("用例类型");
        testCaseTypeCell.setCellStyle(cellStyle);

        Cell importanceCell = testcaseTitle.createCell(TestCaseTemplate.PRIORITY);
        importanceCell.setCellValue("用例等级");
        importanceCell.setCellStyle(cellStyle);

    }

    /***
     * &Desc: 测试用例对象 写入到Excel的指定行
     * @param testcaseRow 写入的Excel sheet 表行
     * @param testCasePO 测试用例对象
     * @return void
     */
    private static void insertTestCase2Excel(Row testcaseRow, TestCasePO testCasePO){
        Cell testCaseCatalogCell = testcaseRow.createCell(TestCaseTemplate.TESTCASECATALOG);
        testCaseCatalogCell.setCellStyle(cellStyle);
        testCaseCatalogCell.setCellValue(testCasePO.getTestCaseCatalog());

        Cell testCaseNameCell = testcaseRow.createCell(TestCaseTemplate.TESTCASENAME);
        testCaseNameCell.setCellStyle(cellStyle);
        testCaseNameCell.setCellValue(testCasePO.getTestCaseName());

        Cell predicationCell = testcaseRow.createCell(TestCaseTemplate.PREDICATION);
        predicationCell.setCellStyle(cellStyle);
        predicationCell.setCellValue(testCasePO.getPredication());

        Cell actionsCell = testcaseRow.createCell(TestCaseTemplate.ACTIONS);
        actionsCell.setCellStyle(cellStyle);
        List<String> actions = testCasePO.getActions();
        StringBuilder sb = new StringBuilder();
        for(int item=0; item<actions.size(); item++){
            sb.append(actions.get(item));
            if(item < actions.size()-1){
                sb.append("\n");
            }
        }
        actionsCell.setCellValue(sb.toString());

        Cell resultsCell = testcaseRow.createCell(TestCaseTemplate.RESULTS);
        resultsCell.setCellStyle(cellStyle);
        List<String> results = testCasePO.getResults();
        sb.setLength(0);
        for(int item=0; item<results.size(); item++){
            sb.append(results.get(item));
            if(item < results.size()-1){
                sb.append("\n");
            }
        }
        resultsCell.setCellValue(sb.toString());

        Cell testCaseTypeCell = testcaseRow.createCell(TestCaseTemplate.TESTCASETYPE);
        testCaseTypeCell.setCellStyle(cellStyle);
        testCaseTypeCell.setCellValue(testCasePO.getTestCaseType());

        Cell priorityCell = testcaseRow.createCell(TestCaseTemplate.PRIORITY);
        priorityCell.setCellStyle(cellStyle);
        priorityCell.setCellValue(testCasePO.getPriority());
    }

    /***
     * &Desc: 解析XMind
     * @param node 指定节点开始解析
     * @return void
     */
    private static void parseXMind(Element node) throws FileNotFoundException {
        // 解析节点
        if (node != null) {
            // 获取root topic信息
            dealXMindRootInfo(node);

            // 分支主题
            if("topics".equals(node.getName()) && "attached".equals(node.attributeValue("type"))){
                // 获取所有的topic分支节点
                List<Element> topicList = node.elements("topic");
                // 遍历所有的topic分支节点
                for (Element el: topicList) {
                    Element eMarker = el.element("marker-refs");
                    // 如果是模块 或 用例节点
                    if(eMarker != null){
                        Element eMarkerRef = eMarker.element("marker-ref");
                        String markerIcon = eMarkerRef.attributeValue("marker-id");
                        // 用例
                        if(markerIcon.startsWith("priority")){
                            logger.info("###########当前node为用例:" + el.element("title").getStringValue());
                            TestCasePO testCasePO = dealXMindTestCase(el, markerIcon);
                            testCases.add(testCasePO);
                        }
                        // 模块
                        if(markerIcon.startsWith("flag") || markerIcon.startsWith("star")){
                            logger.info("###########当前node为模块:" + el.element("title").getStringValue());
                        }
                    }
                }
            }
        }

        // 递归遍历当前节点所有的子节点
        List<Element> listElement = node.elements();
        // 遍历所有一级子节点
        for (Element e : listElement) {
            // 递归
            parseXMind(e);
        }

    }

    /***
     * &Desc: 解析XMind: 根节点的信息
     * @param node 指定节点开始解析
     * @return void
     */
    private static void dealXMindRootInfo(Element node){
        // 获取root topic信息
        if("sheet".equals(node.getName())){
            logger.info("###########解析 root topic 信息###########");
            Element firstElement = node.element("topic");
            Element firstTopicNotes = firstElement.element("notes");
            if(firstTopicNotes !=  null){
                Element html = firstTopicNotes.element("html");
                List<Element> topicNotesProperty = html.elements();
                logger.info("###########开始获取root topic备注信息###########");
                for (Element e: topicNotesProperty) {
                    logger.info("=====================>" + e.getStringValue());
                }
                logger.info("###########开始获取root topic备注信息###########");
            }
        }

    }

    /***
     * &Desc: 解析XMind: 用例节点的相关信息，并组装成一条测试用例，放入 测试用例对象 中
     * @param testCaseTopicNode 测试用例节点
     * @param markerIcon 测试用例节点的Marker标识值
     * @return cn.xxx.xmind2excel.biz.TestCasePO
     */
    private static TestCasePO dealXMindTestCase(Element testCaseTopicNode, String markerIcon){
        TestCasePO testCasePO = new TestCasePO();

        /** 取得：用例目录 */
        StringBuilder testCaseCatalog = new StringBuilder();
        Element prevTopicNode = testCaseTopicNode.getParent().getParent().getParent();
        // 回溯路径找到Test Case目录
        while(!prevTopicNode.getParent().getName().equals("sheet")){
            String catalogItem = prevTopicNode.element("title").getStringValue();
            if(testCaseCatalog.toString().equals("")){
                testCaseCatalog.append(catalogItem);
            }else {
                testCaseCatalog.insert(0, "-")
                        .insert(0,catalogItem);
            }
            prevTopicNode = prevTopicNode.getParent().getParent().getParent();
        }
        // 记录用例模块路径备用
        String testCaseNameOfPrefix = testCaseCatalog.toString();
        // 补充根节点名称
        String rootNodeName = prevTopicNode.element("title").getStringValue();
        testCaseCatalog.insert(0, "-").insert(0,rootNodeName);

        /** 取得：用例名称 */
        String testCaseName = testCaseTopicNode.element("title").getStringValue();
        // 增加用例模块路径前缀
        if(testCaseName.length()>0){
            testCaseName = testCaseNameOfPrefix + ": " + testCaseName;
        }

        /** 取得：前置条件 */
        String predication = "";
        Element notes = testCaseTopicNode.element("notes");
        if(notes != null){
            predication = notes.element("plain").getStringValue();
        }

        /** 取得：优先级数字 */
        String suffix = markerIcon.replaceAll("priority-", "");
        String priority = "";
        // 判断用例优先级
        switch (suffix) {
            case "1":
                priority = "高";
                break;
            case "2":
                priority = "中";
                break;
            case "3":
                priority = "低";
                break;
            default:
                break;
        }

        /** 取得：步骤、期望 */
        List<String> steps = new ArrayList<>();
        List<String> results = new ArrayList<>();
        Element stepChildrenNode = testCaseTopicNode.element("children");
        // 获取用例的所有步骤
        try {
            if(stepChildrenNode != null){
                List<Element> stepTopicNodes = stepChildrenNode.element("topics").elements("topic");
                for (Element e: stepTopicNodes) {
                    // XMind copy节点无title节点问题，容错处理
                    Element step = e.element("title");
                    if(step != null){
                        steps.add(step.getStringValue());
                    }else {
                        steps.add("");
                    }
                    Element resultChildrenNode = e.element("children");
                    // 获取单个步骤的所有期望
                    if(resultChildrenNode != null){
                        List<Element> resultTopicNodes = resultChildrenNode.element("topics").elements("topic");
                        for (Element ee: resultTopicNodes) {
                            // XMind copy节点无title节点问题，容错处理
                            Element expect = ee.element("title");
                            if(expect != null){
                                results.add(expect.getStringValue());
                            }else {
                                results.add("");
                            }
                        }
                    }
                }
            }
        } catch (Exception e) {
            logger.error("******解析测试用例步骤和期望出错：" + testCaseName);
            e.printStackTrace();
        }

        /** 赋值用例对象 */
        testCasePO.setTestCaseCatalog(testCaseCatalog.toString());
        testCasePO.setTestCaseName(testCaseName);
        testCasePO.setPredication(predication);
        testCasePO.setActions(steps);
        testCasePO.setResults(results);
        testCasePO.setPriority(priority);
        testCasePO.setTestCaseType("功能测试");

        logger.info("=====================>解析测试用例： " + testCaseName);
        return testCasePO;
    }


}