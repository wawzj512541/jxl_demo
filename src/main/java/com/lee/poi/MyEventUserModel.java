package com.lee.poi;

import org.apache.commons.compress.utils.Lists;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.parsers.ParserConfigurationException;
import java.io.InputStream;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

/**
 * @author zhengqiang.zq
 * @date 2018/05/04 ,参考链接：https://poi.apache.org/spreadsheet/how-to.html#sxssf
 */
public class MyEventUserModel {
    public static ThreadLocal<List<ParsedRow>> local = new ThreadLocal<List<ParsedRow>>();

    public void processOneSheet(String filename) throws Exception {
        OPCPackage pkg = OPCPackage.open(filename);
        XSSFReader r = new XSSFReader(pkg);
        SharedStringsTable sst = r.getSharedStringsTable();

        XMLReader parser = fetchSheetParser(sst);
        InputStream sheet2 = r.getSheet("rId1");
        InputSource sheetSource = new InputSource(sheet2);
        parser.parse(sheetSource);
        sheet2.close();
    }

    public XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException, ParserConfigurationException {
        XMLReader parser = XMLHelper.newXMLReader();
        ContentHandler handler = new SheetHandler(sst);
        parser.setContentHandler(handler);
        return parser;
    }

    /**
     * See org.xml.sax.helpers.DefaultHandler javadocs
     */
    private class SheetHandler extends DefaultHandler {

        /**
         * excel 常量数据对象，对应的就是sharedStrings.xml文件中的内容，类似excel中的常量池
         */
        private SharedStringsTable sst;
        /**
         * 当前处理的文本值
         */

        private String lastContents;
        /**
         * 下一个文本是不是String类型
         */
        private boolean nextIsString;
        /**
         * 当前单元格的索引值，从0开始，0:第一列
         */
        private Short index;
        /**
         * 自定义数据类型，存储被解析的每一行原始数据
         */
        List<ParsedRow> sheetData = Lists.newArrayList();
        ParsedRow currentRow = new ParsedRow();

        private SheetHandler(SharedStringsTable sst) {
            this.sst = sst;
        }

        @Override
        public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
            //第一行
            if (name.equals("row")) {
                currentRow.setRowNum(new Long(attributes.getValue("r")));
                sheetData.add(currentRow);
            }
            //c => cell 一个单元格，
            if (name.equals("c")) {
                //r属性表示单元格位置，例如A2,C3
                String coordinate = attributes.getValue("r");
                CellReference cellReference = new CellReference(coordinate);
                //根据r属性获取其列下标，从0开始
                index = cellReference.getCol();

                //t:属性代表单元格类型
                String cellType = attributes.getValue("t");
                if (cellType != null && cellType.equals("s")) {
                    //t="s"表示是改单元格是字符串，那么该单元格的实际值值需要去SharedStringsTable中取
                    nextIsString = true;
                } else {
                    nextIsString = false;
                }
            }
            // Clear contents cache
            lastContents = "";
        }

        @Override
        public void endElement(String uri, String localName, String name) throws SAXException {
            if (nextIsString) {
                int idx = Integer.parseInt(lastContents);
                //从SharedStringsTable中取当前单元格的实际值
                lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
                nextIsString = false;
            }

            // v => contents of a cell
            // Output after we've seen the string contents
            if (name.equals("v")) {
                //不管是不是数字还是文本值
                currentRow.getData().put(index, lastContents);
            }
            if (name.equals("row")) {
                currentRow = new ParsedRow();
            }
        }

        @Override
        public void endDocument() throws SAXException {
            local.set(sheetData);
        }

        /**
         * 通知一个元素中的字符，是否处理由自己决定，比如  <v>1</v>，
         *
         * @param ch     The characters. 整个sheet.xml的char[]数组表示
         * @param start  The start position in the character array. 本次处理的元素值的的开始位置
         * @param length The number of characters to use from the ,元素长度
         *               character array.
         * @throws SAXException Any SAX exception, possibly
         *                      wrapping another exception.
         * @see ContentHandler#characters
         */
        @Override
        public void characters(char[] ch, int start, int length) throws SAXException {
            //对于lastContents是String类型来说，lastContent存放的是其在SharedStringsTable中的索引，
            // 对于是数字类型来说，lastContents存放就是该数字的字符串表示
            lastContents += new String(ch, start, length);
        }
    }

    public static void main(String[] args) throws Exception {
        String fileName = "C:\\Users\\lion\\Desktop\\project_info_10000.xlsx";
        MyEventUserModel example = new MyEventUserModel();

        Long startTime = new Date().getTime();
        example.processOneSheet(fileName);
        System.out.println("execution time:"+(new Date().getTime() - startTime));
        List<ParsedRow> parsedRows = local.get();
        for(int i = 1;i<parsedRows.size();i++){
            ParsedRow p = parsedRows.get(i);
            HashMap map = p.getData();
            ProjectInfo projectInfo = new ProjectInfo();
            projectInfo.convertsetProjectName(map.get((short)1).toString());
            projectInfo.convertsetUseAreaExcelLabel(map.get((short)2).toString());
            projectInfo.convertsetSkillAreaExcelLabel(map.get((short)3).toString());
            projectInfo.convertsetValuationExcelLabel(map.get((short)4).toString());
            projectInfo.convertsetFruitTypeExcelLabel(map.get((short)5).toString());
            projectInfo.convertsetYourWantExcelLabel(map.get((short)6).toString());
            projectInfo.convertsetYourWantContent(map.get((short)7).toString());
            projectInfo.convertsetProcessTypeExcelLabel(map.get((short)8).toString());
            projectInfo.convertsetFinancingTypeExcelLabel(map.get((short)9).toString());
            projectInfo.convertsetCooperateTypeExcelLabel(map.get((short)10).toString());
            projectInfo.convertsetCompanyName(map.get((short)11).toString());
            projectInfo.convertsetContactName(map.get((short)12).toString());
            projectInfo.convertsetPhoneNumber(map.get((short)13).toString());
            projectInfo.convertsetProvinceExcelLabel(map.get((short)14).toString());
            projectInfo.convertsetSummary(map.get((short)15).toString());
            projectInfo.convertsetDetail(map.get((short)16).toString());
            projectInfo.convertsetWordsExcelLabel(map.get((short)17).toString());
            System.out.println(projectInfo.toString());
        }
        System.out.println("execution time:"+(new Date().getTime() - startTime));
        Thread.sleep(1 * 1000);
    }
}