package com.xinge.replace;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import org.apache.commons.lang.StringUtils;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

public class MakeTemplateWord
{

    private WordprocessingMLPackage template;
    private String target;


    /**
     *
     */
    public MakeTemplateWord(String source, String target) throws Docx4JException, FileNotFoundException
    {
        this.template = getTemplate(source);
        this.target = target;
    }


    /**
     * @param name
     * @return
     * @throws Docx4JException
     * @throws FileNotFoundException
     */
    private WordprocessingMLPackage getTemplate(String name) throws Docx4JException, FileNotFoundException
    {
        WordprocessingMLPackage template = WordprocessingMLPackage.load(new FileInputStream(new File(name)));
        return template;
    }


    /**
     * @param obj
     * @param toSearch
     * @return
     */
    private static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch)
    {
        List<Object> result = new ArrayList<Object>();
        if (obj instanceof JAXBElement)
        {
            obj = ((JAXBElement<?>) obj).getValue();
        }

        if (obj.getClass().equals(toSearch))
        {
            result.add(obj);
        }
        else if (obj instanceof ContentAccessor)
        {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children)
            {
                result.addAll(getAllElementFromObject(child, toSearch));
            }

        }
        return result;
    }

    /**
     * @param name
     * @param placeholder
     */
    private void replacePlaceholder(String name, String placeholder)
    {
        List<Object> texts = getAllElementFromObject(template.getMainDocumentPart(), Text.class);

        for (Object text : texts)
        {
            Text textElement = (Text) text;
            if (textElement.getValue().equals(placeholder))
            {
                textElement.setValue(name);
            }
        }
    }

    private void replaceParagraph(String placeholder, String textToAdd)
    {
        // 1. get the paragraph
        List<Object> paragraphs = getAllElementFromObject(template.getMainDocumentPart(), P.class);

        P toReplace = null;
        for (Object p : paragraphs)
        {
            List<Object> texts = getAllElementFromObject(p, Text.class);
            for (Object t : texts)
            {
                Text content = (Text) t;
                if (content.getValue().equals(placeholder))
                {
                    toReplace = (P) p;
                    break;
                }
            }
        }

        // we now have the paragraph that contains our placeholder: toReplace
        // 2. split into seperate lines
        String as[] = StringUtils.splitPreserveAllTokens(textToAdd, '\n');

        for (int i = 0; i < as.length; i++)
        {
            String ptext = as[i];

            // 3. copy the found paragraph to keep styling correct
            P copy = (P) XmlUtils.deepCopy(toReplace);

            // replace the text elements from the copy
            List<?> texts = getAllElementFromObject(copy, Text.class);
            if (texts.size() > 0)
            {
                Text textToReplace = (Text) texts.get(0);
                textToReplace.setValue(ptext);
            }

            // add the paragraph to the document
            //			 template.getMainDocumentPart().getContent().add(copy);
            ((ContentAccessor) toReplace.getParent()).getContent().add(copy);
        }

        // 4. remove the original one
        ((ContentAccessor) toReplace.getParent()).getContent().remove(toReplace);

    }


    /**
     * @param tables
     * @param templateKey
     * @return
     * @throws Docx4JException
     * @throws JAXBException
     */
    private Tbl getTemplateTable(List<Object> tables, String templateKey) throws Docx4JException, JAXBException
    {
        for (Iterator<Object> iterator = tables.iterator(); iterator.hasNext(); )
        {
            Object tbl = iterator.next();
            List<?> textElements = getAllElementFromObject(tbl, Text.class);
            for (Object text : textElements)
            {
                Text textElement = (Text) text;
                if (textElement.getValue() != null && textElement.getValue().equals(templateKey))
                {
                    return (Tbl) tbl;
                }
            }
        }
        return null;
    }

    /**
     * @param placeholders
     * @param textToAdd
     * @throws Docx4JException
     * @throws JAXBException
     */
    private void replaceTable(String placeholders, List<Map<String, String>> textToAdd) throws Docx4JException, JAXBException
    {
        List<Object> tables = getAllElementFromObject(template.getMainDocumentPart(), Tbl.class);

        // 1. find the table
        Tbl tempTable = getTemplateTable(tables, placeholders);
        List<Object> rows = getAllElementFromObject(tempTable, Tr.class);

        // first row is header, second row is content
        if (rows.size() == 2)
        {
            // this is our template row
            Tr templateRow = (Tr) rows.get(1);

            for (Map<String, String> replacements : textToAdd)
            {
                // 2 and 3 are done in this method
                addRowToTable(tempTable, templateRow, replacements);
            }

            // 4. remove the template row
            tempTable.getContent().remove(templateRow);
        }
    }

    /**
     * @param reviewtable
     * @param templateRow
     * @param replacements
     */
    private static void addRowToTable(Tbl reviewtable, Tr templateRow, Map<String, String> replacements)
    {
        // 拷贝一行
        Tr workingRow = (Tr) XmlUtils.deepCopy(templateRow);
        //在该行中找到所有列
        List<?> textElements = getAllElementFromObject(workingRow, Tc.class);
        for (Object object : textElements)
        {
            Tc tc = (Tc) object;
            //列中的内容是否在需要替换的内容中
            String replacementValue = (String) replacements.get(tc.getContent().get(0).toString());
            if (replacementValue != null)
            {
                //找到该列中所有文字，用改方式是解决 列中的文字有可能被解析成多个字符串（多个Text对像）
                List<Object> textEs = getAllElementFromObject(tc, Text.class);
                if (textEs.size() == 0)
                {
                    return;
                }
                for (Object tempTex : textEs)
                {
                    //把该列的内容设置为空
                    Text temT = (Text) tempTex;
                    temT.setValue("");
                }
                //随便找一个字符对象填写新的内容
                Text text = (Text) textEs.get(0);
                text.setValue(replacementValue);
            }
        }

        reviewtable.getContent().add(workingRow);
    }


    /**
     * @throws IOException
     * @throws Docx4JException
     */
    public void writeDocxToStream() throws IOException, Docx4JException
    {
        File f = new File(this.target);
        template.save(f);
    }

    /**
     * @param replacements
     */
    public void test(Map<String, String> replacements, Map<String, Object> tblReplacements, Map<String, String> pgReplacements) throws Exception
    {
        processTexts(replacements);
        processTables(tblReplacements);
         processPgTexts(pgReplacements);
        writeDocxToStream();
    }

    public void processTexts(Map<String, String> replacements)
    {
        //进行文本占位符替换
        for (Map.Entry<String, String> entry : replacements.entrySet())
        {
            replacePlaceholder(entry.getValue(), entry.getKey());
        }
    }

    public void processTables(Map<String, Object> tblReplacements) throws Docx4JException, JAXBException
    {
        //表格替换
        for (Map.Entry<String, Object> entry : tblReplacements.entrySet())
        {
            String placeholder = entry.getKey();
            List<Map<String, String>> textToAdd = (List<Map<String, String>>) entry.getValue();
            replaceTable(placeholder, textToAdd);
        }
    }

    public void processPgTexts(Map<String, String> pgReplacements)
    {
        //文本段落替换
        for (Map.Entry<String, String> entry : pgReplacements.entrySet())
        {
            String placeholder = entry.getKey();
            String textToAdd = (String) entry.getValue();
            replaceParagraph(placeholder, textToAdd);
        }
    }


    /**
     * @param args
     */
    public static void main(String[] args) throws Exception
    {
        // TODO Auto-generated method stub
        String source = "c:\\tmp\\docs\\template.docx";
        String target = "c:\\tmp\\docs\\ouput.docx";
        MakeTemplateWord make = new MakeTemplateWord(source, target);

        Map<String, String> repl1 = new HashMap<String, String>();
        repl1.put("SJ_FUNCTION", "function1");
        repl1.put("SJ_DESC", "desc1");
        repl1.put("SJ_PERIOD", "period1");

        Map<String, String> repl2 = new HashMap<String, String>();
        repl2.put("SJ_FUNCTION", "function2");
        repl2.put("SJ_DESC", "desc2");
        repl2.put("SJ_PERIOD", "period2");

        Map<String, String> repl3 = new HashMap<String, String>();
        repl3.put("SJ_FUNCTION", "function3");
        repl3.put("SJ_DESC", "desc3");
        repl3.put("SJ_PERIOD", "period3");


        Map<String, String> replacements = new HashMap<String, String>();
        replacements = getYWData();
        //replacements.put("", "");
        //


        List<Map> tabList = new ArrayList<Map>();
        Map<String, Object> tblReplacements = new HashMap<String, Object>();
        Map<String, String> copy = new HashMap<String, String>();
        copy.put("FIELD0", "复奉直流");
        copy.put("FIELD1", "0");
        copy.put("FIELD2", "0");
        copy.put("FIELD3", "双极停运");
        tabList.add(copy);
        copy = new HashMap<String, String>();
        copy.put("FIELD0", "锦苏直流");
        copy.put("FIELD1", "4367");
        copy.put("FIELD2", "2862");
        copy.put("FIELD3", "双极四阀组大地回线");
        tabList.add(copy);
        copy = new HashMap<String, String>();
        copy.put("FIELD0", "天中直流");
        copy.put("FIELD1", "test1");
        copy.put("FIELD2", "test2");
        copy.put("FIELD3", "双极四阀组大地回线");
        tabList.add(copy);
        tblReplacements.put("FIELD0", tabList);

        Map<String, String> pgReplacements = new HashMap<String, String>();
        pgReplacements.put("LAST_WEEK", "1.继续跟踪复奉直流年度检修消缺，协调复龙站4.1和4.2阀侧套管（换下的400kV换流变）运往沈阳传奇修复。\n2.继续跟踪宾金直流调试工作。按照国网运检部要求，编制报送宾金直流生产准备费分解计划。\n5.按照国网法律部要求，组织开展国网公司第四批通用制度修改意见的征求工作。");
        pgReplacements.put("CURR_WEEK", "7.继续组织开展运维管理提升作业标准编制工作。\n10.参加哈郑工程GRTS避雷器技术讨论会、溪浙工程双龙换流站穿墙套管技术讨论会。组织ABB变压器组件公司技术交流会。");

        make.test(replacements, tblReplacements, pgReplacements);
    }


    public static Map getYWData()
    {
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("WEEK_NO", "2014年第10周");
        map.put("DEPT_NAME", "运维部门");
        map.put("REPORT_NO", "总第400期");
        map.put("DURATION", "(2014-04-30至2014-05-04)");
        Map tmp = null;
        map.put("DC_TRANS_CAPACITY", "截至2014年3月10日00:00，公司系统本年度直流输送电量为78.77亿千瓦时。");
        map.put("ASSET_FLAW", "3月6日，奉贤站发现极I低端Y/D-A相换流变10号冷却器渗油，经厂家、技术监督单位确认，3月9日对该冷却器进行了更换");
        map.put("SYSTEM_FAULT", "3月6日16:13，西锦III线B相故障，重合不成功跳闸，线路故障测距距离锦屏站57.6km，站内一、二次设备检查无异常。16:55按国调令将西锦III线转运行，16:58线路再次跳闸，18:56西锦III线转运行正常");

        List<Map> lastWeekList = new ArrayList<Map>();
        tmp = new HashMap();
        tmp.put("content", "1.继续跟踪复奉直流年度检修消缺，协调复龙站4.1和4.2阀侧套管（换下的400kV换流变）运往沈阳传奇修复。梳理天中直流消缺后的工程遗留问题，目前还有中州站极2高端Y/Y-B换流变铁芯夹件间绝缘异常、高端换流变阀侧套管SF6密度继电器不满足非电量保护3取2配置要求、四阀组运行主泵的工频回路负荷开关换型未完成和天山站换流变充电时西门子阀控系统误报晶闸管无回报信号问题未彻底解决等31个问题未处理或在处理。");
        tmp.put("focus", "Y");
        lastWeekList.add(tmp);
        tmp = new HashMap();
        tmp.put("content", "2.继续跟踪宾金直流调试工作。按照国网运检部要求，编制报送宾金直流生产准备费分解计划。");
        tmp.put("focus", "Y");
        lastWeekList.add(tmp);
        tmp = new HashMap();
        tmp.put("content", "3.天山、中州站完成非电量保护定值清理并完成标准化定值单的编制、审批、执行和备案，组织各站全面推广此项工作，加强公司所辖各站非电量保护定值和定值单管理。以保护定值为突破口，组织开展站用电、阀水冷、非电量保护装置全面排查工作，避免“7.5”类似事件发生");
        tmp.put("focus", "Y");
        lastWeekList.add(tmp);
        tmp = new HashMap();
        tmp.put("content", "4.组织向公司生产管理信息系统关键用户通报PMS1.0应用情况，PMS2.0建设进展和下一步工作计划。");
        tmp.put("focus", "Y");
        lastWeekList.add(tmp);
        tmp = new HashMap();
        tmp.put("content", "5.按照国网法律部要求，组织开展国网公司第四批通用制度修改意见的征求工作。");
        tmp.put("focus", "Y");
        lastWeekList.add(tmp);
        tmp = new HashMap();
        tmp.put("content", "6.编制新工程生产准备工作月报和特高压换流站运行月报并在公司OA系统电子公告栏发布。");
        tmp.put("focus", "N");
        lastWeekList.add(tmp);
        tmp = new HashMap();
        tmp.put("content", "7.继续组织开展运维管理提升作业标准编制工作。");
        tmp.put("focus", "N");
        lastWeekList.add(tmp);
        tmp = new HashMap();
        tmp.put("content", "8.编制完成公司技改大修项目立项原则。协调国网运检部和浙江、江苏公司做好2014年项目计划下达。继续组织完善一体化在线监测平台，梳理讨论在线监测设备类型和数据告警规则");
        tmp.put("focus", "N");
        lastWeekList.add(tmp);
        tmp = new HashMap();
        tmp.put("content", "9.组织开展2014年度国网科技进步奖、专利奖推荐工作，组织征集电网设备状态检测技术应用案例。组织开展公司2014年运维工作会和第一次生产准备例会筹备工作。");
        tmp.put("focus", "N");
        lastWeekList.add(tmp);
        tmp = new HashMap();
        tmp.put("content", "10.参加哈郑工程GRTS避雷器技术讨论会、溪浙工程双龙换流站穿墙套管技术讨论会。组织ABB变压器组件公司技术交流会。");
        tmp.put("focus", "N");
        lastWeekList.add(tmp);


        //		map.put("LAST_WEEK", lastWeekList);


        List<Map> currWeekList = new ArrayList<Map>();
        tmp = new HashMap();
        tmp.put("content", "1.继续跟踪2014年复龙、奉贤换流站年度检修，做好现场安全、质量、进度控制，做好检修总结的编报工作。");
        tmp.put("focus", "Y");
        currWeekList.add(tmp);
        tmp = new HashMap();
        tmp.put("content", "   2.参加宾金直流双极低端系统调试启动会议，继续跟踪宾金直流双极低端系统调试。");
        tmp.put("focus", "Y");
        currWeekList.add(tmp);
        tmp = new HashMap();
        tmp.put("content", "   3.现场检查大修期间备品备件使用、管理情况。");
        tmp.put("focus", "Y");
        currWeekList.add(tmp);
        tmp = new HashMap();
        tmp.put("content", "   4.继续组织开展公司2014第二批物资非招标采购前期准备工作。");
        tmp.put("focus", "Y");
        currWeekList.add(tmp);
        tmp = new HashMap();
        tmp.put("content", "   5.继续开展公司零星物资采购工作。");
        tmp.put("focus", "Y");
        currWeekList.add(tmp);
        tmp = new HashMap();
        tmp.put("content", "   6.筹备与许继集团技术交流会议。");
        tmp.put("focus", "N");
        currWeekList.add(tmp);
        tmp = new HashMap();
        tmp.put("content", "   7.继续开展公司招标资料档案整理工作。");
        tmp.put("focus", "N");
        currWeekList.add(tmp);
        tmp = new HashMap();
        tmp.put("content", "   8.开展前期调研，申报直流自动功率曲线控制技术咨询项目和特高压直流停电检修工期研究技术咨询项目。");
        tmp.put("focus", "N");
        currWeekList.add(tmp);
        //		map.put("CURR_WEEK", currWeekList);
        return map;
    }


}
