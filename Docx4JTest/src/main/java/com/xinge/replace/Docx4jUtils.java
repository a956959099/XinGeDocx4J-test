package com.xinge.replace;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.math.BigInteger;
import java.text.SimpleDateFormat;
import java.util.*;

import javax.imageio.ImageIO;
import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import org.docx4j.XmlUtils;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.model.table.TblFactory;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.DocumentSettingsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;
import org.docx4j.wml.SectPr.PgMar;
import org.docx4j.wml.TcPrInner.HMerge;
import org.docx4j.wml.TcPrInner.TcBorders;
import org.docx4j.wml.TcPrInner.VMerge;

/**
 * 功能描述：报表的工具类
 *
 * @author myclover
 */
@SuppressWarnings("all")
public class Docx4jUtils
{

    /**
     * 功能描述：创建文档处理包对象
     *
     * @return 返回值：返回文档处理包对象
     * @throws Exception
     * @author myclover
     */
    public static WordprocessingMLPackage createWordprocessingMLPackage() throws Exception
    {
        return WordprocessingMLPackage.createPackage();
    }

    /**
     * 功能描述：获取文档的可用宽度
     *
     * @param wordPackage 文档处理包对象
     * @return 返回值：返回值文档的可用宽度
     * @throws Exception
     * @author myclover
     */
    private static int getWritableWidth(WordprocessingMLPackage wordPackage) throws Exception
    {
        return wordPackage.getDocumentModel().getSections().get(0).getPageDimensions().getWritableWidthTwips();
    }

    /**
     * 功能描述：创建文档表格，上下双边框，左右不封口
     *
     * @param rows   行数
     * @param cols   列数
     * @param widths 每列的宽度
     * @return 返回值：返回表格对象
     * @author myclover
     */
    public static Tbl createTable(int rows, int cols, int[] widths)
    {
        ObjectFactory factory = Context.getWmlObjectFactory();
        Tbl tbl = factory.createTbl();
        // w:tblPr
        StringBuffer tblSb = new StringBuffer();
        tblSb.append("<w:tblPr ").append(Namespaces.W_NAMESPACE_DECLARATION).append(">");
        tblSb.append("<w:tblStyle w:val=\"TableGrid\"/>");
        tblSb.append("<w:tblW w:w=\"0\" w:type=\"auto\"/>");
        //上边框双线
        tblSb.append("<w:tblBorders><w:top w:val=\"double\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>");
        //左边无边框
        tblSb.append("<w:left w:val=\"none\" w:sz=\"0\" w:space=\"0\" w:color=\"auto\"/>");
        //下边框双线
        tblSb.append("<w:bottom w:val=\"double\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>");
        //右边无边框
        tblSb.append("<w:right w:val=\"none\" w:sz=\"0\" w:space=\"0\" w:color=\"auto\"/>");
        tblSb.append("</w:tblBorders>");
        tblSb.append("<w:tblLook w:val=\"04A0\"/>");
        tblSb.append("</w:tblPr>");
        TblPr tblPr = null;
        try
        {
            tblPr = (TblPr) XmlUtils.unmarshalString(tblSb.toString());
        } catch (JAXBException e)
        {
            e.printStackTrace();
        }
        tbl.setTblPr(tblPr);
        if (tblPr != null)
        {
            Jc jc = factory.createJc();
            //单元格居中对齐
            jc.setVal(JcEnumeration.CENTER);
            tblPr.setJc(jc);
            CTTblLayoutType tbll = factory.createCTTblLayoutType();
            // 固定列宽
            tbll.setType(STTblLayoutType.FIXED);
            tblPr.setTblLayout(tbll);
        }
        // <w:tblGrid><w:gridCol w:w="4788"/>
        TblGrid tblGrid = factory.createTblGrid();
        tbl.setTblGrid(tblGrid);
        // Add required <w:gridCol w:w="4788"/>
        for (int i = 1; i <= cols; i++)
        {
            TblGridCol gridCol = factory.createTblGridCol();
            gridCol.setW(BigInteger.valueOf(widths[i - 1]));
            tblGrid.getGridCol().add(gridCol);
        }
        // Now the rows
        for (int j = 1; j <= rows; j++)
        {
            Tr tr = factory.createTr();
            tbl.getContent().add(tr);
            // The cells
            for (int i = 1; i <= cols; i++)
            {
                Tc tc = factory.createTc();
                tr.getContent().add(tc);
                TcPr tcPr = factory.createTcPr();
                tc.setTcPr(tcPr);
                // <w:tcW w:w="4788" w:type="dxa"/>
                TblWidth cellWidth = factory.createTblWidth();
                tcPr.setTcW(cellWidth);
                cellWidth.setType("dxa");
                cellWidth.setW(BigInteger.valueOf(widths[i - 1]));
                tc.getContent().add(factory.createP());
            }

        }
        return tbl;
    }

    /**
     * 功能描述：创建文档表格，全部表格边框都为1
     *
     * @param rows   行数
     * @param cols   列数
     * @param widths 每列的宽度
     * @return 返回值：返回表格对象
     * @author myclover
     */
    public static Tbl createBorderTable(int rows, int cols, int[] widths)
    {
        ObjectFactory factory = Context.getWmlObjectFactory();
        Tbl tbl = factory.createTbl();
        // w:tblPr
        StringBuffer tblSb = new StringBuffer();
        tblSb.append("<w:tblPr ").append(Namespaces.W_NAMESPACE_DECLARATION).append(">");
        tblSb.append("<w:tblStyle w:val=\"TableGrid\"/>");
        tblSb.append("<w:tblW w:w=\"0\" w:type=\"auto\"/>");
        tblSb.append("<w:tblBorders><w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>");
        tblSb.append("<w:left w:val=\"single\" w:sz=\"0\" w:space=\"0\" w:color=\"auto\"/>");
        tblSb.append("<w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>");
        tblSb.append("<w:right w:val=\"single\" w:sz=\"0\" w:space=\"0\" w:color=\"auto\"/>");
        tblSb.append("</w:tblBorders>");
        tblSb.append("<w:tblLook w:val=\"04A0\"/>");
        tblSb.append("</w:tblPr>");
        TblPr tblPr = null;
        try
        {
            tblPr = (TblPr) XmlUtils.unmarshalString(tblSb.toString());
        } catch (JAXBException e)
        {
            e.printStackTrace();
        }
        tbl.setTblPr(tblPr);
        if (tblPr != null)
        {
            Jc jc = factory.createJc();
            //单元格居中对齐
            jc.setVal(JcEnumeration.CENTER);
            tblPr.setJc(jc);
            CTTblLayoutType tbll = factory.createCTTblLayoutType();
            // 固定列宽
            tbll.setType(STTblLayoutType.FIXED);
            tblPr.setTblLayout(tbll);
        }
        // <w:tblGrid><w:gridCol w:w="4788"/>
        TblGrid tblGrid = factory.createTblGrid();
        tbl.setTblGrid(tblGrid);
        // Add required <w:gridCol w:w="4788"/>
        for (int i = 1; i <= cols; i++)
        {
            TblGridCol gridCol = factory.createTblGridCol();
            gridCol.setW(BigInteger.valueOf(widths[i - 1]));
            tblGrid.getGridCol().add(gridCol);
        }
        // Now the rows
        for (int j = 1; j <= rows; j++)
        {
            Tr tr = factory.createTr();
            tbl.getContent().add(tr);
            // The cells
            for (int i = 1; i <= cols; i++)
            {
                Tc tc = factory.createTc();
                tr.getContent().add(tc);
                TcPr tcPr = factory.createTcPr();
                tc.setTcPr(tcPr);
                // <w:tcW w:w="4788" w:type="dxa"/>
                TblWidth cellWidth = factory.createTblWidth();
                tcPr.setTcW(cellWidth);
                cellWidth.setType("dxa");
                cellWidth.setW(BigInteger.valueOf(widths[i - 1]));
                tc.getContent().add(factory.createP());
            }

        }
        return tbl;
    }

    /**
     * 功能描述：创建文档表格，上下双边框，左右不封口
     *
     * @param wordPackage      文档处理包对象
     * @param rows             表格行数
     * @param cols             表格列数
     * @param tableWidth       表格的宽度
     * @param style            表格的样式
     * @param jcEnumerationVal 表格的对齐方式
     * @return 返回值：返回表格对象
     * @throws Exception
     * @author myclover
     */
    private static Tbl createTable(WordprocessingMLPackage wordPackage, int rows, int cols, String tableWidth, String style, JcEnumeration jcEnumerationVal) throws Exception
    {
        int writableWidthTwips = getWritableWidth(wordPackage);
        if (cols == 0)
        {
            cols = 1;
        }
        int cellWidth = new Double(Math.floor((writableWidthTwips / cols))).intValue();
        int[] widths = new int[cols];
        for (int i = 0; i < cols; i++)
        {
            widths[i] = cellWidth;
        }
        return createTable(rows, cols, widths);
    }

    /**
     * 功能描述：创建文档表格，表格的默认宽度为：表格样式：dxa，对齐方式：居中
     *
     * @param wordPackage 文档处理包对象
     * @param rows        表格行数
     * @param cols        表格列数
     * @param tableWidth  表格宽度
     * @return 返回值：返回表格对象
     * @throws Exception
     * @author myclover
     */
    public static Tbl createTable(WordprocessingMLPackage wordPackage, int rows, int cols, String tableWidth) throws Exception
    {
        return createTable(wordPackage, rows, cols, tableWidth, "dxa", JcEnumeration.CENTER);
    }

    /**
     * 功能描述：创建文档表格，表格的默认宽度为：9328，表格样式：dxa，对齐方式：居中
     *
     * @param wordPackage 文档处理包对象
     * @param rows        表格行数
     * @param cols        表格列数
     * @return 返回值：返回表格对象
     * @throws Exception
     * @author myclover
     */
    public static Tbl createTable(WordprocessingMLPackage wordPackage, int rows, int cols) throws Exception
    {
        return createTable(wordPackage, rows, cols, "9328", "dxa", JcEnumeration.CENTER);
    }

    /**
     * 功能描述：往文档对象中添加相应的内容
     *
     * @param wordPackage 文档处理包对象
     * @param info        需要添加的信息
     * @param unmarshal   是否有样式，表格对象默认不用
     * @throws Exception
     * @author myclover
     */
    private static void addObject(WordprocessingMLPackage wordPackage, Object info, boolean unmarshal) throws Exception
    {
        if (unmarshal)
        {
            wordPackage.getMainDocumentPart().addObject(org.docx4j.XmlUtils.unmarshalString(String.valueOf(info)));
        }
        else
        {
            wordPackage.getMainDocumentPart().addObject(info);
        }
    }

    /**
     * 功能描述：往文档中添加内容
     *
     * @param wordPackage 文档处理包对象
     * @param info        文档内容
     * @throws Exception
     * @author myclover
     */
    public static void addObject(WordprocessingMLPackage wordPackage, Object info) throws Exception
    {
        addObject(wordPackage, info, true);
    }

    /**
     * 功能描述：往文档中添加表格
     *
     * @param wordPackage 文档处理包对象
     * @param tbl         表格
     * @throws Exception
     * @author myclover
     */
    public static void addObjectForTbl(WordprocessingMLPackage wordPackage, Tbl tbl) throws Exception
    {
        addObject(wordPackage, tbl, false);
    }

    /**
     * 功能描述：保存文档信息
     *
     * @param wordPackage 文档处理包对象
     * @param fileName    完整的输出文件名称，包括路径
     * @throws Exception
     * @author myclover
     */
    public static void saveWordPackage(WordprocessingMLPackage wordPackage, String fileName) throws Exception
    {
        saveWordPackage(wordPackage, new File(fileName));
    }

    /**
     * 功能描述：保存文档信息
     *
     * @param wordPackage 文档处理包对象
     * @param file        文件
     * @throws Exception
     * @author myclover
     */
    public static void saveWordPackage(WordprocessingMLPackage wordPackage, File file) throws Exception
    {
        wordPackage.save(file);
    }

    /**
     * 功能描述：设置页边距
     *
     * @param wordPackage 文档处理包对象
     * @param top         上边距
     * @param bottom      下边距
     * @param left        左边距
     * @param right       右边距
     * @author myclover
     */
    public static void setMarginSpace(WordprocessingMLPackage wordPackage, String top, String bottom, String left, String right)
    {
        ObjectFactory factory = Context.getWmlObjectFactory();
        PgMar pg = factory.createSectPrPgMar();
        pg.setTop(new BigInteger(top));
        pg.setBottom(new BigInteger(bottom));
        pg.setLeft(new BigInteger(left));
        pg.setRight(new BigInteger(right));
        wordPackage.getDocumentModel().getSections().get(0).getSectPr().setPgMar(pg);
    }

    /**
     * 功能描述：设置页边距，上下边距都为1440，2.54厘米
     *
     * @param wordPackage 文档处理包对象
     * @param left        左边距
     * @param right       右边距
     * @author myclover
     */
    public static void setMarginSpace(WordprocessingMLPackage wordPackage, String left, String right)
    {
        setMarginSpace(wordPackage, "1440", "1440", left, right);
    }

    /**
     * 功能描述：设置页边距，上下边距都为1440,2.54厘米，左右边距都为1797,3.17厘米
     *
     * @param wordPackage 文档处理包对象
     * @author myclover
     */
    public static void setMarginSpace(WordprocessingMLPackage wordPackage)
    {
        setMarginSpace(wordPackage, "1440", "1440", "1797", "1797");
    }

    /**
     * 功能描述：设置字体的样式
     *
     * @param fontFamily     字体类型
     * @param colorVal       字体颜色
     * @param hpsMeasureSize 字号大小
     * @param sTHint         字体格式
     * @param isBlod         是否加粗
     * @return 返回值：返回字体样式对象
     * @throws Exception
     * @author myclover
     */
    private static RPr getRPr(String fontFamily, String colorVal, String hpsMeasureSize, STHint sTHint, boolean isBlod)
    {
        ObjectFactory factory = Context.getWmlObjectFactory();
        RPr rPr = factory.createRPr();

        org.docx4j.wml.RFonts rf = new org.docx4j.wml.RFonts();
        rf.setHint(sTHint);
        rf.setAscii(fontFamily);
        rf.setHAnsi(fontFamily);
        rPr.setRFonts(rf);

        BooleanDefaultTrue bdt = Context.getWmlObjectFactory().createBooleanDefaultTrue();
        rPr.setBCs(bdt);
        if (isBlod)
        {
            rPr.setB(bdt);
        }
        org.docx4j.wml.Color color = new org.docx4j.wml.Color();
        color.setVal(colorVal);
        rPr.setColor(color);

        org.docx4j.wml.HpsMeasure sz = new org.docx4j.wml.HpsMeasure();
        sz.setVal(new BigInteger(hpsMeasureSize));
        rPr.setSz(sz);
        rPr.setSzCs(sz);

        return rPr;
    }

    /**
     * 功能描述：设置字体的样式，宋体，黑色，18号
     *
     * @param isBlod 是否加粗
     * @return 返回值：返回字体样式对象
     * @throws Exception
     * @author myclover
     */
    private static RPr getRPr(boolean isBlod)
    {
        return getRPr("宋体", "000000", "18", STHint.EAST_ASIA, isBlod);
    }

    /**
     * 功能描述：设置字体的样式，黑色，18号
     *
     * @param fontFamily 字体
     * @param isBlod     是否加粗
     * @return 返回值：返回字体样式对象
     * @throws Exception
     * @author myclover
     */
    private static RPr getRPr(String fontFamily, boolean isBlod)
    {
        return getRPr(fontFamily, "000000", "18", STHint.EAST_ASIA, isBlod);
    }

    /**
     * 功能描述：设置字体的样式，黑色
     *
     * @param fontFamily     字体
     * @param hpsMeasureSize 字号的大小
     * @param isBlod         是否加粗
     * @return 返回值：返回字体样式对象
     * @throws Exception
     * @author myclover
     */
    private static RPr getRPr(String fontFamily, String hpsMeasureSize, boolean isBlod)
    {
        return getRPr(fontFamily, "000000", hpsMeasureSize, STHint.EAST_ASIA, isBlod);
    }

    /**
     * 功能描述：获取单元格边框样式
     *
     * @param type   单元格类型，0表示无边框，2表示双线边框，其他表示单线边框
     * @param color  边框颜色
     * @param border 边框大小
     * @param space  间距
     * @return 返回值：返回边框对象
     * @author myclover
     */
    private static CTBorder getCTBorder(int type, String color, String border, String space)
    {
        CTBorder ctb = new CTBorder();
        if (type == 0)
        {
            ctb.setVal(STBorder.NIL);
        }
        else
        {
            ctb.setColor(color);
            ctb.setSz(new BigInteger(border));
            ctb.setSpace(new BigInteger(space));
            if (type == 2)
            {
                ctb.setVal(STBorder.DOUBLE);
            }
            else
            {
                ctb.setVal(STBorder.SINGLE);
            }
        }
        return ctb;
    }

    /**
     * 功能描述：设置单元格内容对齐方式
     *
     * @param p             内容
     * @param jcEnumeration 对齐方式
     * @author myclover
     */
    public static void setCellContentStyle(P p, JcEnumeration jcEnumeration)
    {
        PPr pPr = p.getPPr();
        if (pPr == null)
        {
            ObjectFactory factory = Context.getWmlObjectFactory();
            pPr = factory.createPPr();
        }
        org.docx4j.wml.Jc jc = pPr.getJc();
        if (jc == null)
        {
            jc = new org.docx4j.wml.Jc();
        }
        jc.setVal(jcEnumeration);
        pPr.setJc(jc);
        p.setPPr(pPr);
    }

    /**
     * 功能描述：设置单元格内容对齐方式，居中对齐
     *
     * @param p 内容
     * @author myclover
     */
    public static void setCellContentStyle(P p)
    {
        setCellContentStyle(p, JcEnumeration.CENTER);
    }


    /**
     * 功能描述：填充表格内容
     *
     * @param wordPackage    文档处理包对象
     * @param tbl            表格对象
     * @param dataList       表格数据
     * @param titleList      表头数据
     * @param isFixedTitle   是否固定表头
     * @param tFontFamily    表头字体
     * @param tFontSize      表头字体大小
     * @param tIsBlod        表头是否加粗
     * @param tJcEnumeration 表头对齐方式
     * @param fontFamily     表格字体
     * @param fontSize       表格字号
     * @param isBlod         表格内容是否加粗
     * @param jcEnumeration  表格对齐方式
     * @author myclover
     */
    private static void fillTableData(WordprocessingMLPackage wordPackage, Tbl tbl, List dataList, List titleList, boolean isFixedTitle, String tFontFamily, String tFontSize, boolean tIsBlod, JcEnumeration tJcEnumeration, String fontFamily, String fontSize, boolean isBlod, JcEnumeration jcEnumeration)
    {
        List rowList = tbl.getContent();
        //整个表格的行数
        int rows = rowList.size();
        int tSize = 0;
        //表头
        if (titleList != null && titleList.size() > 0)
        {
            tSize = titleList.size();
            for (int t = 0; t < tSize; t++)
            {
                Object[] tobj = (Object[]) titleList.get(t);
                Tr tr0 = (Tr) XmlUtils.unwrap(rowList.get(t));
                List colList = tr0.getContent();
                for (int c = 0; c < colList.size(); c++)
                {
                    Tc tc0 = (Tc) XmlUtils.unwrap(colList.get(c));
                    //填充表头数据
                    fillCellData(tc0, converObjToStr(tobj[c], ""), tFontFamily, tFontSize, tIsBlod, tJcEnumeration);
                }
                if (isFixedTitle)
                {
                    //设置固定表头
                    fixedTitle(tr0);
                }
            }
        }
        int colsSize = 1;
        //表格数据，去掉表头之后的表格数据
        for (int r = tSize; r < rows; r++)
        {
            Tr tr = (Tr) XmlUtils.unwrap(rowList.get(r));
            Object[] objs = null;
            //如果表格内容不为空，则取出相应的数据进行填充
            if (dataList != null && (dataList.size() >= (rows - tSize)))
            {
                objs = (Object[]) dataList.get(r - tSize);
            }
            List colsList = tr.getContent();
            //整个表格的列数
            colsSize = colsList.size();
            for (int i = 0; i < colsSize; i++)
            {
                Tc tc = (Tc) XmlUtils.unwrap(colsList.get(i));
                //填充表格数据
                if (objs != null)
                {
                    fillCellData(tc, converObjToStr(objs[i], ""), fontFamily, fontSize, isBlod, jcEnumeration);
                }
                else
                {
                    fillCellData(tc, "", fontFamily, fontSize, isBlod, jcEnumeration);
                }
            }
        }
    }

    /**
     * 功能描述：填充表格内容，固定表头，表头宋体加粗，小五号，表格内容宋体，小五号，表格居中对齐</BR>
     * 其中表格数据跟表头数据结构要一致，适用于简单的n行m列的普通表格
     *
     * @param wordPackage 文档处理包对象
     * @param tbl         表格对象
     * @param dataList    表格数据
     * @param titleList   表头数据，如果不需要表头信息，则只要传入null即可
     * @author myclover
     */
    public static void fillTableData(WordprocessingMLPackage wordPackage, Tbl tbl, List dataList, List titleList)
    {
        fillTableData(wordPackage, tbl, dataList, titleList, true, "宋体", "18", true, JcEnumeration.CENTER, "宋体", "18", false, JcEnumeration.CENTER);
    }

    /**
     * 功能描述：固定表头
     *
     * @param tr 行对象
     * @author myclover
     */
    public static void fixedTitle(Tr tr)
    {
        ObjectFactory factory = Context.getWmlObjectFactory();
        BooleanDefaultTrue bdt = factory.createBooleanDefaultTrue();
        //表示固定表头
        bdt.setVal(true);
        TrPr trpr = tr.getTrPr();
        if (trpr == null)
        {
            trpr = factory.createTrPr();
        }
        trpr.getCnfStyleOrDivIdOrGridBefore().add(factory.createCTTrPrBaseTblHeader(bdt));
        tr.setTrPr(trpr);
    }

    /**
     * 功能描述：填充单元格内容
     *
     * @param tc            单元格对象
     * @param data          内容
     * @param fontFamily    字体
     * @param fontSize      字号
     * @param isBlod        是否加粗
     * @param jcEnumeration 对齐方式
     * @author myclover
     */
    private static void fillCellData(Tc tc, String data, String fontFamily, String fontSize, boolean isBlod, JcEnumeration jcEnumeration)
    {
        ObjectFactory factory = Context.getWmlObjectFactory();
        org.docx4j.wml.P p = (P) XmlUtils.unwrap(tc.getContent().get(0));
        //设置表格内容的对齐方式
        setCellContentStyle(p, jcEnumeration);
        org.docx4j.wml.Text t = factory.createText();
        t.setValue(data);
        org.docx4j.wml.R run = factory.createR();
        //设置表格内容字体样式
        run.setRPr(getRPr(fontFamily, fontSize, isBlod));
        TcPr tcpr = tc.getTcPr();
        if (tcpr == null)
        {
            tcpr = factory.createTcPr();
        }
        //设置内容垂直居中
        CTVerticalJc valign = factory.createCTVerticalJc();
        valign.setVal(STVerticalJc.CENTER);
        tcpr.setVAlign(valign);
        run.getContent().add(t);
        p.getContent().add(run);
    }

    /**
     * 功能描述：填充单元格内容，小五号，宋体，内容居中
     *
     * @param tc     单元格对象
     * @param data   数据
     * @param isBlod 是否加粗
     * @author myclover
     */
    public static void fillCellData(Tc tc, String data, boolean isBlod)
    {
        fillCellData(tc, data, "宋体", "18", isBlod, JcEnumeration.CENTER);
    }


    /**
     * 功能描述：获取文字信息
     *
     * @param content 文字信息
     * @return 返回值：返回值指定样式的信息
     * @author myclover
     */
    public static String getContentInfo(String content)
    {
        StringBuffer sb = new StringBuffer();
        sb.append("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" >");
        sb.append("<w:pPr><w:spacing w:line=\"360\" w:lineRule=\"auto\"/><w:ind w:firstLineChars=\"200\" w:firstLine=\"480\"/>");
        sb.append("<w:rPr><w:rFonts w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/>");
        sb.append("<w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr></w:pPr><w:r w:rsidRPr=\"00792470\"><w:rPr>");
        sb.append("<w:rFonts w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\" w:hint=\"eastAsia\"/>");
        sb.append("<w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr><w:t>");
        sb.append(content);
        sb.append("</w:t></w:r></w:p>");
        return sb.toString();
    }

    /**
     * 功能描述：合并单元格</BR>
     * 表示合并第startRow（开始行）行中的第startCol（开始列）列到（startCol + colSpan - 1）列 </BR>
     * 表示合并第startCol（开始列）行中的第startRow（开始行）列到（startRow + rowSpan - 1）行
     *
     * @param tc         单元格对象
     * @param currentRow 当前行号，传入的是遍历表格时的行索引参数
     * @param startRow   开始行
     * @param rowSpan    合并的行数，大于1才表示合并
     * @param currentCol 当前列号，传入的是遍历表格时的列索引参数
     * @param startCol   开始列
     * @param colSpan    合并的列数，大于1才表示合并
     * @author myclover
     */
    public static void setCellMerge(Tc tc, int currentRow, int startRow, int rowSpan, int currentCol, int startCol, int colSpan)
    {
        ObjectFactory factory = Context.getWmlObjectFactory();
        TcPr tcpr = tc.getTcPr();
        if (tcpr == null)
        {
            tcpr = factory.createTcPr();
        }
        //表示合并列
        if (colSpan > 1)
        {
            //表示从第startRow行开始
            if (currentRow == startRow)
            {
                //表示从第startRow行的第startCol列开始合并，合并到第startCol + colSpan - 1列
                if (currentCol == startCol)
                {
                    HMerge hm = factory.createTcPrInnerHMerge();
                    hm.setVal("restart");
                    tcpr.setHMerge(hm);
                    tc.setTcPr(tcpr);
                }
                else if (currentCol > startCol && currentCol <= (startCol + colSpan - 1))
                {
                    HMerge hm = factory.createTcPrInnerHMerge();
                    tcpr.setHMerge(hm);
                    tc.setTcPr(tcpr);
                }
            }
        }
        //表示合并行
        if (rowSpan > 1)
        {
            //表示从第startCol列开始
            if (currentCol == startCol)
            {
                //表示从第startCol列的第startRow行始合并，合并到第startRow + rowSpan - 1行
                if (currentRow == startRow)
                {
                    VMerge vm = factory.createTcPrInnerVMerge();
                    vm.setVal("restart");
                    tcpr.setVMerge(vm);
                    tc.setTcPr(tcpr);
                }
                else if (currentRow > startRow && currentRow <= (startRow + rowSpan - 1))
                {
                    VMerge vm = factory.createTcPrInnerVMerge();
                    tcpr.setVMerge(vm);
                    tc.setTcPr(tcpr);
                }
            }
        }

    }

    /**
     * 功能描述：合并单元格，相当于跨列的效果</BR>
     * 表示合并第startRow（开始行）行中的第startCol（开始列）列到（startCol + colSpan - 1）列 </BR>
     *
     * @param tc         单元格对象
     * @param currentRow 当前行号，传入的是遍历表格时的行索引参数
     * @param startRow   开始行
     * @param currentCol 当前列号，传入的是遍历表格时的列索引参数
     * @param startCol   开始列
     * @param colSpan    合并的列数，大于1才表示合并
     * @author myclover
     */
    public static void setCellHMerge(Tc tc, int currentRow, int startRow, int currentCol, int startCol, int colSpan)
    {
        setCellMerge(tc, currentRow, startRow, 1, currentCol, startCol, colSpan);
    }

    /**
     * 功能描述：合并单元格，相当于跨行的效果</BR>
     * 表示合并第startCol（开始列）行中的第startRow（开始行）列到（startRow + rowSpan - 1）行
     *
     * @param tc         单元格对象
     * @param currentRow 当前行号，传入的是遍历表格时的行索引参数
     * @param startRow   开始行
     * @param rowSpan    合并的行数，大于1才表示合并
     * @param currentCol 当前列号，传入的是遍历表格时的列索引参数
     * @param startCol   开始列
     * @author myclover
     */
    public static void setCellVMerage(Tc tc, int currentRow, int startRow, int rowSpan, int currentCol, int startCol)
    {
        setCellMerge(tc, currentRow, startRow, rowSpan, currentCol, startCol, 1);
    }

    /**
     * 功能描述：设置文档是否只读，包括内容和样式
     *
     * @param wordPackage 文档处理包对象
     * @param isReadOnly  是否只读
     * @throws Exception
     * @author myclover
     */
    public static void setReadOnly(WordprocessingMLPackage wordPackage, boolean isReadOnly) throws Exception
    {
        byte[] bt = "".getBytes();
        if (isReadOnly)
        {
            bt = "123456".getBytes();
        }
        ObjectFactory factory = Context.getWmlObjectFactory();
        //创建设置文档对象
        DocumentSettingsPart ds = wordPackage.getMainDocumentPart().getDocumentSettingsPart();
        if (ds == null)
        {
            ds = new DocumentSettingsPart();
        }
        CTSettings cs = ds.getJaxbElement();
        if (cs == null)
        {
            cs = factory.createCTSettings();
        }
        //创建文档保护对象
        CTDocProtect cp = cs.getDocumentProtection();
        if (cp == null)
        {
            cp = new CTDocProtect();
        }
        //设置加密方式
        cp.setCryptProviderType(STCryptProv.RSA_AES);
        cp.setCryptAlgorithmClass(STAlgClass.HASH);
        //设置任何用户
        cp.setCryptAlgorithmType(STAlgType.TYPE_ANY);
        cp.setCryptAlgorithmSid(new BigInteger("4"));
        cp.setCryptSpinCount(new BigInteger("50000"));
        //只读
        if (isReadOnly)
        {
            cp.setEdit(STDocProtect.READ_ONLY);
            cp.setHash(bt);
            cp.setSalt(bt);
            //设置内容不可编辑
            cp.setEnforcement(true);
            //设置格式不可编辑
            cp.setFormatting(true);
        }
        else
        {
            cp.setEdit(STDocProtect.NONE);
            cp.setHash(null);
            cp.setSalt(null);
            //设置内容不可编辑
            cp.setEnforcement(false);
            //设置格式不可编辑
            cp.setFormatting(false);
        }

        cs.setDocumentProtection(cp);
        ds.setJaxbElement(cs);
        //添加到文档主体中
        wordPackage.getMainDocumentPart().addTargetPart(ds);
    }

    /**
     * 功能描述：设置文档是否只读，包括内容和样式
     *
     * @param fileName   文件
     * @param isReadOnly 是否只读
     * @return 返回值：设置成功，则返回true，否则返回false
     * @throws Exception
     * @author myclover
     */
    public static boolean setReadOnly(String fileName, boolean isReadOnly) throws Exception
    {
        try
        {
            File file = new File(fileName);
            if (!file.exists())
            {
                return false;
            }
            //加载需要设置只读的文件
            WordprocessingMLPackage wordPackage = WordprocessingMLPackage.load(file);
            //设置只读
            setReadOnly(wordPackage, isReadOnly);
            //保存文件
            saveWordPackage(wordPackage, file);
        } catch (Exception e)
        {
            e.printStackTrace();
            return false;
        }
        return true;
    }

    /**
     * 功能描述：插入页脚条形码
     *
     * @param wordPackage 文档处理包对象
     * @param code        条形码编码
     * @author myclover
     */
    public static void insertFooterImage(WordprocessingMLPackage wordPackage, String code) throws Exception
    {
        try
        {
            createFooterReference(wordPackage, createFooterPart(wordPackage, code));
        } catch (Exception e)
        {
            e.printStackTrace();
            throw new Exception("insertFooterImage 插入页脚的条形码失败：", e);
        }
    }

    /**
     * 功能描述：创建页脚信息
     *
     * @param wordPackage  文档处理包对象
     * @param relationship 关联对象
     * @throws Exception
     * @author myclover
     */
    private static void createFooterReference(WordprocessingMLPackage wordPackage, Relationship relationship) throws Exception
    {
        org.docx4j.wml.ObjectFactory factory = new org.docx4j.wml.ObjectFactory();
        List<SectionWrapper> sections = wordPackage.getDocumentModel().getSections();
        SectPr sectPr = sections.get(sections.size() - 1).getSectPr();
        if (sectPr == null)
        {
            sectPr = factory.createSectPr();
            wordPackage.getMainDocumentPart().addObject(sectPr);
            sections.get(sections.size() - 1).setSectPr(sectPr);
        }
        FooterReference headerReference = factory.createFooterReference();
        headerReference.setId(relationship.getId());
        headerReference.setType(HdrFtrRef.DEFAULT);
        sectPr.getEGHdrFtrReferences().add(headerReference);
    }

    /**
     * 功能描述：创建页脚部分
     *
     * @param wordPackage 文档处理包对象
     * @param code        条形码编码
     * @return
     * @throws Exception
     * @author myclover
     */
    private static Relationship createFooterPart(WordprocessingMLPackage wordPackage, String code) throws Exception
    {
        FooterPart footerPart = new FooterPart();
        footerPart.setPackage(wordPackage);
        footerPart.setJaxbElement(getFtr(wordPackage, footerPart, code));
        return wordPackage.getMainDocumentPart().addTargetPart(footerPart);
    }

    /**
     * 功能描述：获取页脚内容
     *
     * @param wordPackage 文档处理包对象
     * @param sourcePart  源部分
     * @param code        条形码编码
     * @return
     * @throws Exception
     * @author myclover
     */
    private static Ftr getFtr(WordprocessingMLPackage wordPackage, Part sourcePart, String code) throws Exception
    {
        org.docx4j.wml.ObjectFactory factory = new org.docx4j.wml.ObjectFactory();
        Ftr ftr = factory.createFtr();
        byte[] bt = generateCode39Barcode(code);
        ftr.getContent().add(getPImage(wordPackage, sourcePart, bt, "filename", "alttext", 1, 2));
        return ftr;
    }

    /**
     * 功能描述：创建页脚条形码图片信息
     *
     * @param wordPackage  文档处理包对象
     * @param sourcePart   源部分
     * @param bytes        字节数组
     * @param filenameHint 图片名称
     * @param altText      提示信息
     * @param id1          id1
     * @param id2          id2
     * @return
     * @throws Exception
     * @author myclover
     */
    private static org.docx4j.wml.P getPImage(WordprocessingMLPackage wordPackage, Part sourcePart, byte[] bytes, String filenameHint, String altText, int id1, int id2) throws Exception
    {
        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordPackage, sourcePart, bytes);
        Inline inline = imagePart.createImageInline(filenameHint, altText, id1, id2, false);
        org.docx4j.wml.ObjectFactory factory = new org.docx4j.wml.ObjectFactory();
        org.docx4j.wml.P p = factory.createP();
        org.docx4j.wml.R run = factory.createR();
        p.getContent().add(run);
        org.docx4j.wml.Drawing drawing = factory.createDrawing();
        run.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);
        PPr pPr = p.getPPr();
        if (pPr == null)
        {
            pPr = factory.createPPr();
        }
        Jc jc = pPr.getJc();
        if (jc == null)
        {
            jc = new org.docx4j.wml.Jc();
        }
        //页脚条形码所处位置
        jc.setVal(JcEnumeration.RIGHT);
        pPr.setJc(jc);
        p.setPPr(pPr);
        return p;
    }

    /**
     * 功能描述：生成条形码
     *
     * @param code 条形码编码
     * @return
     * @throws Exception
     * @author myclover
     */
    private static byte[] generateCode39Barcode(String code) throws Exception
    {
        /*Code39Bean bean = new Code39Bean();
	    //每英寸所打印的点数或线数，用来表示打印机打印分辨率。   
	    final int dpi = 200;   
	    bean.setModuleWidth(0.15);   
	    bean.setHeight(10);   
	    bean.setWideFactor(3);   
	    bean.doQuietZone(true);   
	    ByteArrayOutputStream out = null;
	    try {
			out = new ByteArrayOutputStream();
			BitmapCanvasProvider canvas = new BitmapCanvasProvider(dpi,BufferedImage.TYPE_BYTE_GRAY, true, 0);   
	        bean.generateBarcode(canvas, code);   
	        canvas.finish();   
	        BufferedImage barcodeImage = canvas.getBufferedImage();   
	        ImageIO.write(barcodeImage, "jpeg", out);  
	        out.flush();
			return out.toByteArray();
		} catch (Exception e) {
			e.printStackTrace();
		}finally{
			if(out != null){
				try {
					out.close();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		}*/
        return null;
    }

    /**
     * 功能描述：Object数据转换为String类型
     *
     * @param obj
     * @param defaultStr 如果obj对象为空，则返回的值
     * @return
     * @author myclover
     */
    public static String converObjToStr(Object obj, String defaultStr)
    {
        if (obj != null)
        {
            return obj.toString();
        }
        return defaultStr;
    }

    private static void addRowToTable(Tbl reviewtable, Tr templateRow, Map<String, String> replacements)
    {
        Tr workingRow = (Tr) XmlUtils.deepCopy(templateRow);
        List<?> textElements = getAllElementFromObject(workingRow, Text.class);
        for (Object object : textElements)
        {
            Text text = (Text) object;
            String replacementValue = (String) replacements.get(text.getValue());
            if (replacementValue != null)
            {
                text.setValue(replacementValue);
            }
        }

        reviewtable.getContent().add(workingRow);
    }

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

    private void replaceTable(String[] placeholders, List<Map<String, String>> textToAdd, WordprocessingMLPackage template) throws Docx4JException, JAXBException
    {
        List<Object> tables = getAllElementFromObject(template.getMainDocumentPart(), Tbl.class);

        // 1. find the table
        Tbl tempTable = getTemplateTable(tables, placeholders[0]);
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
}

