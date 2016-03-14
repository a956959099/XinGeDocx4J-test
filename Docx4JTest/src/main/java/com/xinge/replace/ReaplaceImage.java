package com.xinge.replace;

import org.docx4j.XmlUtils;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import java.io.*;
import java.math.BigInteger;
import java.util.*;

import static com.sun.xml.internal.fastinfoset.alphabet.BuiltInRestrictedAlphabets.table;

/**
 * Created by AX on 2016/3/8.
 */
public class ReaplaceImage
{
    public static void main(String[] args) throws Exception
    {


        // Input docx has variables in it: ${colour}, ${icecream}
        String inputfilepath = "D:\\month (1).docx";

        String outputfilepath = "D:\\OUT_VariableReplace.docx";


        WordprocessingMLPackage wordMLPackage = getTemplate(inputfilepath);

        org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
        List allElemtns = new ArrayList();
        List elemetns = getAllElementFromObject(wordMLPackage.getMainDocumentPart(), P.class);
        allElemtns.addAll(elemetns);
        elemetns = getAllElementFromObject(wordMLPackage.getMainDocumentPart(), Tbl.class);
        allElemtns.addAll(elemetns);
        for (Object obj : allElemtns)
        {
            if (obj instanceof Tbl)
            {
                Tbl table = (Tbl) obj;
                List rows = getAllElementFromObject(table, Tr.class);
                for (Object trObj : rows)
                {
                    Tr tr = (Tr) trObj;
                    List cols = getAllElementFromObject(tr, Tc.class);
                    for (Object tcObj : cols)
                    {
                        Tc tc = (Tc) tcObj;
                        List texts = getAllElementFromObject(tc, Text.class);
                        for (Object textObj : texts)
                        {
                            Text text = (Text) textObj;
                            if (text.getValue().contains("${name}"))
                            {
                                File file = new File("D:\\pict\\7.jpg");
                                P paragraphWithImage = addInlineImageToParagraph(createInlineImage(file, wordMLPackage));
                                tc.getContent().remove(0);
                                tc.getContent().add(paragraphWithImage);
                            }
                        }
                        System.out.println("here");
                    }
                }
                System.out.println("here");
            }
            else if (obj instanceof P)
            {
                //段落
                P p = (P) obj;
                {
                    if (p.toString().contains("${nameTest}"))
                    {
                        File file = new File("D:\\pict\\7.jpg");
                        P paragraphWithImage = addInlineImageToParagraph(createInlineImage(file, wordMLPackage));
                        p.getContent().removeAll(p.getContent());
                        //wordMLPackage.getMainDocumentPart().addObject(paragraphWithImage);

                        Map<String,String> repl1 = new HashMap<String, String>();
                        repl1.put("SJ_FUNCTION", "function1");
                        repl1.put("SJ_DESC", "desc1");
                        repl1.put("SJ_PERIOD", "period1");

                        Map<String,String> repl2 = new HashMap<String, String>();
                        repl2.put("SJ_FUNCTION", "function2");
                        repl2.put("SJ_DESC", "desc2");
                        repl2.put("SJ_PERIOD", "period2");

                        Map<String,String> repl3 = new HashMap<String, String>();
                        repl3.put("SJ_FUNCTION", "function3");
                        repl3.put("SJ_DESC", "desc3");
                        repl3.put("SJ_PERIOD", "period3");

                      //  Docx4jUtils.replaceTable(new String[]{"SJ_FUNCTION","SJ_DESC","SJ_PERIOD"}, Arrays.asList(repl1,repl2,repl3), wordMLPackage);
                    }
                }
            }

        }

        wordMLPackage.save(new java.io.File(outputfilepath));
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

    private static P addInlineImageToParagraph(Inline inline)
    {
        // Now add the in-line image to a paragraph
        ObjectFactory factory = new ObjectFactory();
        P paragraph = factory.createP();
        R run = factory.createR();
        paragraph.getContent().add(run);
        Drawing drawing = factory.createDrawing();
        run.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);
        return paragraph;
    }


    private static Inline createInlineImage(File file, WordprocessingMLPackage wordMLPackage) throws Exception
    {
        byte[] bytes = convertImageToByteArray(file);

        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);

        int docPrId = 1;
        int cNvPrId = 2;

        return imagePart.createImageInline("Filename hint", "Alternative text", docPrId, cNvPrId, false);
    }


    private static byte[] convertImageToByteArray1(File file) throws FileNotFoundException, IOException
    {
        InputStream is = new FileInputStream(file);
        long length = file.length();
        // You cannot create an array using a long, it needs to be an int.
        if (length > Integer.MAX_VALUE)
        {
            System.out.println("File too large!!");
        }
        byte[] bytes = new byte[(int) length];
        int offset = 0;
        int numRead = 0;
        while (offset == 0)
        {
            offset += numRead;
        }
        // Ensure all the bytes have been read
        if (offset < bytes.length)
        {
            System.out.println("Could not completely read file " + file.getName());
        }
        is.close();
        return bytes;
    }

    private static byte[] convertImageToByteArray(File file) throws FileNotFoundException, IOException
    {
        InputStream is = new FileInputStream(file);
        long length = file.length();
        // You cannot create an array using a long, it needs to be an int.
        if (length > Integer.MAX_VALUE)
        {
            System.out.println("File too large!!");
        }
        byte[] bytes = new byte[(int) length];
        int offset = 0;
        int numRead = 0;
        while (offset < bytes.length && (numRead = is.read(bytes, offset, bytes.length - offset)) >= 0)
        {
            offset += numRead;
        }
        // Ensure all the bytes have been read
        if (offset < bytes.length)
        {
            System.out.println("Could not completely read file " + file.getName());
        }
        is.close();
        return bytes;
    }

    private static WordprocessingMLPackage getTemplate(String name) throws Docx4JException, FileNotFoundException
    {
        WordprocessingMLPackage template = WordprocessingMLPackage.load(new FileInputStream(new File(name)));
        return template;
    }

    private static List getAllElementFromObject(Object obj, Class toSearch)
    {
        List result = new ArrayList();
        if (obj instanceof JAXBElement)
        {
            obj = ((JAXBElement) obj).getValue();
        }

        if (obj.getClass().equals(toSearch))
        {
            result.add(obj);
        }
        else if (obj instanceof ContentAccessor)
        {
            List children = ((ContentAccessor) obj).getContent();
            for (Object child : children)
            {
                result.addAll(getAllElementFromObject(child, toSearch));
            }

        }
        return result;
    }
}
