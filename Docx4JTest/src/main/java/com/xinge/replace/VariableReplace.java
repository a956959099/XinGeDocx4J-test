package com.xinge.replace;

import java.util.HashMap;
import java.util.Map;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

/**
 * Created by AX on 2016/3/8.
 */
public class VariableReplace
{


    public static void main(String[] args) throws Exception
    {

        // Exclude context init from timing
        org.docx4j.wml.ObjectFactory foo = Context.getWmlObjectFactory();

        // Input docx has variables in it: ${colour}, ${icecream}
        String inputfilepath =  "D:\\模板.docx";

        boolean save = true;
        String outputfilepath =  "D:\\OUT_VariableReplace.docx";

        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File(inputfilepath));
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

        HashMap<String, String> mappings = new HashMap<String, String>();
       // Map<String, Object> params = new HashMap<String, Object>();
        mappings.put("myTable","我的表格" );
        mappings.put("name", "小宝");
        mappings.put("age", "xx");
        mappings.put("sex", "男");
        mappings.put("job", "肉盾");
        mappings.put("hobby", "电商");
        mappings.put("phone", "1717");
        //mappings.put("icecream", "chocolate");

        long start = System.currentTimeMillis();

        // Approach 1 (from 3.0.0; faster if you haven't yet caused unmarshalling to occur):

        documentPart.variableReplace(mappings);

/*		// Approach 2 (original)

			// unmarshallFromTemplate requires string input
			String xml = XmlUtils.marshaltoString(documentPart.getJaxbElement(), true);
			// Do it...
			Object obj = XmlUtils.unmarshallFromTemplate(xml, mappings);
			// Inject result into docx
			documentPart.setJaxbElement((Document) obj);
*/

        long end = System.currentTimeMillis();
        long total = end - start;
        System.out.println("Time: " + total);

        // Save it
        if (save)
        {
            SaveToZipFile saver = new SaveToZipFile(wordMLPackage);
            saver.save(outputfilepath);
        }
        else
        {
            System.out.println(XmlUtils.marshaltoString(documentPart.getJaxbElement(), true, true));
        }
    }

}
