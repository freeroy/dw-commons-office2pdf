package org.developerworld.commons.office2pdf.impl.jacob;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang.StringUtils;
import org.developerworld.commons.office2pdf.Convertor;
import org.developerworld.commons.office2pdf.utils.Utils;

/**
 * jacob实现的转换器
 * 
 * @author Roy Huang
 *
 */
public class JacobConvertor implements Convertor {

	private static Map<String, Convertor> convertors = new HashMap<String, Convertor>();

	static {
		convertors.put("ppt", new JacobPowerPoint2PdfConvertor());
		convertors.put("pptx", convertors.get("ppt"));
		convertors.put("xls", new JacobExcel2PdfConvertor());
		convertors.put("xlsx", convertors.get("xls"));
		convertors.put("doc", new JacobWord2PdfConvertor());
		convertors.put("docx", convertors.get("doc"));
	}

	public void conver2Pdf(File origFile, File destFile) {
		Utils.checkFileArgs(origFile, destFile);
		String ext = FilenameUtils.getExtension(origFile.getName());
		if (StringUtils.isBlank(ext))
			throw new RuntimeException("origFile ext can not blank");
		else if (!convertors.containsKey(ext.toLowerCase()))
			throw new RuntimeException("origFile ext:" + ext + " can not conver!");
		convertors.get(ext.toLowerCase()).conver2Pdf(origFile, destFile);
	}

	public void conver2Pdf(String origFilePath, String destFilePath) {
		Utils.checkFileArgs(origFilePath, destFilePath);
		String ext = FilenameUtils.getExtension(origFilePath);
		if (StringUtils.isBlank(ext))
			throw new RuntimeException("origFile ext can not blank");
		else if (!convertors.containsKey(ext.toLowerCase()))
			throw new RuntimeException("origFile ext:" + ext + " can not conver!");
		convertors.get(ext.toLowerCase()).conver2Pdf(origFilePath, origFilePath);
	}

}
