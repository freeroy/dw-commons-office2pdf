package org.developerworld.commons.office2pdf.utils;

import java.io.File;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang.StringUtils;

public class Utils {

	/**
	 * 检查输入文件参数是否正确
	 * 
	 * @param origFile
	 * @param destFile
	 */
	public static void checkFileArgs(File origFile, File destFile) {
		if (origFile == null || destFile == null)
			throw new RuntimeException("origFile&destFile must not null!");
		if (!origFile.exists())
			throw new RuntimeException("origFile is not found!");
		else if (!origFile.isFile())
			throw new RuntimeException("origFile is not a file!");
		else if (!"pdf".equalsIgnoreCase(FilenameUtils.getExtension(destFile.getName())))
			throw new RuntimeException("destFile must is pdf file!");
		else if (destFile.exists())
			destFile.delete();
	}

	/**
	 * 检查输入文件参数是否正确
	 * 
	 * @param origFilePath
	 * @param destFilePath
	 */
	public static void checkFileArgs(String origFilePath, String destFilePath) {
		if (StringUtils.isBlank(origFilePath) || StringUtils.isBlank(destFilePath))
			throw new RuntimeException("origFilePath&destFilePath must not blank!");
		checkFileArgs(new File(origFilePath), new File(destFilePath));
	}
}
