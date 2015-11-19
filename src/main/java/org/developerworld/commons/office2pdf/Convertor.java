package org.developerworld.commons.office2pdf;

import java.io.File;

/**
 * PDF转换器
 * 
 * @author Roy Huang
 *
 */
public interface Convertor {

	/**
	 * 转为pdf文件
	 * 
	 * @param origFile
	 *            转换来源
	 * @param destFile
	 *            目标输出
	 */
	public void conver2Pdf(File origFile, File destFile);

	/**
	 * 转为pdf文件
	 * 
	 * @param origFilePath
	 *            转换来源路径
	 * @param destFilePath
	 *            目标输出路径
	 */
	public void conver2Pdf(String origFilePath, String destFilePath);
}
