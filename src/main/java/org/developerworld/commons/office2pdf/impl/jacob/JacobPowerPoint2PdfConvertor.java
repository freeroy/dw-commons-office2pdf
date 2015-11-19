package org.developerworld.commons.office2pdf.impl.jacob;

import java.io.File;

import org.developerworld.commons.office2pdf.Convertor;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;

/**
 * ppt转换pdf
 * 
 * @author Roy Huang
 *
 */
public class JacobPowerPoint2PdfConvertor implements Convertor {

	public void conver2Pdf(File origFile, File destFile) {
		conver2Pdf(origFile.getPath(), destFile.getPath());
	}

	public void conver2Pdf(String origFilePath, String destFilePath) {
		// 打开应用程序
		ActiveXComponent axc = new ActiveXComponent("PowerPoint.Application");
		// 获得所有打开的文档,返回对象
		Dispatch ppts = axc.getProperty("Presentations").toDispatch();
		// 调用对象中Open方法打开文档，并返回打开的文档对象
		Dispatch ppt = Dispatch.call(ppts, "Open", origFilePath, true// 只读模式
				, true// untitled制定文件是否有标题
				, false// withWindow制定文件是否可见
		).toDispatch();
		// 保存为pdf格式宏，值为0
		Dispatch.call(ppt, "SaveAs", destFilePath, 32);
		// 关闭文档
		Dispatch.call(ppt, "Close");
		// 关闭word应用程序
		axc.invoke("Quit");
	}

}
