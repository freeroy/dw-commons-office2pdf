package org.developerworld.commons.office2pdf.impl.jacob;

import java.io.File;

import org.developerworld.commons.office2pdf.Convertor;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;

/**
 * excel转换pdf
 * 
 * @author Roy Huang
 *
 */
public class JacobExcel2PdfConvertor implements Convertor {

	public void conver2Pdf(File origFile, File destFile) {
		conver2Pdf(origFile.getPath(), destFile.getPath());
	}

	public void conver2Pdf(String origFilePath, String destFilePath) {
		// 打开应用程序
		ActiveXComponent axc = new ActiveXComponent("Excel.Application");
		// 设置不可见
		axc.setProperty("Visible", false);
		// 获得所有打开的文档,返回对象
		Dispatch excels = axc.getProperty("Workbooks").toDispatch();
		// 调用对象中Open方法打开文档，并返回打开的文档对象
		Dispatch excel = Dispatch.call(excels, "Open", origFilePath, false, true).toDispatch();
		// 保存为pdf格式宏，值为0
		Dispatch.call(excel, "ExportAsFixedFormat", 0, destFilePath);
		// 关闭文档
		Dispatch.call(excel, "Close", false);
		// 关闭word应用程序
		axc.invoke("Quit");
	}

}
