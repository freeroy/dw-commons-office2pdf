package org.developerworld.commons.office2pdf.impl.jacob;

import java.io.File;
import java.net.URL;

import org.apache.commons.io.FilenameUtils;
import org.developerworld.commons.office2pdf.Convertor;
import org.junit.Assert;
import org.junit.Test;

public class PdfConvertorTest {

	//@Test
	public void testConver2Pdf() {
		testWord2Pdf();
		testExcel2Pdf();
		testPowerPoint2Pdf();
	}

	private void testPowerPoint2Pdf() {
		Convertor convertor = buildPowerPointPdfConvertor();
		URL origFileUrl = this.getClass().getClassLoader().getResource("powerpoint.ppt");
		File origFile = new File(origFileUrl.getPath());
		File destFile = new File(
				origFile.getParent() + File.separator + FilenameUtils.getBaseName(origFile.getName()) + ".pdf");
		convertor.conver2Pdf(origFile, destFile);
		Assert.assertTrue(true);
	}

	private Convertor buildPowerPointPdfConvertor() {
		JacobPowerPoint2PdfConvertor rst = new JacobPowerPoint2PdfConvertor();
		return rst;
	}

	private void testWord2Pdf() {
		Convertor convertor = buildWordPdfConvertor();
		URL origFileUrl = this.getClass().getClassLoader().getResource("word.doc");
		File origFile = new File(origFileUrl.getPath());
		File destFile = new File(
				origFile.getParent() + File.separator + FilenameUtils.getBaseName(origFile.getName()) + ".pdf");
		convertor.conver2Pdf(origFile, destFile);
		Assert.assertTrue(true);
	}

	private Convertor buildWordPdfConvertor() {
		JacobWord2PdfConvertor rst = new JacobWord2PdfConvertor();
		return rst;
	}

	private void testExcel2Pdf() {
		Convertor convertor = buildExcelPdfConvertor();
		URL origFileUrl = this.getClass().getClassLoader().getResource("excel.xls");
		File origFile = new File(origFileUrl.getPath());
		File destFile = new File(
				origFile.getParent() + File.separator + FilenameUtils.getBaseName(origFile.getName()) + ".pdf");
		convertor.conver2Pdf(origFile, destFile);
		Assert.assertTrue(true);
	}

	private Convertor buildExcelPdfConvertor() {
		JacobExcel2PdfConvertor rst = new JacobExcel2PdfConvertor();
		return rst;
	}
}
