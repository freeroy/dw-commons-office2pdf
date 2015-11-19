package org.developerworld.commons.office2pdf.impl.jodconverter;

import java.io.File;
import java.net.URL;

import org.apache.commons.io.FilenameUtils;
import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

public class PdfConvertorTest {
	
	private JobConvertor convertor;
	
	@Before
	public void before(){
		convertor=new JobConvertor().setOfficeHome("D:\\Program Files (x86)\\OpenOffice 4");
		convertor.startService();
	}
	
	@After
	public void after(){
		convertor.stopService();
		convertor=null;
	}
	
	//@Test
	public void testConver2Pdf() {
		testWord2Pdf();
		testExcel2Pdf();
		testPowerPoint2Pdf();
	}

	private void testWord2Pdf() {
		URL origFileUrl = this.getClass().getClassLoader().getResource("word.doc");
		File origFile = new File(origFileUrl.getPath());
		File destFile = new File(
				origFile.getParent() + File.separator + FilenameUtils.getBaseName(origFile.getName()) + ".pdf");
		convertor.conver2Pdf(origFile, destFile);
		Assert.assertTrue(true);
	}

	private void testExcel2Pdf() {
		URL origFileUrl = this.getClass().getClassLoader().getResource("excel.xls");
		File origFile = new File(origFileUrl.getPath());
		File destFile = new File(
				origFile.getParent() + File.separator + FilenameUtils.getBaseName(origFile.getName()) + ".pdf");
		convertor.conver2Pdf(origFile, destFile);
		Assert.assertTrue(true);
	}

	private void testPowerPoint2Pdf() {
		URL origFileUrl = this.getClass().getClassLoader().getResource("powerpoint.ppt");
		File origFile = new File(origFileUrl.getPath());
		File destFile = new File(
				origFile.getParent() + File.separator + FilenameUtils.getBaseName(origFile.getName()) + ".pdf");
		convertor.conver2Pdf(origFile, destFile);
		Assert.assertTrue(true);
	}

}
