package org.developerworld.commons.office2pdf.impl.jodconverter;

import java.io.File;

import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeConnectionProtocol;
import org.artofsolving.jodconverter.office.OfficeManager;
import org.artofsolving.jodconverter.process.ProcessManager;
import org.developerworld.commons.office2pdf.Convertor;
import org.developerworld.commons.office2pdf.utils.Utils;

/**
 * JobConvertor方式的pdf装换器
 * 
 * @author Roy Huang
 *
 */
public class JobConvertor implements Convertor {

	private DefaultOfficeManagerConfiguration officeManagerConfiguration;

	private OfficeManager officeManager;

	private OfficeDocumentConverter officeDocumentConverter;

	public void setOfficeManagerConfiguration(DefaultOfficeManagerConfiguration officeManagerConfiguration) {
		this.officeManagerConfiguration = officeManagerConfiguration;
	}

	public void setOfficeManager(OfficeManager officeManager) {
		this.officeManager = officeManager;
	}

	public void setOfficeDocumentConverter(OfficeDocumentConverter officeDocumentConverter) {
		this.officeDocumentConverter = officeDocumentConverter;
	}

	/**
	 * 获取配置
	 * 
	 * @return
	 */
	private DefaultOfficeManagerConfiguration getOfficeManagerConfiguration() {
		if (officeManagerConfiguration == null)
			officeManagerConfiguration = new DefaultOfficeManagerConfiguration();
		return officeManagerConfiguration;
	}

	/**
	 * 获取管理器
	 * 
	 * @return
	 */
	private OfficeManager getOfficeManager() {
		if (officeManager == null)
			officeManager = getOfficeManagerConfiguration().buildOfficeManager();
		return officeManager;
	}

	/**
	 * 获取转换器
	 * 
	 * @return
	 */
	private OfficeDocumentConverter getOfficeDocumentConverter() {
		if (officeDocumentConverter == null)
			officeDocumentConverter = new OfficeDocumentConverter(getOfficeManager());
		return officeDocumentConverter;
	}

	public JobConvertor setOfficeHome(String officeHome) throws NullPointerException, IllegalArgumentException {
		getOfficeManagerConfiguration().setOfficeHome(officeHome);
		return this;
	}

	public JobConvertor setOfficeHome(File officeHome) throws NullPointerException, IllegalArgumentException {
		getOfficeManagerConfiguration().setOfficeHome(officeHome);
		return this;
	}

	public JobConvertor setConnectionProtocol(OfficeConnectionProtocol connectionProtocol) throws NullPointerException {
		getOfficeManagerConfiguration().setConnectionProtocol(connectionProtocol);
		return this;
	}

	public JobConvertor setPortNumber(int portNumber) {
		getOfficeManagerConfiguration().setPortNumber(portNumber);
		return this;
	}

	public JobConvertor setPortNumbers(int... portNumbers) throws NullPointerException, IllegalArgumentException {
		getOfficeManagerConfiguration().setPortNumbers(portNumbers);
		return this;
	}

	public JobConvertor setPipeName(String pipeName) throws NullPointerException {
		getOfficeManagerConfiguration().setPipeName(pipeName);
		return this;
	}

	public JobConvertor setPipeNames(String... pipeNames) throws NullPointerException, IllegalArgumentException {
		getOfficeManagerConfiguration().setPipeNames(pipeNames);
		return this;
	}

	public JobConvertor setRunAsArgs(String... runAsArgs) {
		getOfficeManagerConfiguration().setRunAsArgs(runAsArgs);
		return this;
	}

	public JobConvertor setTemplateProfileDir(File templateProfileDir) throws IllegalArgumentException {
		getOfficeManagerConfiguration().setTemplateProfileDir(templateProfileDir);
		return this;
	}

	public JobConvertor setWorkDir(File workDir) {
		getOfficeManagerConfiguration().setWorkDir(workDir);
		return this;
	}

	public JobConvertor setTaskQueueTimeout(long taskQueueTimeout) {
		getOfficeManagerConfiguration().setTaskQueueTimeout(taskQueueTimeout);
		return this;
	}

	public JobConvertor setTaskExecutionTimeout(long taskExecutionTimeout) {
		getOfficeManagerConfiguration().setTaskExecutionTimeout(taskExecutionTimeout);
		return this;
	}

	public JobConvertor setMaxTasksPerProcess(int maxTasksPerProcess) {
		getOfficeManagerConfiguration().setMaxTasksPerProcess(maxTasksPerProcess);
		return this;
	}

	public JobConvertor setProcessManager(ProcessManager processManager) throws NullPointerException {
		getOfficeManagerConfiguration().setProcessManager(processManager);
		return this;
	}

	public JobConvertor setRetryTimeout(long retryTimeout) {
		getOfficeManagerConfiguration().setRetryTimeout(retryTimeout);
		return this;
	}

	/**
	 * 启动后台进程
	 * 
	 * @return
	 */
	public JobConvertor startService() {
		if (!getOfficeManager().isRunning())
			getOfficeManager().start();
		return this;
	}

	/**
	 * 关闭后台进程
	 * 
	 * @return
	 */
	public JobConvertor stopService() {
		if (getOfficeManager().isRunning())
			getOfficeManager().stop();
		return this;
	}

	public void conver2Pdf(File origFile, File destFile) {
		Utils.checkFileArgs(origFile, destFile);
		getOfficeDocumentConverter().convert(origFile, destFile);
	}

	public void conver2Pdf(String origFilePath, String destFilePath) {
		Utils.checkFileArgs(origFilePath, destFilePath);
		getOfficeDocumentConverter().convert(new File(origFilePath), new File(destFilePath));
	}

}
