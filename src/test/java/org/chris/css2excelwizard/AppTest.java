package org.chris.css2excelwizard;

import junit.framework.Test;
import junit.framework.TestCase;
import junit.framework.TestSuite;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

/**
 * Unit test for simple App.
 */
public class AppTest
		extends TestCase
{
	/**
	 * Create the test case
	 *
	 * @param testName name of the test case
	 */
	public AppTest(String testName)
	{
		super(testName);
	}

	/**
	 * @return the suite of tests being tested
	 */
	public static Test suite()
	{
		return new TestSuite(AppTest.class);
	}

	/**
	 * Rigourous Test :-)
	 */
	public void testApp()
	{

		SheetMap bitmap = new SheetMap();

		// 设置某些位置为true
		bitmap.set(2, 3, true);
		bitmap.set(4, 5, true);
		bitmap.set(10, 12, true);

		// 查询位置的值
		System.out.println(bitmap.get(2, 3));  // 输出: true
		System.out.println(bitmap.get(4, 5));  // 输出: true
		System.out.println(bitmap.get(10, 12));  // 输出: true

		bitmap.print();

		// 查询超出范围的位置
		System.out.println(bitmap.get(0, 0));  // 输出: false
		System.out.println(bitmap.get(15, 15));  // 输出: false
		System.out.println(bitmap.get(20, 12));  // 输出: false

		bitmap.set(10, 12, true);


		// 设置超出范围的位置
		bitmap.set(20, 20, true);

		// 再次查询超出范围的位置
		System.out.println(bitmap.get(20, 20));  // 输出: true
		System.out.println(bitmap.get(40, 20));  // 输出: true

		bitmap.print();
	}

	public void testExcel()
	{
		Workbook workbook = new XSSFWorkbook();
		Css2ExcelWizard css2ExcelWizard = new Css2ExcelWizard(workbook);

		css2ExcelWizard.css.defineRootStyle("font-family: Calibri; font-size: 10;border-color: red; border-style: thin");

		css2ExcelWizard.css.defineStyle("style1", "color: red; background-color: yellow; vertical-align: top");
		css2ExcelWizard.css.defineStyle("style2", "font:Calibri 26 green bold italic;align: left middle;editable: lock;indention: 1; number-format: 0.0%; font-style: underline-double");
		css2ExcelWizard.css.defineStyle("style3", "font-size: 26;color: blue; font-weight: bold;font-style: normal;border-color: red; border-style:double;text-align: right");

		css2ExcelWizard.createSheet("Test");
		css2ExcelWizard.layout.setWidthPadding(0.9f);
		css2ExcelWizard.layout.setWidth(30, 15, 15, 20);
		css2ExcelWizard.layout
				.newRow(20).cell().cell("style1", "Test Text")
				.newRow(80).cell("style2", 0.54)
				.newRow().cellSpan(2, 2, "style3", "TEST3")
				.newRow().cell(null, "start");


		css2ExcelWizard.layout.appendMergedCells();


		try (FileOutputStream fileOutputStream = new FileOutputStream(new File("/Users/chris/Downloads/test.xlsx")))
		{
			workbook.write(fileOutputStream);
		}
		catch (IOException e)
		{
			throw new RuntimeException(e);
		}

	}
}
