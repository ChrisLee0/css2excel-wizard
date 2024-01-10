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

	public void testExcel() throws IOException
	{
		Css2ExcelWizard css2ExcelWizard = new Css2ExcelWizard();
		css2ExcelWizard.createSheet("SVP");
		css2ExcelWizard.getLayout().setWidth(2, 12, 8, 8, 8, 10); //define root class, all class will default inherit properties from here
		css2ExcelWizard.getCss().defineRootStyle("font-family: Calibri; font-size: 10;border-color: black;;vertical-align: middle");

		//define a class called 'cell'
		css2ExcelWizard.getCss().defineStyle("cell", "text-align: center");
		//define a class called 'heading' which inherit properties from class 'cell'
		css2ExcelWizard.getCss().extendStyle("cell", "heading", "font-weight: bold; font-size: 14;")
				.extendStyle("cell", "header", "background-color: red; color: white; font-weight: bold; ")
				//defined a style called 'rate' with number with 2 decimals
				.extendStyle("cell", "rate", "number-format: 0.00").extendStyle("cell", "grey-cell", "background-color: rgb(191,191,191)");
		css2ExcelWizard.getCss().defineStyle("note", "text-align: left;border: none; wrap-text: true");
		//defined a style with formatted date
		css2ExcelWizard.getCss().defineStyle("date", "text-align: left;border: none; font-weight: bold;date-format: dd-mmm-yyyy");

		css2ExcelWizard.getLayout().newRow()
				.newRow().cell().cellSpan(1, 5, "heading", "Streamlined Variable Pay")
				.newRow().cell().cell("header", "Ratings", "Top", "Strong", "Good", "Inconsistent")
				.newRow().cell().cell("header", "Role Model").cell("rate", 3, 2, 1.25).cell("grey-cell")
				.newRow().cell().cell("header", "Strong").cell("rate", 2.5, 1.5, 1).cell("grey-cell")
				.newRow().cell().cell("header", "Good").cell("rate", 2.25, 1.25, 0.75).cell("grey-cell")
				.newRow().cell().cell("header", "Unacceptable").cell("grey-cell").cell("grey-cell").cell("grey-cell").cell("grey-cell").newRow()
				//create a row with height 35, which takes 1 row and 7 cols, a rich text inside. //grammar: normal text {fontId|text....} normal text
				.newRow(40).cell().cellSpan(1, 7, "note", css2ExcelWizard.getCss().richText("As an example using the grid shown, an employee rated {good|“good performer”} for performance and {str|“strong”} for behaviours will receive {emp|one month's fixed pay as their variable pay}.",
						css2ExcelWizard.getLayout().cssManager.newRichTexTLocalStyle().defineDefaultFont("10 black normal").defineFont("good", "12 orange bold").defineFont("str", "11 red bold").defineFont("emp", "11 green double"))).newRow().newRow().cell().cell("date", new Date());


		//Important: do not forget call appendMergedCells at the end if you have merged cell
		css2ExcelWizard.getLayout().appendMergedCells();
		//hide grid lines
		css2ExcelWizard.setDisplayGridLines(false);
		css2ExcelWizard.createSheet("Test");
		css2ExcelWizard.getLayout().newRow().cellSpan(4, 4, "heading", "Merged Cell").cell(null, 1, 2, 3, 4).newRow().cell(null, "New line start from here").newRow().cell().cellSpan(4, 2, "heading", "Merged Cell").newRow().cell(null, "New line start from here").newRow().newRow().newRow();

		css2ExcelWizard.getLayout().autoWidthFitsContent(4);
		css2ExcelWizard.getLayout().appendMergedCells();

		css2ExcelWizard.writeFile("./Streamlined Variable Pay.xlsx");
	}
}
