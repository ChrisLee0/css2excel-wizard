package org.chris.css2excelwizard;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class Css2ExcelWizard
{
	private LayoutManager layout;

	private CssManager css;

	private Workbook workbook;

	private Sheet sheet;

	public Css2ExcelWizard()
	{
		this(new XSSFWorkbook());
	}

	public Css2ExcelWizard(Workbook workbook)
	{
		this.workbook = workbook;
		this.css = new CssManager(workbook);
	}

	public void createSheet(String sheetName)
	{
		if(sheet!=null && layout!=null)
			layout.appendMergedCells();

		this.sheet = this.workbook.createSheet(sheetName);
		this.layout = new LayoutManager(css, sheet);
	}

	public Workbook getWorkbook()
	{
		return workbook;
	}

	public Sheet getSheet()
	{
		return sheet;
	}

	public LayoutManager getLayout()
	{
		return layout;
	}

	public CssManager getCss()
	{
		return css;
	}

	public void setDisplayGridLines(boolean display)
	{
		sheet.setDisplayGridlines(display);
	}

	public void ignoreError()
	{
		//not work in current version, only work in 2.X
		/*((XSSFSheet) sheet).addIgnoredErrors(new CellRangeAddress(1, layout.getCurrentRow(), 1, layout.getMaxCol()),
				IgnoredErrorType.TWO_DIGIT_TEXT_YEAR, IgnoredErrorType.NUMBER_STORED_AS_TEXT);*/
	}
	public void write(FileOutputStream outputStream) throws IOException
	{
		workbook.write(outputStream);
	}

	public void writeFile(File file) throws IOException
	{
		try (FileOutputStream fileOutputStream = new FileOutputStream(file))
		{
			write(fileOutputStream);
		}
	}

	public void writeFile(String path) throws IOException
	{
		writeFile(new File(path));
	}
}
