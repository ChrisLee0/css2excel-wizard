package org.chris.css2excelwizard;

import org.apache.poi.ss.usermodel.*;

public class Css2ExcelWizard
{
	public LayoutManager layout;

	public CssManager css;

	private Workbook workbook;

	private Sheet sheet;

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
}
