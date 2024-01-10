package org.chris.css2excelwizard;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class LayoutManager
{
	public void autoWidthFitsContent(int index)
	{
		sheet.autoSizeColumn(index);
	}

	public int getCurrentRow()
	{
		return row;
	}

	public int getMaxCol()
	{
		return maxCol;
	}

	private class MergedCell
	{
		int row;
		int rowSpan;
		int col;
		int colSpan;
		String styleId;

		public MergedCell(int row, int rowSpan, int col, int colSpan, String styleId)
		{
			this.row = row;
			this.rowSpan = rowSpan;
			this.col = col;
			this.colSpan = colSpan;
			this.styleId = styleId;
		}
	}

	private int row = 0;
	private int col = 0;

	private int maxCol = 1;
	private Sheet sheet;
	private Row rowObj;

	private float widthPadding = 1.0f;

	private SheetMap sheetMap;

	CssManager cssManager;

	private List<MergedCell> mergedCellList = new ArrayList<>();

	public LayoutManager(CssManager cssManager, Sheet sheet)
	{
		this(16, 16, cssManager, sheet);
	}

	public LayoutManager(int width, int height, CssManager cssManager, Sheet sheet)
	{
		this.cssManager = cssManager;
		this.sheet = sheet;
		this.sheetMap = new SheetMap(height, width);
	}

	public void setWidth(float... widths)
	{
		for (int i = 0; i < widths.length; i++)
		{
			sheet.setColumnWidth(i, ((int) ((widths[i] + widthPadding) * 256f)));
		}
	}

	public void setWidthPadding(float widthPadding)
	{
		this.widthPadding = widthPadding;
	}

	private void setContent(Cell cell, Object content)
	{
		if (content == null)
			return;

		if (content instanceof String)
			cell.setCellValue((String) content);
		else if (content instanceof Integer)
			cell.setCellValue((Integer) content);
		else if (content instanceof Long)
			cell.setCellValue((Long) content);
		else if (content instanceof Double)
			cell.setCellValue((Double) content);
		else if (content instanceof Float)
			cell.setCellValue((Float) content);
		else if (content instanceof Date)
			cell.setCellValue((Date) content);
		else if (content instanceof RichTextString)
			cell.setCellValue((RichTextString) content);
	}

	private void applyStyle(Cell cell, String styleId, Object content)
	{
		if (content != null)
			setContent(cell, content);

		CellStyle style = cssManager.getStyle(styleId);
		if (style != null)
			cell.setCellStyle(style);
	}

	private void applyStyle(Cell cell, CellStyle style, Object content)
	{
		if (content != null)
			setContent(cell, content);

		if (style != null)
			cell.setCellStyle(style);
	}

	LayoutManager newRow()
	{
		rowObj = sheet.createRow(row++);
		col = findFreeCol(row - 1);
		return this;
	}

	LayoutManager newRow(int height)
	{
		newRow();
		rowObj.setHeight((short) (height * 20));
		return this;
	}

	private int findFreeCol(int row)
	{
		for (int i = 0; i < 65536; i++)
			if (!sheetMap.get(row, i))
				return i;
		return -1;
	}

	LayoutManager cell()
	{
		col++;
		return this;
	}

	public LayoutManager cell(String styleId, Object... contents)
	{
		CellStyle style = styleId == null ? cssManager.rootStyle : cssManager.getStyle(styleId);

		if (contents == null || contents.length == 0)
			applyStyle(rowObj.createCell(col++), style, null);
		else
			for (Object content : contents)
				applyStyle(rowObj.createCell(col++), style, content);
		return this;
	}

	LayoutManager cellSpan(int row, int col)
	{
		return cellSpan(row, col, null, null);
	}

	LayoutManager cellRowSpan(int row)
	{
		if (row == 1)
			return this;

		for (int i = this.row; i < this.row + row; i++)
			sheetMap.set(i, col, true);

		this.col++;

		return this;
	}

	LayoutManager cellColSpan(int col)
	{
		for (int j = this.col; j < this.col + col; j++)
			sheetMap.set(row, j, true);

		this.col += col;
		return this;
	}

	LayoutManager cellSpan(int row, int col, String styleId, Object content)
	{
		if (row == 1 && col == 1)
			return cell(styleId, content);

		for (int i = this.row - 1; i < this.row - 1 + row; i++)
			for (int j = this.col; j < this.col + col; j++)
				sheetMap.set(i, j, true);


		mergedCellList.add(new MergedCell(this.row - 1, row, this.col, col, styleId));

		if (styleId != null || content != null)
			applyStyle(rowObj.createCell(this.col), styleId, content);

		this.col += col;

		return this;
	}

	LayoutManager cellRowSpan(int row, String styleId, Object content)
	{
		return cellSpan(row, 1, styleId, content);
	}

	LayoutManager cellColSpan(int col, String styleId, Object content)
	{
		return cellSpan(1, col, styleId, content);
	}

	LayoutManager cell(String styleId, Object content)
	{
		if (styleId != null || content != null)
			applyStyle(rowObj.createCell(col), styleId, content);
		col++;
		return this;
	}

	LayoutManager skipRow(int row)
	{
		this.row += row + 1;
		this.col = 0;
		return this;
	}

	LayoutManager skipCell(int col)
	{
		this.col += col + 1;
		return this;
	}

	public void appendMergedCells()
	{
		for (MergedCell mergedCell : mergedCellList)
		{
			CellRangeAddress range = new CellRangeAddress(mergedCell.row, mergedCell.row + mergedCell.rowSpan - 1, mergedCell.col, mergedCell.col + mergedCell.colSpan - 1);
			this.cssManager.applyBorderToMergedCells(sheet, range, mergedCell.styleId);
			sheet.addMergedRegion(range);
		}

		mergedCellList.clear();
	}
}
