package org.chris.css2excelwizard;

import java.util.BitSet;

public class SheetMap
{
	private BitSet data;
	private int rows;
	private int cols;

	public SheetMap()
	{
		this(16, 16);
	}

	public SheetMap(int initialRows, int initialCols)
	{
		rows = initialRows;
		cols = initialCols;
		data = new BitSet(rows * cols);
	}

	public boolean get(int row, int col)
	{
		if (row >= rows || col >= cols)
		{
			return false;
		}
		int index = row * cols + col;
		return data.get(index);
	}

	public void set(int row, int col, boolean value)
	{
		if (row >= rows || col >= cols)
		{
			expand(row, col);
		}
		int index = row * cols + col;
		data.set(index, value);
	}

	private void expand(int targetRow, int targetCol)
	{
		int newRows = Math.max(targetRow + 1, (int) (rows * 1.5));
		int newCols = Math.max(targetCol + 1, (int) (cols * 1.2));
		BitSet newData = new BitSet(newRows * newCols);

		for (int row = 0; row < rows; row++)
		{
			for (int col = 0; col < cols; col++)
			{
				int newIndex = row * newCols + col;
				int oldIndex = row * cols + col;
				newData.set(newIndex, data.get(oldIndex));
			}
		}

		rows = newRows;
		cols = newCols;
		data = newData;
	}

	public void print()
	{
		System.out.println(rows + "X" + cols);
		for (int row = 0; row < rows; row++)
		{
			for (int col = 0; col < cols; col++)
			{
				System.out.print(get(row, col) ? "1" : "0");
			}
			System.out.println();
		}
	}
}
