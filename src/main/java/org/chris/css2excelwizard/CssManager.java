package org.chris.css2excelwizard;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.awt.Color;
import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;

public class CssManager
{
	Workbook workbook;

	Map<String, CellStyle> styleMap = new HashMap<>();

	Map<String, Map<String, String>> styleRuleMap = new HashMap<>();

	Map<String, String> rootStyle = new HashMap<>();
	Map<String, Color> builtInColor = new HashMap<>();
	Map<String, IndexedColors> indexedColorsMap = new HashMap<>();


	public CssManager(Workbook workbook)
	{
		this.workbook = workbook;
		workbook.getCreationHelper();
		initBuildIn();
	}

	private void initBuildIn()
	{
		try
		{
			Field[] declaredFields = Color.class.getDeclaredFields();
			for (Field field : declaredFields)
				if (field.getType() == Color.class)
					builtInColor.put(field.getName().toLowerCase().replace("_", ""), (Color) field.get(null));
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}

		for (IndexedColors values : IndexedColors.values())
			indexedColorsMap.put(values.name().toLowerCase().replace("_", ""), values);
	}


	public void defineRootStyle(String css)
	{
		Map<String, String> cssRule = parseCssRule(css);
		rootStyle.putAll(cssRule);
	}

	public void defineStyle(String styleId, String css)
	{
		if (styleMap.containsKey(styleId))
			throw new IllegalArgumentException(String.format("Style '%s' already defined", styleId));

		Map<String, String> rules = new HashMap<>(rootStyle);

		Map<String, String> computedRules = parseCssRule(css);
		rules.putAll(computedRules);

		styleRuleMap.put(styleId, rules);

		styleMap.put(styleId, createStyle(rules));
	}

	public void extendStyle(String parentStyleId, String styleId, String css)
	{
		if (!styleRuleMap.containsKey(parentStyleId))
			throw new IllegalArgumentException(String.format("Style '%s' does not define", parentStyleId));

		if (styleMap.containsKey(styleId))
			throw new IllegalArgumentException(String.format("Style '%s' already defined", styleId));

		Map<String, String> rules = new HashMap<>(styleRuleMap.get(parentStyleId));

		Map<String, String> computedRules = parseCssRule(css);
		rules.putAll(computedRules);

		styleRuleMap.put(styleId, rules);

		styleMap.put(styleId, createStyle(rules));
	}

	public CellStyle getStyle(String id)
	{
		return styleMap.get(id);
	}

	private Map<String, String> parseCssRule(String css)
	{
		Map<String, String> map = new HashMap<>();

		String[] cssRules = css.split(";");
		for (String rule : cssRules)
		{
			int pos = rule.indexOf(':');

			if (pos == -1)
				continue;

			String ruleName = rule.substring(0, pos).trim();
			String ruleValue = rule.substring(pos + 1).trim();

			map.put(ruleName, ruleValue);
		}

		return map;
	}

	private Color parseColor(String colorRule)
	{
		Color color = builtInColor.get(colorRule);
		if (color != null)
			return color;

		colorRule = colorRule.toLowerCase().replace(" ", "");

		if (colorRule.startsWith("rgb("))
		{
			String content = colorRule.replace("rgb(", "").replace(")", "");
			String[] rgb = content.split(",");
			if (rgb.length != 3)
				return builtInColor.get("black");

			return new Color(Integer.parseInt(rgb[0]), Integer.parseInt(rgb[1]), Integer.parseInt(rgb[2]));
		}
		else
			return builtInColor.get("black");
	}

	private void parseAlign(XSSFCellStyle style, String ruleValue, String type)
	{
		if ("text-align".equals(type) || "align".equals(type))
		{
			switch (ruleValue)
			{
				case "left":
					style.setAlignment(HorizontalAlignment.LEFT);
					break;
				case "center":
					style.setAlignment(HorizontalAlignment.CENTER);
					break;
				case "right":
					style.setAlignment(HorizontalAlignment.RIGHT);
					break;
				case "fill":
					style.setAlignment(HorizontalAlignment.FILL);
					break;
				case "justify":
					style.setAlignment(HorizontalAlignment.JUSTIFY);
					break;
				case "center-selection":
					style.setAlignment(HorizontalAlignment.CENTER_SELECTION);
					break;
				case "distributed":
					style.setAlignment(HorizontalAlignment.DISTRIBUTED);
					break;
				case "general":
					style.setAlignment(HorizontalAlignment.GENERAL);
					break;
			}
		}

		if ("vertical-align".equals(type) || "align".equals(type))
		{
			switch (ruleValue)
			{
				case "top":
					style.setVerticalAlignment(VerticalAlignment.TOP);
					break;
				case "center":
				case "middle":
					style.setVerticalAlignment(VerticalAlignment.CENTER);
					break;
				case "bottom":
					style.setVerticalAlignment(VerticalAlignment.BOTTOM);
					break;
				case "justify":
					style.setVerticalAlignment(VerticalAlignment.JUSTIFY);
					break;
				case "distributed":
					style.setVerticalAlignment(VerticalAlignment.DISTRIBUTED);
					break;
			}
		}
	}


	private XSSFColor parseXssColor(String colorRule)
	{
		Color color = builtInColor.get(colorRule);
		if (color == null)
			color = parseColor(colorRule);
		return color == null ? null : new XSSFColor(color);
	}

	private int parseBorderStyle(String borderStyle)
	{
		switch (borderStyle)
		{
			case "none":
				return CellStyle.BORDER_NONE;
			case "thin":
				return CellStyle.BORDER_THIN;
			case "medium":
				return CellStyle.BORDER_MEDIUM;
			case "dashed":
				return CellStyle.BORDER_DASHED;
			case "hair":
				return CellStyle.BORDER_HAIR;
			case "thick":
				return CellStyle.BORDER_THICK;
			case "double":
				return CellStyle.BORDER_DOUBLE;
			case "dotted":
				return CellStyle.BORDER_DOTTED;
			case "medium-dashed":
				return CellStyle.BORDER_MEDIUM_DASHED;
			case "dash-dot":
				return CellStyle.BORDER_DASH_DOT;
			case "medium-dash-dot":
				return CellStyle.BORDER_MEDIUM_DASH_DOT;
			case "dash-dot-dot":
				return CellStyle.BORDER_DASH_DOT_DOT;
			case "medium-dash-dot-dot":
				return CellStyle.BORDER_MEDIUM_DASH_DOT_DOT;
			case "slanted-dash-dot":
				return CellStyle.BORDER_SLANTED_DASH_DOT;
			default:
				return -1;
		}
	}

	private BorderStyle parseBorderStyleEnum(String borderStyle)
	{
		switch (borderStyle)
		{
			case "none":
				return BorderStyle.NONE;
			case "thin":
				return BorderStyle.THIN;
			case "medium":
				return BorderStyle.MEDIUM;
			case "dashed":
				return BorderStyle.DASHED;
			case "dotted":
				return BorderStyle.DOTTED;
			case "thick":
				return BorderStyle.THICK;
			case "double":
				return BorderStyle.DOUBLE;
			case "hair":
				return BorderStyle.HAIR;
			case "medium-dashed":
				return BorderStyle.MEDIUM_DASHED;
			case "dash-dot":
				return BorderStyle.DASH_DOT;
			case "medium-dash-dot":
				return BorderStyle.MEDIUM_DASH_DOT;
			case "dash-dot-dot":
				return BorderStyle.DASH_DOT_DOT;
			case "medium-dash-dot-dotc":
				return BorderStyle.MEDIUM_DASH_DOT_DOTC;
			case "slanted-dash-dot":
				return BorderStyle.SLANTED_DASH_DOT;
			default:
				return null;
		}
	}


	private CellStyle createStyle(Map<String, String> rules)
	{
		XSSFFont font = (XSSFFont) workbook.createFont();
		XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
		String[] ruleVals;

		for (String rule : rules.keySet())
		{
			String ruleValue = rules.get(rule);
			switch (rule)
			{
				case "background-color":
					style.setFillForegroundColor(parseXssColor(ruleValue));
					style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					break;
				case "color":
					font.setColor(parseXssColor(ruleValue));
					break;
				case "font":
					parseFont(font, ruleValue);
					break;
				case "font-size":
					font.setFontHeightInPoints(Short.parseShort(ruleValue));
					break;
				case "font-weight":
					if (ruleValue.equals("bold"))
						font.setBold(true);
					else if (ruleValue.equals("normal"))
						font.setBold(false);
					break;
				case "font-style":
					ruleVals = ruleValue.split(" ");
					if (ruleVals.length == 1 && ruleVals[0].equals("none"))
					{
						font.setUnderline(FontUnderline.NONE);
						font.setItalic(false);
					}
					else
					{
						for (String val : ruleVals)
						{
							switch (val)
							{
								case "italic":
									font.setItalic(true);
									break;
								case "underline":
								case "underline-single":
									font.setUnderline(FontUnderline.SINGLE);
									break;
								case "underline-double":
									font.setUnderline(FontUnderline.DOUBLE);
									break;
								case "underline-single-accounting":
									font.setUnderline(FontUnderline.SINGLE_ACCOUNTING);
									break;
								case "underline-double-accounting":
									font.setUnderline(FontUnderline.DOUBLE_ACCOUNTING);
									break;
							}
						}
					}

					break;
				case "font-family":
					if (!ruleValue.equals("auto"))
						font.setFontName(ruleValue);
					break;
				case "indention":
					if ("none".equals(ruleValue))
						style.setIndention((short) 0);
					else
						style.setIndention(Short.parseShort(ruleValue));
					break;
				case "editable":
					if ("lock".equals(ruleValue))
						style.setLocked(true);
					else if ("free".equals(ruleValue))
						style.setLocked(false);
					break;
				case "align":
					ruleVals = ruleValue.split(" ");
					if (ruleVals.length == 1)
						parseAlign(style, ruleValue, rule);
					else
					{
						parseAlign(style, ruleVals[0], "text-align");
						parseAlign(style, ruleVals[1], "vertical-align");
					}
					break;
				case "text-align":
				case "vertical-align":
					parseAlign(style, ruleValue, rule);
					break;
				case "border-color":
					XSSFColor color = parseXssColor(ruleValue);
					if (color != null)
					{
						style.setTopBorderColor(color);
						style.setRightBorderColor(color);
						style.setBottomBorderColor(color);
						style.setLeftBorderColor(color);
					}
					break;
				case "border-style":
					BorderStyle borderStyle = parseBorderStyleEnum(rule);
					if (borderStyle != null)
					{
						style.setBorderTop(borderStyle);
						style.setBorderRight(borderStyle);
						style.setBorderBottom(borderStyle);
						style.setBorderLeft(borderStyle);
					}
					break;
				case "date-format":
					CreationHelper createHelper = workbook.getCreationHelper();
					short dateFormat = createHelper.createDataFormat().getFormat(ruleValue);
					style.setDataFormat(dateFormat);
					break;
				case "number-format":
					createHelper = workbook.getCreationHelper();
					dateFormat = createHelper.createDataFormat().getFormat(ruleValue);
					style.setDataFormat(dateFormat);
			}
		}

		style.setFont(font);

		return style;
	}

	private void parseFont(XSSFFont font, String ruleValue)
	{
		String[] values = ruleValue.split(" ");
		for (String val : values)
		{
			val = val.trim();
			if (val.length() == 0)
				continue;

			switch (val)
			{
				case "normal":
					break;
				case "bold":
					font.setBold(true);
					break;
				case "italic":
					font.setItalic(true);
					break;
				case "single":
					font.setUnderline(FontUnderline.SINGLE);
					break;
				case "double":
					font.setUnderline(FontUnderline.DOUBLE);
					break;
				default:
					val = val.replace(" ", "");

					boolean isDigit = true;
					for (int i = 0; i < val.length(); i++)
					{
						if (val.charAt(i) < '0' || val.charAt(i) > '9')
						{
							isDigit = false;
							break;
						}
					}

					if (isDigit)
					{
						font.setFontHeightInPoints(Short.parseShort(val));
						break;
					}


					if (val.startsWith("#") || val.toLowerCase().startsWith("rgb("))
					{
						font.setColor(parseXssColor(val));
						break;
					}

					Color color = builtInColor.get(val);
					if (color != null)
					{
						font.setColor(new XSSFColor(color));
						break;
					}

					font.setFontName(val);
			}
		}
	}

	public void applyBorderToMergedCells(Sheet sheet, CellRangeAddress range, String styleId)
	{
		Map<String, String> ruleMap = styleRuleMap.get(styleId);
		String borderColor = ruleMap.get("border-color");
		String borderStyle = ruleMap.get("border-style");

		if (borderColor != null)
		{
			IndexedColors indexedColors = indexedColorsMap.get(borderColor);
			if (indexedColors != null)
			{
				RegionUtil.setTopBorderColor(indexedColors.index, range, sheet, workbook);
				RegionUtil.setRightBorderColor(indexedColors.index, range, sheet, workbook);
				RegionUtil.setBottomBorderColor(indexedColors.index, range, sheet, workbook);
				RegionUtil.setLeftBorderColor(indexedColors.index, range, sheet, workbook);
			}
		}

		if (borderStyle != null)
		{
			int borderStyleVal = parseBorderStyle(borderStyle);
			if (borderStyleVal != -1)
			{
				RegionUtil.setBorderTop(borderStyleVal, range, sheet, workbook);
				RegionUtil.setBorderRight(borderStyleVal, range, sheet, workbook);
				RegionUtil.setBorderBottom(borderStyleVal, range, sheet, workbook);
				RegionUtil.setBorderLeft(borderStyleVal, range, sheet, workbook);
			}
		}
	}

	public RichTextString richText(String text, String defaultFontStyle, String... restFontStyles)
	{

		return null;
	}
}
