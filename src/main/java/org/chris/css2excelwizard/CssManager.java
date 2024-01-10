package org.chris.css2excelwizard;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.awt.Color;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class CssManager
{
	Workbook workbook;

	Map<String, CellStyle> styleMap = new HashMap<>();
	Map<String, CellStyle> parentStyleMap = new HashMap<>();

	Map<String, Map<String, String>> styleRuleMap = new HashMap<>();

	static Map<String, Color> builtInColor = new HashMap<>();
	Map<String, IndexedColors> indexedColorsMap = new HashMap<>();

	XSSFCellStyle rootStyle;
	XSSFFont rootFont;


	public CssManager(Workbook workbook)
	{
		this.workbook = workbook;
		workbook.getCreationHelper();
		initBuildIn();
		this.rootFont = (XSSFFont) workbook.createFont();
		this.rootStyle = (XSSFCellStyle) workbook.createCellStyle();
		rootStyle.setFont(rootFont);
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

	private XSSFFont cloneFont(XSSFFont font)
	{
		XSSFFont newFont = (XSSFFont) workbook.createFont();

		newFont.setFontName(font.getFontName());
		newFont.setFontHeightInPoints(font.getFontHeightInPoints());
		//newFont.setFamily(font.getFamily());
		newFont.setColor(font.getXSSFColor());
		newFont.setBold(font.getBold());
		newFont.setItalic(font.getItalic());
		newFont.setUnderline(font.getUnderline());

		return newFont;
	}

	public void defineRootStyle(String css)
	{
		Map<String, String> cssRule = parseCssRule(css);
		XSSFFont font = (XSSFFont) workbook.createFont();
		XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
		setStyle(style, font, cssRule);
		this.rootStyle = style;
		this.rootFont = font;
	}

	public CssManager defineStyle(String styleId, String css)
	{
		if (styleMap.containsKey(styleId))
			throw new IllegalArgumentException(String.format("Style '%s' already defined", styleId));

		Map<String, String> computedRules = parseCssRule(css);
		styleRuleMap.put(styleId, computedRules);

		XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
		style.cloneStyleFrom(rootStyle);
		XSSFFont font = cloneFont(rootFont);
		setStyle(style, font, computedRules);

		styleMap.put(styleId, style);

		return this;
	}

	public CssManager extendStyle(String parentStyleId, String styleId, String css)
	{
		if (!styleMap.containsKey(parentStyleId))
			throw new IllegalArgumentException(String.format("Style '%s' does not define", parentStyleId));

		if (styleMap.containsKey(styleId))
			throw new IllegalArgumentException(String.format("Style '%s' already defined", styleId));

		Map<String, String> computedRules = parseCssRule(css);
		styleRuleMap.put(styleId, computedRules);

		XSSFCellStyle parentStyle = (XSSFCellStyle) styleMap.get(parentStyleId);
		XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
		style.cloneStyleFrom(parentStyle);
		XSSFFont font = cloneFont(parentStyle.getFont());
		setStyle(style, font, computedRules);

		styleMap.put(styleId, style);
		parentStyleMap.put(styleId, parentStyle);

		return this;
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

	private static Color parseColor(String colorRule)
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


	private static XSSFColor parseXssColor(String colorRule)
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


	private CellStyle setStyle(XSSFCellStyle style, XSSFFont font, Map<String, String> rules)
	{
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
				case "wrap-text":
					switch (ruleValue)
					{
						case "true":
							style.setWrapText(true);
							break;
						case "false":
							style.setWrapText(false);
							break;
						default:
							break;
					}
					break;
				case "border":
					if ("none".equals(ruleValue))
					{
						style.setBorderTop(BorderStyle.NONE);
						style.setBorderRight(BorderStyle.NONE);
						style.setBorderBottom(BorderStyle.NONE);
						style.setBorderLeft(BorderStyle.NONE);
					}
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

		postProcessRule(rules, style);

		return style;
	}

	private void postProcessRule(Map<String, String> rules, XSSFCellStyle style)
	{
		if (rules.containsKey("border-color") && !rules.containsKey("border-style"))
		{
			if (style.getBorderTop() == BorderStyle.NONE.ordinal())
				style.setBorderTop(BorderStyle.THIN);

			if (style.getBorderRight() == BorderStyle.NONE.ordinal())
				style.setBorderRight(BorderStyle.THIN);

			if (style.getBorderBottom() == BorderStyle.NONE.ordinal())
				style.setBorderBottom(BorderStyle.THIN);

			if (style.getBorderLeft() == BorderStyle.NONE.ordinal())
				style.setBorderLeft(BorderStyle.THIN);
		}
	}

	private static void parseFont(XSSFFont font, String ruleValue)
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
					if (color != null && color != Color.BLACK)
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
		if (styleId == null)
			return;

		Map<String, String> ruleMap = styleRuleMap.get(styleId);
		if (ruleMap == null)
			return;

		String border = ruleMap.get("border");

		if ("none".equals(border))
			return;

		String borderColor = ruleMap.get("border-color");
		String borderStyle = ruleMap.get("border-style");

		if (borderColor == null)
			borderColor = "black";

		IndexedColors indexedColors = indexedColorsMap.get(borderColor);
		if (indexedColors != null)
		{
			RegionUtil.setTopBorderColor(indexedColors.index, range, sheet, workbook);
			RegionUtil.setRightBorderColor(indexedColors.index, range, sheet, workbook);
			RegionUtil.setBottomBorderColor(indexedColors.index, range, sheet, workbook);
			RegionUtil.setLeftBorderColor(indexedColors.index, range, sheet, workbook);
		}

		short borderStyleTop;
		short borderStyleRight;
		short borderStyleBottom;
		short borderStyleLeft;

		if (borderStyle == null)
		{
			CellStyle parentStyle = parentStyleMap.get(styleId);
			if (parentStyle == null)
			{
				borderStyleTop = (short) BorderStyle.THIN.ordinal();
				borderStyleRight = (short) BorderStyle.THIN.ordinal();
				borderStyleBottom = (short) BorderStyle.THIN.ordinal();
				borderStyleLeft = (short) BorderStyle.THIN.ordinal();
			}
			else
			{
				borderStyleTop = parentStyle.getBorderTop();
				borderStyleRight = parentStyle.getBorderRight();
				borderStyleBottom = parentStyle.getBorderBottom();
				borderStyleLeft = parentStyle.getBorderLeft();
			}
		}
		else
		{
			short borderStyleEnum = (short) parseBorderStyle(borderStyle);
			borderStyleTop = borderStyleEnum;
			borderStyleRight = borderStyleEnum;
			borderStyleBottom = borderStyleEnum;
			borderStyleLeft = borderStyleEnum;
		}


		RegionUtil.setBorderTop(borderStyleTop, range, sheet, workbook);
		RegionUtil.setBorderRight(borderStyleRight, range, sheet, workbook);
		RegionUtil.setBorderBottom(borderStyleBottom, range, sheet, workbook);
		RegionUtil.setBorderLeft(borderStyleLeft, range, sheet, workbook);
	}

	public RichTexTLocalStyle newRichTexTLocalStyle()
	{
		return new RichTexTLocalStyle();
	}

	public class RichTexTLocalStyle
	{
		private Font defaultFont;
		Map<String, Font> map = new HashMap<>();

		public RichTexTLocalStyle defineDefaultFont(String css)
		{
			XSSFFont font = (XSSFFont) workbook.createFont();
			parseFont(font, css);
			defaultFont = font;
			return this;
		}

		public RichTexTLocalStyle defineFont(String fontId, String css)
		{
			XSSFFont font = (XSSFFont) workbook.createFont();
			parseFont(font, css);
			if (map.containsKey(fontId))
				throw new IllegalArgumentException(String.format("Font ID '%s' already defined", fontId));
			map.put(fontId, font);
			return this;
		}

		public Font getDefaultFont()
		{
			return defaultFont;
		}

		public Font get(String fontId)
		{
			Font font = map.get(fontId);
			if (font == null) throw new IllegalStateException("Not found font called '" + fontId + "'");
			return font;
		}
	}

	private static class RichTexFragment
	{
		public final String id;
		public final int start;
		public final int end;

		public RichTexFragment(String id, int start, int end)
		{
			this.id = id;
			this.start = start;
			this.end = end;
		}
	}

	public RichTextString richText(String text, RichTexTLocalStyle localStyle)
	{
		StringBuilder rawText = new StringBuilder();
		List<RichTexFragment> position = new ArrayList<>(); // 0:normal  1: in {brace}
		boolean inBrace = false;
		int braceStart = -1;
		String fontId = null;
		for (int i = 0; i < text.length(); i++)
		{
			char ch = text.charAt(i);
			switch (ch)
			{
				case '{':
					if (inBrace) throw new IllegalStateException("Incorrect Grammar: Nested {brace} not allowed");
					inBrace = true;
					int vbar = text.indexOf('|', i + 1);
					if (vbar == -1 || i + 1 == vbar)
						throw new IllegalStateException("Incorrect Grammar: No font id followed by between { and |");
					fontId = text.substring(i + 1, vbar);
					braceStart = rawText.length();
					i = vbar;
					break;
				case '}':

					if (!inBrace)
						throw new IllegalStateException("Incorrect Grammar: Unable to end an not started {brace}");
					inBrace = false;
					position.add(new RichTexFragment(fontId, braceStart, rawText.length()));
					break;
				case '\\':
					if (i >= text.length() - 1)
						throw new IllegalStateException("Incorrect Grammar: No character followed by \\");
					char escape = text.charAt(i + 1);
					switch (escape)
					{
						case '{':
						case '}':
						case '\\':
							rawText.append(escape);
							i++;
							break;
						default:
							throw new IllegalStateException("Incorrect Grammar: Unsupported escape character " + escape);
					}
					break;
				default:
					rawText.append(ch);
					break;
			}
		}

		XSSFRichTextString richTextString = new XSSFRichTextString(rawText.toString());
		if (position.isEmpty()) richTextString.applyFont(0, rawText.length(), localStyle.getDefaultFont());
		else
		{
			int startCursor = 0;
			for (RichTexFragment richTexFragment : position)
			{
				if (startCursor < richTexFragment.start && localStyle.getDefaultFont() != null)
					richTextString.applyFont(startCursor, richTexFragment.start, localStyle.getDefaultFont());
				richTextString.applyFont(richTexFragment.start, richTexFragment.end, localStyle.get(richTexFragment.id));
				startCursor = richTexFragment.end;
			}
			if (startCursor < rawText.length() - 1 && localStyle.getDefaultFont() != null)
				richTextString.applyFont(startCursor, rawText.length(), localStyle.getDefaultFont());
		}
		return richTextString;
	}
}
