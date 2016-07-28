package com.yyy.transfer;

import java.awt.Color;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Random;

import org.apache.poi.ss.usermodel.Row;
import org.json.JSONArray;
import org.json.JSONObject;

import com.yyy.utils.ExcelUtils;
import com.yyy.utils.JsonUtils;

public class ExcelToJson {
	private static final String FILE_NAME = "test.xlsx";
	private static final String OUTPUT_PATH = "D:\\tomcat\\webapps\\demo\\wheel1.json";
	private static final String HYPHEN = "-";
	private List<String> lStrings = new ArrayList<>();

	public static void main(String[] args) throws IOException {
		ExcelToJson pe = new ExcelToJson();
		pe.readXls();
		// pe.createTable(TABLE_NAME, strHeaders);
	}

	public void readXls() throws IOException {
		List<Row> lRows = ExcelUtils.getRows(FILE_NAME, 0);
		for (Row row : lRows) {
			StringBuilder sb = new StringBuilder();
			for (int i = 0; i < row.getLastCellNum(); i++) {
				sb.append(row.getCell(i));
				sb.append("-");
			}
			lStrings.add(sb.toString());
		}
		sortStrings();
	}

	public void sortStrings() {
		Collections.sort(lStrings);
		for (String string : lStrings) {
		}
		createJSON();
	}

	public void createJSON() {
		int nColumnLength = lStrings.get(0).split(HYPHEN).length;
		JSONArray jArray = new JSONArray();
		JSONObject jObj = new JSONObject();

		String strNewName = "";
		String[] strArray;
		int nColumnCurrentIndex = 0;
		int nLastRowIndex = 0;
		boolean bIsStart = false;
		for (int i = 0; i < lStrings.size(); i++) {
			strArray = lStrings.get(i).split(HYPHEN);
			if (!strArray[nColumnCurrentIndex].equals(strNewName)) {
				if (bIsStart) {
					jObj.put("name", strNewName);
					jObj.put("children", getChildren(nLastRowIndex, i, nColumnCurrentIndex + 1));
					jArray.put(jObj);
				}
				bIsStart = true;
				strNewName = strArray[nColumnCurrentIndex];
				nLastRowIndex = i;
				jObj = new JSONObject();
			}
		}
		jObj.put("name", strNewName);
		jObj.put("children", getChildren(nLastRowIndex, lStrings.size(), nColumnCurrentIndex + 1));
		jArray.put(jObj);

		System.out.println(jArray.toString());
		JsonUtils.write(jArray.toString(), OUTPUT_PATH);
	}

	public JSONArray getChildren(int nStartRowIndex, int nEndRowIndex, int nColumnCurrentIndex) {
		int nColumnLength = lStrings.get(0).split(HYPHEN).length;

		JSONArray jArray = new JSONArray();
		JSONObject jObj = new JSONObject();

		String strNewName = "";
		String[] strArray;
		int nLastRowIndex = 0;
		boolean bIsStart = false;
		for (int i = nStartRowIndex; i < nEndRowIndex; i++) {
			strArray = lStrings.get(i).split(HYPHEN);
			if (!strArray[nColumnCurrentIndex].equals(strNewName)) {
				if (bIsStart) {
					jObj.put("name", strNewName);
					if (nColumnCurrentIndex < nColumnLength - 1) {
						jObj.put("children", getChildren(nLastRowIndex, i, nColumnCurrentIndex + 1));
					} else {
						jObj.put("colour", getRandColorCode());

					}
					jArray.put(jObj);
				}
				bIsStart = true;
				strNewName = strArray[nColumnCurrentIndex];
				nLastRowIndex = i;
				jObj = new JSONObject();
			}
		}
		jObj.put("name", strNewName);
		if (nColumnCurrentIndex < nColumnLength - 1) {
			jObj.put("children", getChildren(nLastRowIndex, nEndRowIndex, nColumnCurrentIndex + 1));
		} else {
			jObj.put("colour", getRandColorCode());

		}
		jArray.put(jObj);
		return jArray;
	}

	/**
	 * 获取十六进制的随机颜色代码.例如 "#6E36B4" , For HTML ,
	 * 
	 * @return String
	 */
	public String getRandColorCode() {
		String r, g, b;
		Random random = new Random();
		r = Integer.toHexString(random.nextInt(256)).toUpperCase();
		g = Integer.toHexString(random.nextInt(256)).toUpperCase();
		b = Integer.toHexString(random.nextInt(256)).toUpperCase();

		r = r.length() == 1 ? "0" + r : r;
		g = g.length() == 1 ? "0" + g : g;
		b = b.length() == 1 ? "0" + b : b;

		return "#" + r + g + b;
	}

	/**
	 * 获取十六进制的渐变颜色代码.例如 "#6E36B4" , For HTML ,
	 * 
	 * @return
	 */
	private int nColorIndex = 0;

	public String getGradualChangeColorCode() {
		int oldR = 50, oldG = 100, oldB = 250;
		int newR = 250, newG = 100, newB = 50;
		Color oldColor = new Color(oldR, oldG, oldB); // 初始颜色
		Color newColor = new Color(newR, newG, newB); // 结束颜色

		int step = lStrings.size() - 1; // 分多少步完成
		int or = oldColor.getRed() + (newColor.getRed() - oldColor.getRed()) * nColorIndex / step;
		int og = oldColor.getGreen() + (newColor.getGreen() - oldColor.getGreen()) * nColorIndex / step;
		int ob = oldColor.getBlue() + (newColor.getBlue() - oldColor.getBlue()) * nColorIndex / step;
		++nColorIndex;
		String r, g, b;
		r = Integer.toHexString(or).toUpperCase();
		g = Integer.toHexString(og).toUpperCase();
		b = Integer.toHexString(ob).toUpperCase();
		r = r.length() == 1 ? "0" + r : r;
		g = g.length() == 1 ? "0" + g : g;
		b = b.length() == 1 ? "0" + b : b;
		return "#" + r + g + b;
	}
}
