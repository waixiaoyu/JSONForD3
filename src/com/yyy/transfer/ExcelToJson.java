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

	/**
	 * ���ַ����������򣬷�����Ƶ�ʱ�����
	 */
	public void sortStrings() {
		Collections.sort(lStrings);
		JsonUtils.write(getChildren(0, lStrings.size(), 0).toString(), OUTPUT_PATH);
	}
	
	/**
	 * �ݹ���ã�Ƕ������json
	 * @param nStartRowIndex ��ʼλ��
	 * @param nEndRowIndex ����λ��
	 * @param nColumnCurrentIndex ��ǰ�ڼ���column
	 * @return JSONArray ֱ�ӿ��Ա�d3��ȡ��json�ַ���
	 */
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
						jObj.put("colour", getGradualChangeColorCode());

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
			jObj.put("colour", getGradualChangeColorCode());

		}
		jArray.put(jObj);
		return jArray;
	}

	/**
	 * ��ȡʮ�����Ƶ������ɫ����.���� "#6E36B4" , For HTML ,
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
	 * ��ȡʮ�����ƵĽ�����ɫ����.���� "#6E36B4" , For HTML ,
	 * 
	 * @return String
	 */
	private int nColorIndex = 0;

	public String getGradualChangeColorCode() {
		Color oldColor = new Color(255, 10, 10); // ��ʼ��ɫ
		Color newColor = new Color(253, 245, 230); // ������ɫ

		int step = lStrings.size() - 1; // �ֶ��ٲ����
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
