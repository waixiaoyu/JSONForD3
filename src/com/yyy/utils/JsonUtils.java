package com.yyy.utils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;

import org.json.JSONArray;
import org.json.JSONObject;

public class JsonUtils {
	public static void main(String[] args) {
		JSONObject obj6 = new JSONObject();
		obj6.put("title", "book1").put("price", "$11");
		JSONObject obj3 = new JSONObject();

		obj3.put("author", new JSONObject().put("name", "author-1"));

		JSONArray a1 = new JSONArray();
		JSONObject a2 = new JSONObject();
		a2.put("name", "进口");
		a2.put("children", "");
		JSONObject a3 = new JSONObject();
		a3.put("name", "出口");
		a3.put("children", "");
		a1.put(a2);
		a1.put(a3);
		System.out.println(a1.toString());
	}


	public static void write(String jsonString, String filePath) {
		try {
			FileOutputStream fos = null;
			PrintWriter writer = null;
			fos = new FileOutputStream(filePath);
			writer = new PrintWriter(new OutputStreamWriter(fos, "utf-8"));
			writer.write(jsonString);
			writer.close();
			fos.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
