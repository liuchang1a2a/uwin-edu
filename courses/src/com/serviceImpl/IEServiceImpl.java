package com.serviceImpl;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.HashSet;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.service.IEService;


@Service
public class IEServiceImpl implements IEService {
	
	public JSONObject getXlsInfo(MultipartFile file){  
		JSONObject result = new JSONObject();
		try{
//			inputStream = new FileInputStream(new File(fileAbsolutePath));
			InputStream is = file.getInputStream();
	        XSSFWorkbook book = new XSSFWorkbook(is);
	        book.getNumberOfSheets();
	        
		}catch(Exception e){
			result.put("success", false);
			result.put("error", "上传文件读取失败,请重试");
		}
		return result;
	}
	
	
	public JSONObject getXlsxInfo(MultipartFile file){
		JSONObject result = new JSONObject();
		try{
			// sheets:{
//			sheetName:{info:{"students":[],"teachers":[],courTypes:[],dates:[]}},
//			"list":[{0:"",1:"",2:""},...]
//			}
			InputStream is = file.getInputStream();
	        XSSFWorkbook book = new XSSFWorkbook(is);
	        List<String> snames = new ArrayList<>();
	        JSONObject sheets = new JSONObject();
	        for(int i=0;i<book.getNumberOfSheets();i++){
	        	String sname = book.getSheetName(i);
	        	snames.add(sname);
	        	JSONObject sheet = new JSONObject();
	        	XSSFSheet xs = book.getSheetAt(i);
	        	JSONObject info = new JSONObject();
	        	JSONArray ths = new JSONArray();
//	        	List<String> ths = new ArrayList<>(); // 表头  {"value"："col1","label":"12.2周一"}
	        	
	        	JSONArray dates = new JSONArray();	// 日期选项
	        	JSONArray list = new JSONArray();  // 表中数据
	        	List<String> teachers = new ArrayList<>(); // 教师数据
	        	Set<String> students = new LinkedHashSet<>(); // 学生数据
	        	Set<String> extras = new LinkedHashSet<>();  // 非正规学生数据
	        	String[] ctypes = {"一对一","一对二","一对三","班课","试听"};
	        	Set<String> courTypes = new LinkedHashSet<>();  // 课程类型
	        	String illegals= "休,歇班,放假,请假,休息,有事,会,会议";  //排除字符
	        	
	        	
	        	int rowNum = xs.getLastRowNum() + 1;  // 不准确,可能出现多行无数据
	            int coloumNum = xs.getRow(0).getPhysicalNumberOfCells();
	            JSONObject colRow = new JSONObject();  
    			JSONObject rowData = new JSONObject();
	            for(int j=0;j<rowNum;j++){   // 行数
	        		XSSFRow row = xs.getRow(j);
	        		if(row!=null){
	        			if(row.getCell(0)==null){ // 行中第一个单元格出现null, 直接退出
	        				break;
	        			}
	        			JSONObject obj = new JSONObject();
	        			for(int k=0;k<row.getLastCellNum();k++){
	        				if(j==0){ //第一行
	        					String thLabel = getXCellFormatValue(row.getCell(k));
	        					JSONObject th = new JSONObject();
	        					th.put("value", "col"+k);
	        					th.put("label", thLabel);
	        					ths.add(th);	// 根据k比较两日期先后顺序
			        			if(k>1){
			        				JSONObject date = new JSONObject();
			        				date.put("label", thLabel);
			        				date.put("value", k);
			        				dates.add(date);
			        			}			        			
			        		}else{
			        			// 根据rowk锁定当前行数据
			        			
			        			// 判断是否跨行， 跨行即 记录需跨行数， 同一k下， 当前行单元格记为跨行数据
//			        			记录j(行号), k列号, 数据 row.getCell(k)  k:j  k:data
			        			
			        			obj.put("idx", j);
			        			String col = "col"+String.valueOf(k);
			        			try{
				        			int rspan = getRowSpan(row.getCell(k),xs);
				        			String cellData = getXCellFormatValue(row.getCell(k));
				        			if(colRow.containsKey(col)){
				        				obj.put(col, rowData.getString(col));
				        				int max = colRow.getInteger(col); // 除去采集完成的跨行数据
				        				if(j==max){
				        					colRow.remove(col);
				        					rowData.remove(col);
				        				}			        				
				        			}else if(rspan>1){		//多行单元格
				        				colRow.put("col"+k,j+rspan-1);
				        				rowData.put("col"+k, cellData);
				        				obj.put(col, cellData);
				        				if(col.equals("col0")){
				        					teachers.add(cellData);
				        				}else if(k>1&&cellData.length()<=4&&!"".equals(cellData)&&illegals.indexOf(cellData)<0){
				        					students.add(cellData);
				        				}else if(k>1&&!"".equals(cellData)&&illegals.indexOf(cellData)<0){
				        					extras.add(cellData);
				        				}
				        			}else if(k==0&&"".equals(cellData)){  // 去除无数据行
				        				obj.remove("idx");
				        				break;
				        			}else{		//单行单元格
				        				obj.put(col, cellData);		
				        				if(col.equals("col0")){
				        					teachers.add(cellData);
				        				}else if(k>1&&cellData.length()<=4&&!"".equals(cellData)&&illegals.indexOf(cellData)<0){
				        					students.add(cellData);
				        				}else if(k>1&&!"".equals(cellData)&&illegals.indexOf(cellData)<0){
				        					extras.add(cellData);
				        				}
				        			}      			
			        			}catch(Exception e){
			        				obj.put(col, "");
			        			}
			        		}	        				
	        			}

	        			if(!obj.isEmpty()){
	        				list.add(obj);	 
	        			}	
	        			
	        		}	        		
	        	}
	        	
	        	info.put("ths", ths);
	        	info.put("dates", dates);
	        	info.put("teachers", formatList(teachers));
	        	info.put("students", formatList(students));
	        	info.put("extras", formatList(extras));
	        	if(extras.size()>0){
	        		for(String ext:extras){
	        			for(String ct:ctypes){
	        				if(ext.indexOf(ct)>-1){
	        					courTypes.add(ct);
	        				}
	        			}
	        		}	        		
	        	}
	        	info.put("courTypes", formatList(courTypes));
	        	
	        	
	        	sheet.put("info", info);
	        	sheet.put("list", list);
	        	sheets.put(sname, sheet);
	        }
	        
	        result.put("snames", snames);
	        result.put("sheets", sheets);
	        result.put("filename", file.getOriginalFilename());
	        result.put("success", true);
			result.put("error", null);
			
		}catch(Exception e){
			e.printStackTrace();
			result.put("success", false);
			result.put("error", "上传文件读取失败,请重试");
		}
		return result;
	}
	
	public static JSONArray formatList(Collection<String> list){
		JSONArray result = new JSONArray();
		for(String str:list){
			JSONObject jso = new JSONObject();
			jso.put("value", str);
			jso.put("label", str);
			result.add(jso);
		}
		return result;
	}
	
	public static int getRowSpan(Cell cell, XSSFSheet sheet) {
		int rowSpan = 1;
		try{
			List<CellRangeAddress> list = sheet.getMergedRegions();
			for (CellRangeAddress cellRangeAddress : list) {
				if (cellRangeAddress.isInRange(cell)) {
					rowSpan = cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow() + 1; 
					// +1是因为如果没跨，就算1
					break;
				}
			}
		}catch(Exception e){
			System.out.println(sheet.getSheetName()+" x:"+cell.getRowIndex()+" y:"+cell.getColumnIndex()+" --- "+e.getMessage());
		}
		return rowSpan;
	}
	
	private static String getXCellFormatValue(XSSFCell cell) {
        String cellValue = "";
        if (null != cell) {
            switch (cell.getCellType()) {
                case STRING:
                    cellValue = cell.getRichStringCellValue().getString().trim();
                    break;
                case NUMERIC:
                    cellValue = (new Double(cell.getNumericCellValue())).intValue() + "";
                    break;
                default:
                    cellValue = "";
            }
        } else {
            cellValue = "";
        }
        return cellValue;
    }
	
}
