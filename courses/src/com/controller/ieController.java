package com.controller;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpSession;

import org.apache.commons.io.FileUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import com.alibaba.fastjson.JSONObject;
import com.service.IEService;

@Controller
@RequestMapping("/ie")
public class ieController {
	
	@Autowired
	private IEService ieService;
	
	
	
	@RequestMapping("/getFileInfo")
	@ResponseBody
	public String getFileInfo(HttpServletRequest req,@RequestParam(value="file")MultipartFile file){	
		// 获取文件， excel解析，拼接，返回列表数据
//		String fileSizeLimit = req.getParameter("fileSizeLimit") == null ? "0" : req.getParameter("fileSizeLimit");
//		String fileTypeExts = req.getParameter("fileTypeExts") == null ? "0" : req.getParameter("fileTypeExts");
		JSONObject jso = new JSONObject();
		try{
			String ext = file.getOriginalFilename().split("\\.")[1];
			// xls,xlsx
			if("xls".equals(ext)){
				jso = ieService.getXlsInfo(file);
			}else if("xlsx".equals(ext)){
				jso = ieService.getXlsxInfo(file);
			}			
		}catch(Exception e){
			e.printStackTrace();
			jso.put("success", false);
			jso.put("error", "上传文件读取失败,请重试");
		}
		
		return jso.toString();
	}	
	
	
	@RequestMapping("/delFile")
	@ResponseBody
	public String delFile(HttpServletRequest req){
		String fpath = req.getParameter("fpath");
		System.out.println(fpath);
		JSONObject jso = new JSONObject();
		jso.put("code", "0");
		jso.put("desc", "ok");
		jso.put("fpath", "");
		
		return jso.toString();
	}
	
	
	
}
