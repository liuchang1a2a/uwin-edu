package com.service;

import org.springframework.web.multipart.MultipartFile;

import com.alibaba.fastjson.JSONObject;

public interface IEService {
	
	public JSONObject getXlsInfo(MultipartFile file);
	public JSONObject getXlsxInfo(MultipartFile file);
	
}
