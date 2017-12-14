package com.github.sep6th.core;

import java.util.List;

import org.apache.commons.configuration.HierarchicalConfiguration.Node;
import org.apache.commons.configuration.tree.ConfigurationNode;

public class XmlCURD {

	private static Node root = ModelConfig.cfg.getRoot();
	
	//获取 sheet 节点
	public static ConfigurationNode getSheet(Integer sheetCode){
		sheetCode = sheetCode == null ? 0 : sheetCode;
		return root.getChildren("sheets").get(0).getChildren("sheet").get(sheetCode);
	}
	
	//获取第 sheetCode 个 sheet 的 columns节点
	public static ConfigurationNode getColumns(Integer sheetCode){
		return getSheet(sheetCode).getChildren("columns").get(0);
	}
	
	
	//获取
    public static String getXmlStringValue(String key) {
        return ModelConfig.cfg.getString(key);
    }
    
    public static List<Object> getList(String key) {
        return ModelConfig.cfg.getList(key);
    }
    
	
}
