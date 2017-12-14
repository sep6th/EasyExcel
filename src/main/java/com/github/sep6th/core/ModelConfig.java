package com.github.sep6th.core;

import org.apache.commons.configuration.ConfigurationException;
import org.apache.commons.configuration.XMLConfiguration;
import org.apache.commons.configuration.reloading.FileChangedReloadingStrategy;

public class ModelConfig {
	
	public static final String xmlFileName = "easyExcel-config.xml";
    public static XMLConfiguration cfg = null;
    static {
        try {
            cfg = new XMLConfiguration(xmlFileName);
        } catch (ConfigurationException e) {
            e.printStackTrace();
        }
        // 配置文件 发生变化就重新加载
        cfg.setReloadingStrategy(new FileChangedReloadingStrategy());
    }

}
