package de.kehrer.rath.msgtoexcel.config;

import java.nio.file.Path;
import java.nio.file.Paths;

import org.springframework.core.convert.converter.Converter;
import org.springframework.stereotype.Component;

@Component
public class StringToPathConverter implements Converter<String,Path>{

	 @Override
	 public Path convert(String path) {
	     return Paths.get(path);
	 }
	
}
