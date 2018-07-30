package de.kehrer.rath.msgtoexcel.mapi;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Scope;
import org.springframework.stereotype.Component;
import org.springframework.util.Assert;

import de.kehrer.rath.msgtoexcel.msg.pojo.MsgContent;

@Component
@Scope("prototype")
public class MAPIMessageFileReader {

	private static final String NOT_NULL_MSG = "Directory to read msg files from (MSG_DIR) must not be null.";


	public static final String DEFAULT_MSG_DIR = "./messages";
	public static final String MSG_DIR_KEY = "MSG_DIR";
	
	@Autowired
	private MAPIMessageFileVisitor fileVisitor;

	@Value("${" + MSG_DIR_KEY + ":" + DEFAULT_MSG_DIR + "}")
	private Path excelFilesDir;
	
	public List<MsgContent> readMAPIFiles() throws IOException {		
		Assert.notNull(excelFilesDir, NOT_NULL_MSG);
		Assert.isTrue(Files.exists(excelFilesDir), "Specified messages directory " + excelFilesDir + " does not exist!");
		Files.walkFileTree(excelFilesDir, fileVisitor);
		List<MsgContent> parsedContents = fileVisitor.getMsgContents();
		return parsedContents;
	}

}
