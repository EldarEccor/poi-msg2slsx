package de.kehrer.rath.msgtoexcel.mapi;

import java.io.IOException;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.FileVisitResult;
import java.nio.file.FileVisitor;
import java.nio.file.Path;
import java.nio.file.PathMatcher;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Scope;
import org.springframework.stereotype.Component;

import de.kehrer.rath.msgtoexcel.msg.pojo.MsgContent;
import lombok.extern.slf4j.Slf4j;

@Slf4j
@Component
@Scope("prototype")
public class MAPIMessageFileVisitor implements FileVisitor<Path> {

	static final String SUCCESS_MSG = "Successfully parsed msg file: {}. Result is {}";

	static final String GLOB_MSG_MATCH_EXPR = "glob:*.msg";

	static final String EXCEPTION_OCCURED_MSG = "Exception occured during processing of {}.";

	static final String FINISHED_ANALYSIS_MSG = "Finished analysis of directory: {}.";

	static final String START_MSG = "Start analysis of directory: {}.";
	
	private List<MsgContent> msgContents = new ArrayList<>();

	@Autowired
	private MAPIMessageFileParser fileParser;

	@Override
	public FileVisitResult preVisitDirectory(Path path, BasicFileAttributes attrs) throws IOException {
		log.info(START_MSG, path.getFileName());
		return FileVisitResult.CONTINUE;
	}

	@Override
	public FileVisitResult postVisitDirectory(Path path, IOException ex) throws IOException {
		log.info(FINISHED_ANALYSIS_MSG, path.getFileName());
		if (ex != null) {
			log.error(EXCEPTION_OCCURED_MSG, path.getFileName(), ex);
		}
		return FileVisitResult.CONTINUE;
	}

	@Override
	public FileVisitResult visitFile(Path path, BasicFileAttributes attrs) throws IOException {
		FileSystem fileSystem = FileSystems.getDefault();
		PathMatcher pathMatcher = fileSystem.getPathMatcher(GLOB_MSG_MATCH_EXPR);
		if (pathMatcher.matches(path.getFileName())) {
			MsgContent parsedContent = fileParser.parseMsgFile(path);
			log.info(SUCCESS_MSG, path.getFileName(), parsedContent);
			msgContents.add(parsedContent);
		}
		return FileVisitResult.CONTINUE;
	}

	@Override
	public FileVisitResult visitFileFailed(Path file, IOException exc) throws IOException {
		log.error(EXCEPTION_OCCURED_MSG, file.getFileName(), exc);
		return FileVisitResult.CONTINUE;
	}
	
	public List<MsgContent> getMsgContents() {
		return new ArrayList<>(msgContents);
	}
}
