package de.kehrer.rath.msgtoexcel;

import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import de.kehrer.rath.msgtoexcel.mapi.MAPIMessageFileReader;
import de.kehrer.rath.msgtoexcel.msg.pojo.MsgContent;
import de.kehrer.rath.msgtoexcel.xlsx.XlsxSummaryWriter;
import lombok.extern.slf4j.Slf4j;

@Slf4j
@SpringBootApplication
public class MsgToExcelApplication implements CommandLineRunner {

	private static final String CLI_BOOTSTRAP_MSG = "Application started with CLI arguments: {}. To kill this application, press Ctrl + C.";

	@Autowired
	private MAPIMessageFileReader msgReader;
	
	@Autowired
	private XlsxSummaryWriter msgWriter;

	public static void main(String[] args) {
		SpringApplication.run(MsgToExcelApplication.class, args);		
	}
	
	@Override
	public void run(String... args) throws Exception {
		log.info(CLI_BOOTSTRAP_MSG, Arrays.toString(args));
		List<MsgContent> msgContents = msgReader.readMAPIFiles();
		
		String excelFilename = new SimpleDateFormat("yyyy.MM.dd_HH-mm-ss").format(new Date()) + ".xlsx";
		
		msgWriter.writeSummary(excelFilename, msgContents);
	}
	
	
}
