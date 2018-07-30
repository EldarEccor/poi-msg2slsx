package de.kehrer.rath.msgtoexcel.mapi;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Calendar;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hsmf.MAPIMessage;
import org.apache.poi.hsmf.exceptions.ChunkNotFoundException;
import org.springframework.stereotype.Component;
import org.springframework.util.Assert;
import org.springframework.util.StringUtils;

import de.kehrer.rath.msgtoexcel.msg.pojo.MsgContent;
import de.kehrer.rath.msgtoexcel.msg.pojo.MsgContent.MsgContentBuilder;
import de.kehrer.rath.msgtoexcel.msg.pojo.PostalAddress;
import de.kehrer.rath.msgtoexcel.msg.pojo.PostalAddress.PostalAddressBuilder;
import lombok.extern.slf4j.Slf4j;

@Slf4j
@Component
public class MAPIMessageFileParser {

	private static final String DEFAULT_COUNTRY = "Deutschland";

	private static final String ADDRESS_PATTERN = "(Anschrift:)\\r\\n(.*)\\r\\n(.*)\\r\\n(.*)\\r\\n(.*)\\r\\n";
	private static final String VOUCHER_CODE_PATTERN = "(Gutscheincode:)(.*)\\r\\n";
	private static final String EMAIL_PATTERN = "(E-Mail:)(.*)\\r\\n";

	public MsgContent parseMsgFile(Path msgFile) throws IOException {
		Assert.notNull(msgFile, "parseMsgFile: msgFile must not be null.");
		String absoluteMsgFileName = msgFile.toString();
		log.trace("Start parsing MSG file: {}.", absoluteMsgFileName);
		boolean msgFileExists = Files.exists(msgFile);
		if (!msgFileExists) {
			log.error("MSG file does not exist: {}.", absoluteMsgFileName);
			throw new FileNotFoundException(absoluteMsgFileName);
		}

		MsgContentBuilder contentBuilder = MsgContent.builder();
		PostalAddressBuilder addressBuilder = PostalAddress.builder();

		try (MAPIMessage mapiMsg = new MAPIMessage(absoluteMsgFileName)) {
			try {
				String textBody = mapiMsg.getTextBody();
				log.trace("Message body: {} ", textBody);
				extractVoucherCode(contentBuilder, textBody);
				extractAddress(addressBuilder, textBody);
				extractMsgDate(contentBuilder, mapiMsg);
				extractEmail(contentBuilder, textBody);
			} catch (ChunkNotFoundException e) {
				log.error("{} : Message text body is corrupt or cannot be found.", absoluteMsgFileName);
			}
		}
		contentBuilder.address(addressBuilder.build());
		MsgContent content = contentBuilder.build();
		log.trace("Message parsing finished: {} ", content);
		return content;
	}

	private void extractEmail(MsgContentBuilder contentBuilder, String textBody) {
		try {

			Pattern pattern = Pattern.compile(EMAIL_PATTERN);
			Matcher matcher = pattern.matcher(textBody);
			if (matcher.find()) {
				String eMail = matcher.group(2).trim();
				contentBuilder.eMail(eMail);
			} else {
				log.warn("Email regexp did not match the msg body: {}", textBody);
			}
		} catch (Exception e) {
			log.error("Error while extracting eMail", e);
		}
	}

	private void extractMsgDate(MsgContentBuilder contentBuilder, MAPIMessage mapiMsg) {
		try {
			Calendar messageDate = mapiMsg.getMessageDate();
			contentBuilder.messageDate(messageDate);
		} catch (ChunkNotFoundException e) {
			log.error("Error while extracting message date.", e);
		}
	}

	private void extractAddress(PostalAddressBuilder addressBuilder, String textBody) {
		try {
			Pattern pattern = Pattern.compile(ADDRESS_PATTERN);
			Matcher matcher = pattern.matcher(textBody);
			if (matcher.find()) {
				String fullName = matcher.group(2).trim();
				if (fullName.startsWith("Herr ") || fullName.startsWith("Frau ")) {
					fullName = fullName.substring(5);
				}
				int lastSpaceIndex = fullName.lastIndexOf(" ");
				String firstName = fullName.substring(0, lastSpaceIndex).trim();
				String lastName = fullName.substring(lastSpaceIndex).trim();
				String streetAndHouseNbr = matcher.group(3).trim();
				String zipCodeAndCity = matcher.group(4).trim();

				int firstSpaceIndex = zipCodeAndCity.indexOf(" ");
				String zipCode = zipCodeAndCity.substring(0, firstSpaceIndex).trim();
				String city = zipCodeAndCity.substring(firstSpaceIndex).trim();

				addressBuilder.firstName(firstName).lastName(lastName).streetAndHouseNumber(streetAndHouseNbr)
						.zipCode(zipCode).city(city).country(DEFAULT_COUNTRY);

				String possibleCountry = matcher.group(5).trim();
				if (!StringUtils.isEmpty(possibleCountry)) {
					addressBuilder.country(possibleCountry);
				}
			} else {
				log.warn("Address regexp did not match the msg body: {}", textBody);
			}
		} catch (Exception e) {
			log.error("Error trying to parse address.", e);
		}
	}

	private void extractVoucherCode(MsgContentBuilder contentBuilder, String textBody) {
		try {
			Pattern pattern = Pattern.compile(VOUCHER_CODE_PATTERN);
			Matcher matcher = pattern.matcher(textBody);
			if (matcher.find()) {
				contentBuilder.voucherCode(matcher.group(2));
			} else {
				log.warn("Voucher regexp did not match the msg body: {}", textBody);
			}
		} catch (Exception e) {
			log.error("Error trying to parse voucher.", e);
		}
	}

}
