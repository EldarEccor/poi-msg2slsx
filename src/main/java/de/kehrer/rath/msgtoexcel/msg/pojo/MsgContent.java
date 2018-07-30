package de.kehrer.rath.msgtoexcel.msg.pojo;

import java.io.Serializable;
import java.util.Calendar;

import lombok.Builder;
import lombok.Data;

@Builder
@Data
public class MsgContent implements Serializable {

	private static final long serialVersionUID = -5911347915590466046L;
	
	PostalAddress address;
	
	String voucherCode;	
	
	Calendar messageDate;
	
	String eMail;
}
