package de.kehrer.rath.msgtoexcel.msg.pojo;

import java.io.Serializable;

import lombok.Builder;
import lombok.Data;

@Builder
@Data
public class PostalAddress implements Serializable {

	private static final long serialVersionUID = -8452202015055849886L;

	String zipCode;
	
	String city;

	String country;

	String streetAndHouseNumber;
	
	String firstName;
	
	String lastName;

}
