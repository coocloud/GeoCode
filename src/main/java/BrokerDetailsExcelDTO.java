import lombok.Data;

import java.util.Date;

/**
 * Created by coocloud on 7/3/2017.
 */
@Data
public class BrokerDetailsExcelDTO {
    private String brokerageName;
    private String fspNumber;
    private String product;
    private String contactPerson;
    private String phoneOne;
    private String phoneTwo;
    private String fax;
    private String email;
    private String website;
    private String physicalAddress;
    private String postalAddress;
    private String country;
    private String province;
    private String city;
    private String suburb;
    private Date dateJoined;
    private Date lastActivity;
    private String status;
}