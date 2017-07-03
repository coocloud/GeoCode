import com.google.maps.GeoApiContext;
import com.google.maps.GeocodingApi;
import com.google.maps.GeocodingApiRequest;
import com.google.maps.errors.ApiException;
import com.google.maps.model.GeocodingResult;
import com.sun.media.sound.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class Main {

    static GeoApiContext context = new GeoApiContext().setApiKey("AIzaSyDq8KUfDO5zGqotJGcuddusBZSs3vAgm1A");

    static final String FILE_NAME = "C:\\Users\\coocloud\\Projects\\Misc\\src\\main\\resources\\brokers_one.xlsx";

    static final String FOUT_FILE_NAME = "C:\\Users\\coocloud\\Projects\\Misc\\src\\main\\resources\\brokers_one_fail.xlsx";

    static final String POUT_FILE_NAME = "C:\\Users\\coocloud\\Projects\\Misc\\src\\main\\resources\\brokers_one_pass.xlsx";

    public static void main(String[] args) throws InterruptedException, ApiException, IOException, org.apache.poi.openxml4j.exceptions.InvalidFormatException {

//        GeocodingResult[] results =  GeocodingApi.geocode(context, "22 Oxford Rd, Parktown, Johannesburg, 2193").await();

        GeocodingApiRequest req = GeocodingApi.newRequest(context).address("SRC 68/A , South Ridge Centra");
        try {
            GeocodingResult[] await = req.await();
            if (await.length == 0){
                System.out.println("Cannot resolve to geo code");
            } else {
                System.out.println(await[0].formattedAddress);
                System.out.println(await[0].geometry.location);
                System.out.println("Hello, World!");
            }

        } catch (Exception e) {
            System.out.println("ERROR");
        }
        List<BrokerDetailsExcelDTO> listFromExcel = getListFromExcel();
        String something = "";
        List<BrokerDetailsExcelDTO> failed = listFromExcel.subList(1, 4);
        System.out.println(failed.size());
        System.out.println(failed.get(0).getBrokerageName());
        writeToExcel(FOUT_FILE_NAME, failed);
//        System.out.println(results[0].formattedAddress);
//        System.out.println(results[0].geometry.location);
//        System.out.println("Hello, World!");
    }

    public static List<BrokerDetailsExcelDTO> getListFromExcel() throws IOException, org.apache.poi.openxml4j.exceptions.InvalidFormatException {
        Workbook workbook = null;
        List<BrokerDetailsExcelDTO> brokerList = new ArrayList<BrokerDetailsExcelDTO>();
        try {
            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
            workbook = WorkbookFactory.create(excelFile);
            Sheet sheet = workbook.getSheetAt(0);

            for (Row row: sheet){
                if (row.getRowNum() == 0){
                    //extract property names
                    //Brokerage Name, FSP Number, Product, Contact Person, Phone 1, Phone 2, Fax, Email, Website, Physical Address, Postal Address, Country, Province, City, Suburb, Date Joined, Last Activity, Status
                }else {
                    //TODO: do cell type check
                    BrokerDetailsExcelDTO brokerDetailsExcelDTO = new BrokerDetailsExcelDTO();
                    brokerDetailsExcelDTO.setBrokerageName(row.getCell(0).getStringCellValue());
                    brokerDetailsExcelDTO.setFspNumber(String.valueOf(row.getCell(1).getNumericCellValue()));
                    brokerDetailsExcelDTO.setProduct(row.getCell(2).getStringCellValue());
                    brokerDetailsExcelDTO.setContactPerson(row.getCell(3).getStringCellValue());

                    Cell cell = row.getCell(4);
                    cell.setCellType(CellType.STRING); //formatted cell phone 1
                    brokerDetailsExcelDTO.setPhoneOne(cell.getStringCellValue());


                    cell = row.getCell(5);
                    cell.setCellType(CellType.STRING); //formatted cell phone 2
                    brokerDetailsExcelDTO.setPhoneTwo(cell.getStringCellValue());


                    cell = row.getCell(6);
                    cell.setCellType(CellType.STRING); //formatted cell fax
                    brokerDetailsExcelDTO.setFax(cell.getStringCellValue());

                    brokerDetailsExcelDTO.setEmail(row.getCell(7).getStringCellValue());
                    brokerDetailsExcelDTO.setWebsite(row.getCell(8).getStringCellValue());
                    brokerDetailsExcelDTO.setPhysicalAddress(row.getCell(9).getStringCellValue());
                    brokerDetailsExcelDTO.setPostalAddress(row.getCell(10).getStringCellValue());
                    brokerDetailsExcelDTO.setCountry(row.getCell(11).getStringCellValue());
                    brokerDetailsExcelDTO.setProvince(row.getCell(12).getStringCellValue());
                    brokerDetailsExcelDTO.setCity(row.getCell(13).getStringCellValue());
                    brokerDetailsExcelDTO.setSuburb(row.getCell(14).getStringCellValue());
                    brokerDetailsExcelDTO.setDateJoined(row.getCell(15).getDateCellValue());
                    brokerDetailsExcelDTO.setLastActivity(row.getCell(16).getDateCellValue());
                    brokerDetailsExcelDTO.setStatus(row.getCell(17).getStringCellValue());
                    brokerList.add(brokerDetailsExcelDTO);
//                    for (Cell cell: row) {
//
//                    }
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
            //throw something
        } finally {
            if (workbook != null) {
                workbook.close();
            }

        }
        return brokerList;
    }

    public static void writeToExcel(String THE_FILE_NAME, List<BrokerDetailsExcelDTO> myArrayList) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("sampleExcel");
        Object[][] datatypes = {
                {"Datatype", "Type", "Size(in bytes)"},
                {"int", "Primitive", 2},
                {"float", "Primitive", 4},
                {"double", "Primitive", 8},
                {"char", "Primitive", 1},
                {"String", "Non-Primitive", "No fixed size"}
        };

        int rowNum = 0;
        System.out.println("Creating excel");

        for (BrokerDetailsExcelDTO brokerDetailsExcelDTO: myArrayList) {
            Row row = sheet.createRow(rowNum++);
            int colNum = 0;
            Cell cell = row.createCell(colNum++);
            cell.setCellValue((String) brokerDetailsExcelDTO.getBrokerageName());
            cell = row.createCell(colNum++);
            cell.setCellValue(brokerDetailsExcelDTO.getFspNumber());
            cell = row.createCell(colNum++);
            cell.setCellValue(brokerDetailsExcelDTO.getProduct());
            cell = row.createCell(colNum++);
            cell.setCellValue(brokerDetailsExcelDTO.getContactPerson());
            cell = row.createCell(colNum++);
            cell.setCellValue(brokerDetailsExcelDTO.getPhoneOne());
            cell = row.createCell(colNum++);
            cell.setCellValue(brokerDetailsExcelDTO.getPhoneTwo());
            cell = row.createCell(colNum++);
            cell.setCellValue(brokerDetailsExcelDTO.getFax());
            cell = row.createCell(colNum++);
            cell.setCellValue(brokerDetailsExcelDTO.getEmail());
            cell = row.createCell(colNum++);
            cell.setCellValue(brokerDetailsExcelDTO.getWebsite());
            cell = row.createCell(colNum++);
            cell.setCellValue(brokerDetailsExcelDTO.getPhysicalAddress());
            cell = row.createCell(colNum++);
            cell.setCellValue(brokerDetailsExcelDTO.getPostalAddress());
            cell = row.createCell(colNum++);
            cell.setCellValue(brokerDetailsExcelDTO.getCountry());
            cell = row.createCell(colNum++);
            cell.setCellValue(brokerDetailsExcelDTO.getProvince());
            cell = row.createCell(colNum++);
            cell.setCellValue(brokerDetailsExcelDTO.getCity());
            cell = row.createCell(colNum++);
            cell.setCellValue(brokerDetailsExcelDTO.getSuburb());
            cell = row.createCell(colNum++);
            cell.setCellValue(brokerDetailsExcelDTO.getDateJoined());
            cell = row.createCell(colNum++);
            cell.setCellValue(brokerDetailsExcelDTO.getLastActivity());
            cell = row.createCell(colNum++);
            cell.setCellValue(brokerDetailsExcelDTO.getStatus());
        }

//        for (Object[] datatype : datatypes) {
//            Row row = sheet.createRow(rowNum++);
//            int colNum = 0;
//            for (Object field : datatype) {
//                Cell cell = row.createCell(colNum++);
//                if (field instanceof String) {
//                    cell.setCellValue((String) field);
//                } else if (field instanceof Integer) {
//                    cell.setCellValue((Integer) field);
//                }
//            }
//        }

        try {
            FileOutputStream outputStream = new FileOutputStream(THE_FILE_NAME);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Done");
    }

}