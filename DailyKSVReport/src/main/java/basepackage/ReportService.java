package basepackage;

import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.io.File;
import java.io.IOException;
import java.time.LocalDate;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.fasterxml.jackson.databind.ObjectMapper;

public class ReportService {
	String workingDirectory = System.getProperty("user.dir");
	File excelFile = new File(workingDirectory+"\\Freshers Track Sheet.xlsx");
	File dataFile = new File(workingDirectory+"\\data.json");
	
	public void fetchData(){
		try {
			ObjectMapper mapper = new ObjectMapper();
			Data data = mapper.readValue(dataFile, Data.class);
			LocalDate date = LocalDate.now();
			System.out.println("Getting excel file................");
			Workbook wb = WorkbookFactory.create(excelFile);
			Sheet sheet1 = wb.getSheetAt(data.getSheetNo());
			System.out.println("Reading Contents..............");
			Iterator<Row> it = sheet1.rowIterator();
			it.next();
			StringBuilder out = new StringBuilder();
			System.out.println("Generating report for "+sheet1.getSheetName()+"...........................\n");
			out.append("*Daily Report    Date: "+date.toString()+"*\n-----------------------------------------\n");
			while(it.hasNext()) {
				Row r = it.next();
				
				Cell cell1 = r.getCell(0);
				Double cellValue = cell1.getNumericCellValue();
				LocalDate currDate = convertExcelSerialDateToLocalDate(cellValue);
				if(currDate!=null &&currDate.compareTo(date)==0) {
				boolean topic1Flag = false;
				boolean topic2Flag = false;
				boolean topic3Flag = false;
				boolean topic4Flag = false;
				
				String topic1 = r.getCell(1).getStringCellValue();
				if(!topic1.isEmpty()&&!topic1.equals("")&&!topic1.equals(" ")) topic1Flag = true;
				String topic1_details = r.getCell(2).getStringCellValue();
				String topic1_hoursSpent = r.getCell(3).getNumericCellValue()+" hrs";
				
				String topic2 = r.getCell(4).getStringCellValue();
				if(!topic2.isEmpty()&&!topic1.equals("")&&!topic1.equals(" ")) topic2Flag = true;
				String topic2_details = r.getCell(5).getStringCellValue();
				String topic2_hoursSpent = r.getCell(6).getNumericCellValue()+ " hrs";
				
				String topic3 = r.getCell(7).getStringCellValue();
				if(!topic3.isEmpty()&&!topic1.equals("")&&!topic1.equals(" ")) topic3Flag = true;
				String topic3_details = r.getCell(8).getStringCellValue();
				String topic3_hoursSpent = r.getCell(9).getNumericCellValue()+" hrs";
				
				String topic4 = r.getCell(10).getStringCellValue();
				if(!topic4.isEmpty()&&!topic1.equals("")&&!topic1.equals(" ")) topic4Flag = true;
				String topic4_details = r.getCell(11).getStringCellValue();
				String topic4_hoursSpent = r.getCell(12).getNumericCellValue()+ " hrs";
				
				String totalHoursSpent = r.getCell(13).getNumericCellValue()+" hrs";
				
				
					
					if(topic1Flag) {
						out.append("*Topic 1:* "+topic1+"\n");
						out.append("*Topic 1 Details:* "+topic1_details+"\n");
						out.append("*Topic 1 Hours Spent:* *_"+topic1_hoursSpent+"_*\n\n");
					}
					
					if(topic2Flag) {
						out.append("*Topic 2:* "+topic2+"\n");
						out.append("*Topic 2 Details:* "+topic2_details+"\n");
						out.append("*Topic 2 Hours Spent:* *_"+topic2_hoursSpent+"_*\n\n");
					}
					
					if(topic3Flag) {
						out.append("*Topic 3:* "+topic3+"\n");
						out.append("*Topic 3 Details:* "+topic3_details+"\n");
						out.append("*Topic 3 Hours Spent:* *_"+topic3_hoursSpent+"_*\n\n");
					}
					
					if(topic4Flag) {
						out.append("*Topic 4:* "+topic4+"\n");
						out.append("*Topic 4 Details:* "+topic4_details+"\n");
						out.append("*Topic 4 Hours Spent:* *_"+topic4_hoursSpent+"_*\n\n");
					}
					
					out.append("*Total hours spent for the day:* *_"+totalHoursSpent+"_*");
					break;
				}
			}
			System.out.println(out.toString());
			StringSelection stringSelection = new StringSelection(out.toString());
			Toolkit.getDefaultToolkit().getSystemClipboard().setContents(stringSelection, null);
	        System.out.println("\nReport copied to clipboard!");
		} catch (EncryptedDocumentException e) {
			System.out.println("File is encrypted!!");
			e.printStackTrace();
		} catch (IOException e) {
			System.out.println("IO error!!");
			e.printStackTrace();
		}
	}
	// Method to convert Excel serial date to LocalDate
    public static LocalDate convertExcelSerialDateToLocalDate(double excelDate) {
    	if((int)excelDate ==0) {
    		return null;
    	}
        // Excel's base date (1900-01-01)
        LocalDate startDate = LocalDate.of(1900, 1, 1);

        // Adjust for Excel's leap year bug (subtract 2 days)
        return startDate.plusDays((int) excelDate - 2);
    }
}
