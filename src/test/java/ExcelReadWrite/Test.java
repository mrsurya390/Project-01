package ExcelReadWrite;
import java.io.*; import java.util.*; import org.apache.poi.xssf.usermodel.*;
public class Test {
	public static void main(String[] args) {
		File file = new File("C:\\Users\\Karna\\eclipse-workspace\\Test\\target\\Book.xlsx");
		try {
			FileOutputStream f = new FileOutputStream(file);
			XSSFWorkbook work = new XSSFWorkbook();
			XSSFSheet sheet = work.createSheet("Student details");
			XSSFRow row = sheet.createRow(0);
			String arr[] = {"Name", "Age", "Marks"};
			for(int i=0; i<arr.length;i++) {
				XSSFCell c = row.createCell(i);
				c.setCellValue(arr[i]);
			}
			Random rand = new Random();
			List<String> l = new ArrayList<>();
			l.add("Vijay"); l.add("Sam"); l.add("Suriya"); l.add("Vikram");
			for(int i=0; i<l.size();i++) {
				XSSFRow r = sheet.createRow(i+1);
				for(int j=0; j<3;j++) {
					XSSFCell c = r.createCell(j);
					if (j==0) c.setCellValue(l.get(i));
					else if(j==1) c.setCellValue(rand.nextInt(19,24));
					else c.setCellValue(rand.nextInt(100));
				}
			}
			work.write(f); f.close();
		} catch (Exception e) {
			e.printStackTrace();}}}
