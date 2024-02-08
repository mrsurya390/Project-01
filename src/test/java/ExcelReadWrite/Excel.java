package ExcelReadWrite; import java.io.*;

import org.apache.poi.xssf.usermodel.*;
public class Excel {
	public static void main(String[] args) throws Throwable {
		File file = new File("C:\\Users\\Karna\\eclipse-workspace\\Test\\target\\Book.xlsx");
		FileInputStream f = new FileInputStream(file);
		XSSFWorkbook work = new XSSFWorkbook(f);
		XSSFSheet s = work.getSheet("Student details");
		for(int i=0;i<s.getPhysicalNumberOfRows();i++) {
			XSSFRow r = s.getRow(i);
			for(int j=0;j<r.getPhysicalNumberOfCells();j++) {
				XSSFCell c = r.getCell(j);
				int ct = c.getCellType();
				if(ct==1) {
					String str = c.getStringCellValue();
					if (str.equals("Sam")) {
						c.setCellValue("Ajith");
					}
					System.out.print(str+"  ");
				}
				else {
						double d = c.getNumericCellValue();
						long l = (long) d;
						String strInt = String.valueOf(l);
						System.out.print(strInt+"  ");}
				}System.out.println();
			}
		FileOutputStream fi = new FileOutputStream(file);
		work.write(fi);
		fi.close();
		}}
