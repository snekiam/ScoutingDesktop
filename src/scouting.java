import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.FileSystems;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardWatchEventKinds;
import java.nio.file.WatchEvent;
import java.nio.file.WatchKey;
import java.nio.file.WatchService;
import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class scouting {
	public void Run() throws IOException, InterruptedException{
		/*Path dataFolder = Paths.get("C:\\Scouting Data Folder");
		WatchService watchService = FileSystems.getDefault().newWatchService();
		dataFolder.register(watchService, StandardWatchEventKinds.ENTRY_CREATE);
		String fileName = null;
		boolean valid = true;
		do {
			WatchKey watchKey = watchService.take();

			for (WatchEvent event : watchKey.pollEvents()) {
				WatchEvent.Kind kind = event.kind();
				if (StandardWatchEventKinds.ENTRY_CREATE.equals(event.kind())) {
					fileName = event.context().toString();
					System.out.println("File Created:" + fileName);
				}
			}
			valid = watchKey.reset();

		} while (true);*/
		while(true){
			File f = new File("C:\\Scouting Data Folder\\scouting(5).xls");
			if(f.exists() && !f.isDirectory()) {
				int files = 0;
				File file = new File("");
				while(files<6){
					switch (files){
					case 0:{
						file =new File("C:\\Scouting Data Folder\\scouting.xls");
						System.out.println(file);
						break;}
					case 1:{
						file =new File("C:\\Scouting Data Folder\\scouting(1).xls");
						System.out.println(file);
						break;}
					case 2:{
						file =new File("C:\\Scouting Data Folder\\scouting(2).xls");
						System.out.println(file);
						break;}
					case 3:{
						file =new File("C:\\Scouting Data Folder\\scouting(3).xls");
						System.out.println(file);
						break;}
					case 4:{
						file =new File("C:\\Scouting Data Folder\\scouting(4).xls");
						System.out.println(file);
						break;}
					case 5:{
						file =new File("C:\\Scouting Data Folder\\scouting(5).xls");
						System.out.println(file);
						break;}
					}
					System.out.println("Hi");
					

					FileInputStream alldata= null;
					FileInputStream input = null;
					try{
						alldata = new FileInputStream(new File("scouting.xls"));
					}catch (FileNotFoundException e1){
						try{
							JOptionPane.showMessageDialog(null, "No scouting data document found. A new one has been created.  Please run the program again. ");
							FileOutputStream out = new FileOutputStream("scouting.xls");
							HSSFWorkbook scoutingwb = new HSSFWorkbook();
							HSSFSheet scoutingsheet = scoutingwb.createSheet("scoutingdata");
							Row rowNull = scoutingsheet.createRow(0);
							scoutingwb.write(out);
							out.close();
							alldata = new FileInputStream(new File("scouting.xls"));
						}
						catch(IOException e2){
							JOptionPane.showMessageDialog(null, "The file can't be written to. Check the permissions");
						}
					}
					try{
						input = new FileInputStream(file);
					}catch(FileNotFoundException e2){
						JOptionPane.showMessageDialog(null,"There is no scouting data to input.  Try again.");
					}
					HSSFWorkbook scoutingwb = new HSSFWorkbook(alldata);
					HSSFSheet scoutingsheet = scoutingwb.getSheetAt(0);
					HSSFWorkbook inputwb = new HSSFWorkbook(input);
					HSSFSheet inputsheet = inputwb.getSheetAt(0);
					int lastrow = scoutingsheet.getLastRowNum();
					int nextrow = lastrow+1;
					int i = 0;
					HSSFSheet sheet = scoutingwb.getSheetAt(0);
					Row row = sheet.createRow(nextrow);
					while(i<16){
						Cell cell = row.createCell(i);
						Row row1 = inputsheet.getRow(0);
						Cell cell1 = row1.getCell(i);
						int cellType = cell1.getCellType();
						if(cellType == 0){
							double cellvalue1 = cell1.getNumericCellValue();
							cell.setCellValue(cellvalue1);
							i++;
						}
						if(cellType == 1){
							String cellvalue2 = cell1.getStringCellValue();
							System.out.println(cellvalue2);
							cell.setCellValue(cellvalue2);
							i++;
						}
						FileOutputStream output = new FileOutputStream("scouting.xls");
						scoutingwb.write(output);
						output.close();
						file.delete();
					}					 
					files++;
				}
			}
		}
	}
}