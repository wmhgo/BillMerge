import java.util.*;
import java.io.*;
import java.io.File;

import jxl.Workbook;
import jxl.Sheet;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.Number;
import jxl.write.biff.*;

class CallEntry implements Comparable<CallEntry> {
	public String number;
	public int localCallTime;
	public int longDistanceCallTime;
	public int totalCallTime;
	public double localCallFee;
	public double longDistanceCallFee;
	public double totalCallFee;

	public int compareTo(CallEntry ent) {
		return totalCallFee > ent.totalCallFee ? -1
				: totalCallFee == ent.totalCallFee ? 0 : 1;
	}

	public String toString() {
		return "\t" + number + "," + localCallTime + "," + longDistanceCallTime
				+ "," + totalCallTime + "," + localCallFee + ","
				+ longDistanceCallFee + "," + totalCallFee;
	}
}

public class BillTune {
	private static final String LOCAL_CALL = "市话";
	private static final String LONG_CALL = "长途";

	private static class DataFileFilter implements FilenameFilter {
		private String pat;

		public DataFileFilter(String pattern) {
			pat = pattern;
		}

		public boolean accept(File dir, String name) {
			return name.matches(pat);
		}
	}

	public static void main(String[] args) throws FileNotFoundException, IOException, BiffException, RowsExceededException, WriteException{
		String[] dataFiles;
		File curDir = new File(".");
		dataFiles = curDir.list(new DataFileFilter(".*\\.xls$"));

		for (String dataFile : dataFiles) {
			HashMap<String, CallEntry> callMap = new HashMap<String, CallEntry>();
			CallEntry callEntry;
			ArrayList<CallEntry> entryArr;
			Workbook book = Workbook.getWorkbook(new File(dataFile));
			int sheetNum = book.getNumberOfSheets();

			for (int j = 0; j < sheetNum; j++) {
				Sheet sheet = book.getSheet(j);
				int rows = sheet.getRows();

				for (int i = 1; i < rows; i++) {
					String callNumber = sheet.getCell(0, i).getContents().substring(2);
					int callTime = Integer.parseInt(sheet.getCell(4, i)
							.getContents());
					double callFee = Double.parseDouble(sheet.getCell(5, i)
							.getContents());
					String callType = sheet.getCell(1, i).getContents();

					callEntry = callMap.get(callNumber);

					if (null == callEntry) {
						CallEntry newEntry = new CallEntry();
						newEntry.number = callNumber;
						
						if (callType.matches(LOCAL_CALL)) {
							newEntry.localCallTime = callTime;
							newEntry.localCallFee = callFee;
							newEntry.totalCallTime = callTime;
							newEntry.totalCallFee = callFee;
						} else if (callType.matches(LONG_CALL)) {
							newEntry.longDistanceCallTime = callTime;
							newEntry.longDistanceCallFee = callFee;
							newEntry.totalCallTime = callTime;
							newEntry.totalCallFee = callFee;
						} else {
							System.out.println("unknown call type:" + callType);
							continue;
						}

						callMap.put(callNumber, newEntry);
					} else if (callType.matches(LOCAL_CALL)) {
						callEntry.localCallTime += callTime;
						callEntry.localCallFee += callFee;
						callEntry.totalCallTime += callTime;
						callEntry.totalCallFee += callFee;
					} else if (callType.matches(LONG_CALL)) {
						callEntry.longDistanceCallTime += callTime;
						callEntry.longDistanceCallFee += callFee;
						callEntry.totalCallTime += callTime;
						callEntry.totalCallFee += callFee;
					} else {
						System.out.println("unknown call type:" + callType);
						continue;
					}
				}
			}

			book.close();
			
			File resultDir = new File("results");
			if (!resultDir.exists())
				resultDir.mkdir();
			
			WritableWorkbook outXls = Workbook.createWorkbook(new File("results" + File.separator + dataFile));
			WritableSheet ws = outXls.createSheet("Sheet1", 0);
			Label label1 = new Label(0, 0, "号码");
			Label label2 = new Label(1, 0, "市话时长");
			Label label3 = new Label(2, 0, "长途时长");
			Label label4 = new Label(3, 0, "总时长");
			Label label5 = new Label(4, 0, "市话费用");
			Label label6 = new Label(5, 0, "长途费用");
			Label label7 = new Label(6, 0, "总费用");

			ws.addCell(label1);
			ws.addCell(label2);
			ws.addCell(label3);
			ws.addCell(label4);
			ws.addCell(label5);
			ws.addCell(label6);
			ws.addCell(label7);
			
			entryArr = new ArrayList<CallEntry>(callMap.values());
			Collections.sort(entryArr);
			int k = 1;
			
			for (CallEntry entry : entryArr) {
				ws.addCell(new Label(0, k, entry.number));
				ws.addCell(new Number(1, k, entry.localCallTime));
				ws.addCell(new Number(2, k, entry.longDistanceCallTime));
				ws.addCell(new Number(3, k, entry.totalCallTime));
				ws.addCell(new Number(4, k, entry.localCallFee));
				ws.addCell(new Number(5, k, entry.longDistanceCallFee));
				ws.addCell(new Number(6, k, entry.totalCallFee));
				k++;
			}

			outXls.write();
			outXls.close();

		}
	}
}
