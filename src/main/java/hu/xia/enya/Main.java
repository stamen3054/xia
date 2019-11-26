package hu.xia.enya;

public class Main {

	public static void main(String[] args) {
		// read enya data and pull picture to local
		ExcelConvertor excelConvertor = new ExcelConvertor();
		try {
			excelConvertor.pullEnyaData();
			excelConvertor.constructExcel();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
