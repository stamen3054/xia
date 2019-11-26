package hu.xia.enya;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFWorkbookFactory;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;

import hu.xia.enya.model.EnyaProduct;
import hu.xia.enya.model.Order;

public class ExcelConvertor {

	private static final Logger logger = Logger.getLogger(ExcelConvertor.class.getSimpleName());

	private Workbook enyaWorkbook;
	private Workbook clientWorkbook;
	private Workbook exportWorkbook;
	private String clientFilePath = "src/main/resources/excel/client.xlsx";
	private String enyaFilePath = "src/main/resources/excel/enya.xlsx";
	private String exportFilePath = "src/main/resources/excel/export.xlsx";
	private String imagesFilePath = "src/main/resources/images";

	private List<EnyaProduct> enyaProducts;
	private Map<String, Integer> indexMap;
	private List<Order> orders;

	public ExcelConvertor() {
		logger.info("Initializing ExcelConvertor...");
		enyaProducts = new ArrayList<EnyaProduct>();
		indexMap = new HashMap<String, Integer>();
		orders = new ArrayList<Order>();
		try {
			logger.info("Trying HSSF...");
			clientWorkbook = HSSFWorkbookFactory.create(new File(clientFilePath));
			enyaWorkbook = HSSFWorkbookFactory.create(new File(enyaFilePath));
			exportWorkbook = XSSFWorkbookFactory.createWorkbook();
		} catch (Exception e) {
			logger.warning("Trying XSSF...");
			try {
				clientWorkbook = XSSFWorkbookFactory.create(new File(clientFilePath));
				enyaWorkbook = XSSFWorkbookFactory.create(new File(enyaFilePath));
				exportWorkbook = XSSFWorkbookFactory.createWorkbook();
			} catch (Exception e2) {
				logger.severe("Failed to initilize ExcelConvertor...");
				logger.severe(MessageFormat.format("Exception:[{0}]", e2.getStackTrace()));
			}
		}
	}

	private void savePicture(String filePath, PictureData pictureData) throws Exception {
		logger.info(MessageFormat.format("Saving picture to {0}", filePath));
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(new File(filePath));
			fos.write(pictureData.getData());
		} catch (Exception e) {
			logger.severe(MessageFormat.format("Failed to save picture {0}", filePath));
		} finally {
			fos.close();
		}
		logger.info(MessageFormat.format("Done saving picture {0}", filePath));
	}

	public void pullEnyaData() throws Exception {
		logger.info("Pulling Enya Data...");
		Sheet sheet = enyaWorkbook.getSheetAt(0);
		List<? extends PictureData> pictureData = enyaWorkbook.getAllPictures();
		int size = pictureData.size();
		int start = 0;
		if (sheet.getRow(0).getCell(0).getStringCellValue().equals("型号")) {
			start = 1;
			size++;
		}
		int count = 0;

		for (; start < size; start++) {
			EnyaProduct enyaProduct = new EnyaProduct();
			String fileName = sheet.getRow(start).getCell(0).getStringCellValue();
			enyaProduct.setModel(sheet.getRow(start).getCell(0).getStringCellValue());
			enyaProduct.setPrice(sheet.getRow(start).getCell(1).getNumericCellValue());
			String filePath = imagesFilePath + "/" + fileName + ".png";
			savePicture(filePath, enyaWorkbook.getAllPictures().get(start - 1));
			enyaProduct.setPicturePath(filePath);

			indexMap.put(fileName, count);
			count++;
			enyaProducts.add(enyaProduct);
		}
		logger.info("Finished pulling Enya Data...");
	}

	private void pullClientData() {
		logger.info("Pulling Client Data...");
		Sheet sheet = clientWorkbook.getSheetAt(0);
		Iterator<Row> iterator = sheet.rowIterator();
		if (sheet.getRow(0).getCell(0).getStringCellValue().equals("型号")) {
			iterator.next();
		}
		while (iterator.hasNext()) {
			Row row = iterator.next();
			String model = row.getCell(0).getStringCellValue();
			Integer index = indexMap.get(model);
			if (index == null) {
				logger.severe(MessageFormat.format("Couldn't find product {0} in Enya List", model));
			} else {
				Order order = new Order();
				order.setModel(model);
				order.setPrice(enyaProducts.get(index).getPrice());
				order.setQuantity((int) row.getCell(1).getNumericCellValue());
				order.setImagePath(enyaProducts.get(index).getPicturePath());
				orders.add(order);
			}
		}
		logger.info("Finished pulling Client Data...");
	}

	public void constructExcel() throws Exception {
		logger.info("Constructing Excel...");
		pullClientData();
		Sheet sheet = exportWorkbook.createSheet();
		int i = 0;
		int pictureIdx = 4;
		sheet.setColumnWidth(pictureIdx, 60 * 256);

		Row firstRow = sheet.createRow(0);
		firstRow.createCell(0).setCellValue("Model");
		firstRow.createCell(1).setCellValue("Price");
		firstRow.createCell(2).setCellValue("Quantity");
		firstRow.createCell(3).setCellValue("Total");
		firstRow.createCell(pictureIdx).setCellValue("Picture");
		firstRow.setHeight((short) 700);

		for (; i < orders.size(); i++) {
			Row row = sheet.createRow(i + 1);
			row.setHeight((short) 5000);
			row.createCell(0).setCellValue(orders.get(i).getModel());
			row.createCell(1).setCellValue(orders.get(i).getPrice());
			row.createCell(2).setCellValue(orders.get(i).getQuantity());
			row.createCell(3).setCellValue(orders.get(i).calculateTotalCost());
			insertPicture(exportWorkbook, sheet, i + 1, pictureIdx, orders.get(i).getImagePath());
		}

		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(new File(exportFilePath));
			exportWorkbook.write(fos);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			fos.close();
		}
		
		logger.info(MessageFormat.format("Finished Excel construction at {0}", exportFilePath));
	}

	public void insertPicture(Workbook workbook, Sheet sheet, int rowIndex, int colIndex, String imagePath)
			throws Exception {
		logger.info(MessageFormat.format("Inserting picture {0} into Excel", imagePath));
		InputStream is = new FileInputStream(imagePath);
		byte[] bytes = IOUtils.toByteArray(is);
		int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);

		CreationHelper helper = workbook.getCreationHelper();
		Drawing drawing = sheet.createDrawingPatriarch();
		ClientAnchor anchor = helper.createClientAnchor();

		anchor.setRow1(rowIndex);
		anchor.setCol1(colIndex);

		Picture picture = drawing.createPicture(anchor, pictureIdx);
		picture.resize(1, 1);
		logger.info("Finished inserting picture...");
	}
}
