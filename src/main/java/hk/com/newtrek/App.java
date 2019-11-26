package hk.com.newtrek;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import hk.com.newtrek.OrderEntity.ProductEntity;
import hk.com.newtrek.OrderEntity.ShipmentEntity;

public class App {
	private static CellStyle unlockedTextCellStyle;
	private static CellStyle textCellStyle;
	private static final int rowStartIndex = 2;
	private static final DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
	private static final String password = "password";

	public static void main(String[] args) throws IOException {
		Path targetPath = Paths.get("./target/output_" + LocalDateTime.now().format(formatter) + ".xlsx");
		try (OutputStream os = Files.newOutputStream(targetPath);
				InputStream is = App.class.getClassLoader().getResourceAsStream("template.xlsx");) {
			Workbook wb = new XSSFWorkbook(is);
			DataFormat fmt = wb.createDataFormat();
			unlockedTextCellStyle = wb.createCellStyle();
			unlockedTextCellStyle.setLocked(false);
			unlockedTextCellStyle.setDataFormat(fmt.getFormat("@"));

			textCellStyle = wb.createCellStyle();
			textCellStyle.setDataFormat(fmt.getFormat("@"));

			Sheet sheet = wb.getSheet("Order");
			int rowIdx = rowStartIndex;

			List<OrderEntity> orders = initOrderList();

			for (OrderEntity order : orders) {
				final int orderSpanRow = (int) order.getTotalShipmentProduct();
				int colIdx = 0;
				int lastOrderColIdx = 5;

				Cell cell;
				for (int i = 0; i <= lastOrderColIdx; i++) {
					// merged cell region for order fields
					CellRangeAddress mergedRegion = new CellRangeAddress(rowIdx, rowIdx + orderSpanRow - 1, i, i);
					sheet.addMergedRegion(mergedRegion);
				}

				Row row = sheet.createRow(rowIdx);

				cell = createTextStyleCell(row, colIdx++);
				cell.setCellValue(order.getOrderNo());

				cell = createTextStyleCell(row, colIdx++);
				cell.setCellValue(order.getOpcoOrderNo());

				cell = createTextStyleCell(row, colIdx++);
				cell.setCellValue(order.getStatus());

				cell = createTextStyleCell(row, colIdx++);
				cell.setCellValue(order.getFromInterface());

				cell = createTextStyleCell(row, colIdx++);
				cell.setCellValue(order.getOpco());

				cell = createTextStyleCell(row, colIdx++);
				cell.setCellValue(order.getKsoLocation());

				for (ShipmentEntity shipment : order.getShipmentEntities()) {
					final int shipmentSpanRow = (int) shipment.getTotalProduct();
					int shipColIdx = colIdx;
					final int totalShipCol = 10;

					for (int i = colIdx; i < colIdx + totalShipCol; i++) {
						// merged cell region for order fields
						CellRangeAddress mergedRegion = new CellRangeAddress(rowIdx, rowIdx + shipmentSpanRow - 1, i,
								i);
						sheet.addMergedRegion(mergedRegion);
					}

					row = sheet.getRow(rowIdx);
					if (row == null) {
						row = sheet.createRow(rowIdx);
					}

					cell = createTextStyleCell(row, shipColIdx++);
					cell.setCellValue(shipment.getShipmentNo());

					cell = createTextStyleCell(row, shipColIdx++);
					cell.setCellValue(shipment.getOpcoRequestedLrd());

					cell = createTextStyleCell(row, shipColIdx++);
					cell.setCellValue(shipment.getExpectedLrd());

					cell = createTextStyleCell(row, shipColIdx++);
					cell.setCellValue(shipment.getExpectedLrd());
					if ("PROFORMA".equalsIgnoreCase(order.getStatus())) {
						cell.setCellStyle(unlockedTextCellStyle);
					}

					cell = createTextStyleCell(row, shipColIdx++);
					cell.setCellValue(shipment.getLrdChangeReason());

					cell = createTextStyleCell(row, shipColIdx++);
					cell.setCellValue(shipment.getLrdChangeReason());
					if ("PROFORMA".equalsIgnoreCase(order.getStatus())) {
						cell.setCellStyle(unlockedTextCellStyle);
					}

					cell = createTextStyleCell(row, shipColIdx++);
					cell.setCellValue(shipment.getRevisedLrd());

					cell = createTextStyleCell(row, shipColIdx++);
					cell.setCellValue(shipment.getRevisedLrd());
					if ("OFFICIAL".equalsIgnoreCase(order.getStatus())) {
						cell.setCellStyle(unlockedTextCellStyle);
					}

					cell = createTextStyleCell(row, shipColIdx++);
					cell.setCellValue(shipment.getReason());

					cell = createTextStyleCell(row, shipColIdx++);
					cell.setCellValue(shipment.getReason());
					if ("OFFICIAL".equalsIgnoreCase(order.getStatus())) {
						cell.setCellStyle(unlockedTextCellStyle);
					}

					for (ProductEntity product : shipment.getProductEntities()) {
						int productColIdx = shipColIdx;

						row = sheet.getRow(rowIdx);
						if (row == null) {
							row = sheet.createRow(rowIdx);
						}

						cell = createTextStyleCell(row, productColIdx++);
						cell.setCellValue(product.getProductCode());

						cell = createTextStyleCell(row, productColIdx++);
						cell.setCellValue(product.getOpcoProductCode());

						cell = createTextStyleCell(row, productColIdx++);
						cell.setCellValue(product.getDescription());

						cell = createTextStyleCell(row, productColIdx++);
						cell.setCellValue(product.getExpectedQty());

						cell = createTextStyleCell(row, productColIdx++);
						cell.setCellValue(product.getRevisedQty());

						cell = createTextStyleCell(row, productColIdx++);
						cell.setCellValue(product.getRevisedQty());
						if ("OFFICIAL".equalsIgnoreCase(order.getStatus())) {
							cell.setCellStyle(unlockedTextCellStyle);
						}

						rowIdx++;
					}
				}
			}

			// protected all the worksheets
			for (int i = 0; i < wb.getNumberOfSheets(); i++) {
				Sheet s = wb.getSheetAt(i);
				s.protectSheet(password);
			}

			wb.write(os);
		}

		System.out.println("... finish generating order excel: " + targetPath.toAbsolutePath().toString() +"...");
	}
	
	private static Cell createTextStyleCell(Row row, int col)
	{
		Cell cell = row.createCell(col);
		cell.setCellStyle(textCellStyle);
		return cell;
	}

	private static List<OrderEntity> initOrderList() {
		List<OrderEntity> orders = new ArrayList<>();

		OrderEntity order;
		OrderEntity.ShipmentEntity shipment;
		OrderEntity.ProductEntity product;

		// 1st order
		order = new OrderEntity();
		order.setOrderNo("PO19000046").setOpcoOrderNo("07231901").setStatus("OFFICIAL").setFromInterface("SAP EDT")
				.setOpco("B&Q plc").setKsoLocation("KSO Hong Kong");

		shipment = new OrderEntity.ShipmentEntity();
		shipment.setShipmentNo("1").setOpcoRequestedLrd("11/8/2019").setExpectedLrd("11/8/2019")
				.setLrdChangeReason("KSO - CPI - Delay LRD").setRevisedLrd("11/11/2019")
				.setReason("KSO - Late LC Issuance");

		product = new OrderEntity.ProductEntity();
		product.setProductCode("PC17000253").setOpcoProductCode("11281706").setDescription("RP Test Quotation 11281706")
				.setExpectedQty("100").setRevisedQty("");
		shipment.addProduct(product);

		product = new OrderEntity.ProductEntity();
		product.setProductCode("PC17000252").setOpcoProductCode("11281705").setDescription("RP Test Quotation 11281705")
				.setExpectedQty("100").setRevisedQty("");
		shipment.addProduct(product);

		product = new OrderEntity.ProductEntity();
		product.setProductCode("PC17000251").setOpcoProductCode("78945623").setDescription("RP TEST QUOTATION 11281704")
				.setExpectedQty("100").setRevisedQty("10");
		shipment.addProduct(product);

		order.addShipment(shipment);

		shipment = new OrderEntity.ShipmentEntity();
		shipment.setShipmentNo("2").setOpcoRequestedLrd("11/8/2019").setExpectedLrd("12/10/2019").setLrdChangeReason("")
				.setRevisedLrd("11/11/2019").setReason("KSO - Late LC Issuance");

		product = new OrderEntity.ProductEntity();
		product.setProductCode("PC17000253").setOpcoProductCode("11281706").setDescription("RP Test Quotation 11281706")
				.setExpectedQty("100").setRevisedQty("");
		shipment.addProduct(product);

		product = new OrderEntity.ProductEntity();
		product.setProductCode("PC17000252").setOpcoProductCode("11281705").setDescription("RP Test Quotation 11281705")
				.setExpectedQty("100").setRevisedQty("");
		shipment.addProduct(product);

		order.addShipment(shipment);

		orders.add(order);

		// 2nd order
		order = new OrderEntity();
		order.setOrderNo("PO19000047").setOpcoOrderNo("07231902").setStatus("PROFORMA").setFromInterface("SAP EDT")
				.setOpco("B&Q plc").setKsoLocation("KSO Hong Kong");

		shipment = new OrderEntity.ShipmentEntity();
		shipment.setShipmentNo("1").setOpcoRequestedLrd("11/8/2019").setExpectedLrd("1/7/2019")
				.setLrdChangeReason("KSO - CPI - Delay LRD").setRevisedLrd("").setReason("");

		product = new OrderEntity.ProductEntity();
		product.setProductCode("PC17000253").setOpcoProductCode("11281706").setDescription("RP Test Quotation 11281706")
				.setExpectedQty("100").setRevisedQty("");
		shipment.addProduct(product);

		product = new OrderEntity.ProductEntity();
		product.setProductCode("PC17000252").setOpcoProductCode("11281705").setDescription("RP Test Quotation 11281705")
				.setExpectedQty("100").setRevisedQty("");
		shipment.addProduct(product);

		product = new OrderEntity.ProductEntity();
		product.setProductCode("PC17000251").setOpcoProductCode("78945623").setDescription("RP TEST QUOTATION 11281704")
				.setExpectedQty("100").setRevisedQty("");
		shipment.addProduct(product);

		order.addShipment(shipment);

		orders.add(order);

		for (OrderEntity o : orders) {
			System.out.println("total shipment product count: " + o.getTotalShipmentProduct());
		}

		return orders;
	}
}
