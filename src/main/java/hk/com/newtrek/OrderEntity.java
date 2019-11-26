package hk.com.newtrek;

import java.util.ArrayList;
import java.util.List;

public class OrderEntity {
	private String orderNo;
	private String opcoOrderNo;
	private String status;
	private String fromInterface;
	private String opco;
	private String ksoLocation;
	private List<ShipmentEntity> shipmentEntities = new ArrayList<>();
	
	public static class ShipmentEntity {
		private String shipmentNo;
		private String opcoRequestedLrd;
		private String expectedLrd;
		private String lrdChangeReason;
		private String revisedLrd;
		private String reason;
		private List<ProductEntity> productEntities = new ArrayList<>();
		
		public String getShipmentNo() {
			return shipmentNo;
		}

		public ShipmentEntity setShipmentNo(String shipmentNo) {
			this.shipmentNo = shipmentNo;
			return this;
		}

		public String getOpcoRequestedLrd() {
			return opcoRequestedLrd;
		}

		public ShipmentEntity setOpcoRequestedLrd(String opcoRequestedLrd) {
			this.opcoRequestedLrd = opcoRequestedLrd;
			return this;
		}

		public String getExpectedLrd() {
			return expectedLrd;
		}

		public ShipmentEntity setExpectedLrd(String expectedLrd) {
			this.expectedLrd = expectedLrd;
			return this;
		}

		public String getLrdChangeReason() {
			return lrdChangeReason;
		}

		public ShipmentEntity setLrdChangeReason(String lrdChangeReason) {
			this.lrdChangeReason = lrdChangeReason;
			return this;
		}

		public String getRevisedLrd() {
			return revisedLrd;
		}

		public ShipmentEntity setRevisedLrd(String revisedLrd) {
			this.revisedLrd = revisedLrd;
			return this;
		}

		public String getReason() {
			return reason;
		}

		public ShipmentEntity setReason(String reason) {
			this.reason = reason;
			return this;
		}

		public List<ProductEntity> getProductEntities() {
			return productEntities;
		}
		
		public ShipmentEntity addProduct(ProductEntity product)
		{
			productEntities.add(product);
			return this;
		}
		
		public long getTotalProduct()
		{
			return productEntities.stream().count();
		}
	}
	
	public static class ProductEntity {
		private String productCode;
		private String opcoProductCode;
		private String description;
		private String expectedQty;
		private String revisedQty;
		public String getProductCode() {
			return productCode;
		}
		public ProductEntity setProductCode(String productCode) {
			this.productCode = productCode;
			return this;
		}
		public String getOpcoProductCode() {
			return opcoProductCode;
		}
		public ProductEntity setOpcoProductCode(String opcoProductCode) {
			this.opcoProductCode = opcoProductCode;
			return this;
		}
		public String getDescription() {
			return description;
		}
		public ProductEntity setDescription(String description) {
			this.description = description;
			return this;
		}
		public String getExpectedQty() {
			return expectedQty;
		}
		public ProductEntity setExpectedQty(String expectedQty) {
			this.expectedQty = expectedQty;
			return this;
		}
		public String getRevisedQty() {
			return revisedQty;
		}
		public ProductEntity setRevisedQty(String revisedQty) {
			this.revisedQty = revisedQty;
			return this;
		}

	}

	public String getOrderNo() {
		return orderNo;
	}

	public OrderEntity setOrderNo(String orderNo) {
		this.orderNo = orderNo;
		return this;
	}

	public String getOpcoOrderNo() {
		return opcoOrderNo;
	}

	public OrderEntity setOpcoOrderNo(String opcoOrderNo) {
		this.opcoOrderNo = opcoOrderNo;
		return this;
	}

	public String getStatus() {
		return status;
	}

	public OrderEntity setStatus(String status) {
		this.status = status;
		return this;
	}

	public String getFromInterface() {
		return fromInterface;
	}

	public OrderEntity setFromInterface(String fromInterface) {
		this.fromInterface = fromInterface;
		return this;
	}

	public String getOpco() {
		return opco;
	}

	public OrderEntity setOpco(String opco) {
		this.opco = opco;
		return this;
	}

	public String getKsoLocation() {
		return ksoLocation;
	}

	public OrderEntity setKsoLocation(String ksoLocation) {
		this.ksoLocation = ksoLocation;
		return this;
	}

	public List<ShipmentEntity> getShipmentEntities() {
		return shipmentEntities;
	}
	
	public OrderEntity addShipment(ShipmentEntity shipment)
	{
		shipmentEntities.add(shipment);
		return this;
	}
	
	public long getTotalShipment()
	{
		return shipmentEntities.stream().count();
	}
	
	public long getTotalShipmentProduct()
	{
		return shipmentEntities.stream()
				.mapToLong(shipment -> shipment.getTotalProduct())
				.sum()
				;
	}
}
