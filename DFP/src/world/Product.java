package world;

import com.google.api.ads.dfp.axis.v201605.Date;

public class Product {
private double advertiserId;
private String advertiser;
private double orderId;
private String orderName;
private double adUnitId;
private String adUnitName;
private double lineItemId;
private String lineItemName;
private double deviceId;
private String deviceName;
private String trafficker;
private String startDate;
private String endDate;
private String salesperson;
private double bookedCPC;
private double bookedCPM;
private double adServerImpressions;
private double totalClicks;
private Date endOrderDate;
private double totalOrderClicks;
private double totalOrderImpressions;
private String dayly;
public Product(double advertiserId2, String advertiser, double orderId2, String orderName, double adUnitId2, String adUnitName,
double lineItemId2, String lineItemName, double deviceId2, String deviceName, String trafficker, String startDate2,
String endDate2, String salesperson, double bookedCPC2, double bookedCPM2, double adServerImpressions2,
double totalClicks2, Date endOrderDate, double totalOrderClicks, double totalOrderImpressions, String dayly) {
super();
this.advertiserId = advertiserId2;
this.advertiser = advertiser;
this.orderId = orderId2;
this.orderName = orderName;
this.adUnitId = adUnitId2;
this.adUnitName = adUnitName;
this.lineItemId = lineItemId2;
this.lineItemName = lineItemName;
this.deviceId = deviceId2;
this.deviceName = deviceName;
this.trafficker = trafficker;
this.startDate = startDate2;
this.endDate = endDate2;
this.salesperson = salesperson;
this.bookedCPC = bookedCPC2;
this.bookedCPM = bookedCPM2;
this.adServerImpressions = adServerImpressions2;
this.totalClicks = totalClicks2;
this.endOrderDate = endOrderDate;
this.totalOrderClicks = totalOrderClicks;
this.totalOrderImpressions = totalOrderImpressions;
this.dayly = dayly;
}




public Date getEndOrderDate() {
return endOrderDate;
}




public void setEndOrderDate(Date endOrderDate) {
this.endOrderDate = endOrderDate;
}




public double getTotalOrderClicks() {
return totalOrderClicks;
}


public String getDayly()
{
return dayly;
}

public void setTotalOrderClicks(double totalOrderClicks) {
this.totalOrderClicks = totalOrderClicks;
}




public double getTotalOrderImpressions() {
return totalOrderImpressions;
}




public void setTotalOrderImpressions(double totalOrderImpressions) {
this.totalOrderImpressions = totalOrderImpressions;
}




public double getAdvertiserId() {
return advertiserId;
}




public void setAdvertiserId(int advertiserId) {
this.advertiserId = advertiserId;
}




public String getAdvertiser() {
return advertiser;
}




public void setAdvertiser(String advertiser) {
this.advertiser = advertiser;
}




public double getOrderId() {
return orderId;
}




public void setOrderId(int orderId) {
this.orderId = orderId;
}




public String getOrderName() {
return orderName;
}




public void setOrderName(String orderName) {
this.orderName = orderName;
}




public double getAdUnitId() {
return adUnitId;
}




public void setAdUnitId(int adUnitId) {
this.adUnitId = adUnitId;
}




public String getAdUnitName() {
return adUnitName;
}




public void setAdUnitName(String adUnitName) {
this.adUnitName = adUnitName;
}




public double getLineItemId() {
return lineItemId;
}




public void setLineItemId(int lineItemId) {
this.lineItemId = lineItemId;
}




public String getLineItemName() {
return lineItemName;
}




public void setLineItemName(String lineItemName) {
this.lineItemName = lineItemName;
}




public double getDeviceId() {
return deviceId;
}




public void setDeviceId(int deviceId) {
this.deviceId = deviceId;
}




public String getDeviceName() {
return deviceName;
}




public void setDeviceName(String deviceName) {
this.deviceName = deviceName;
}




public String getTrafficker() {
return trafficker;
}




public void setTrafficker(String trafficker) {
this.trafficker = trafficker;
}




public String getStartDate() {
return startDate;
}




public void setStartDate(String  startDate) {
this.startDate = startDate;
}




public String getEndDate() {
return endDate;
}




public void setEndDate(String endDate) {
this.endDate = endDate;
}




public String getSalesperson() {
return salesperson;
}




public void setSalesperson(String salesperson) {
this.salesperson = salesperson;
}



public double getBookedCPC() {
return bookedCPC;
}




public void setBookedCPC(int bookedCPC) {
this.bookedCPC = bookedCPC;
}




public double getBookedCPM() {
return bookedCPM;
}




public void setBookedCPM(int bookedCPM) {
this.bookedCPM = bookedCPM;
}




public double getAdServerImpressions() {
return adServerImpressions;
}




public void setAdServerImpressions(int adServerImpressions) {
this.adServerImpressions = adServerImpressions;
}




public double getTotalClicks() {
return totalClicks;
}




public void setTotalClicks(int totalClicks) {
this.totalClicks = totalClicks;
}
public String toString()
{
return this.advertiser + "/" + this.adServerImpressions + "/" + this.orderName;
}

}
