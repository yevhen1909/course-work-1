const apiKey = "656965e036b7c746ed512ff08ef5e35a";
const url = "https://api.novaposhta.ua/v2.0/json/";
const ss=SpreadsheetApp.getActiveSpreadsheet();
const sheetNP=ss.getSheetByName("НП");
const sheetCities=ss.getSheetByName("Міста");
const sheetWarehouses=ss.getSheetByName("Відділення");
const sheetData=ss.getSheetByName("Дані");


function getCities() {

let data={
  "modelName": "Address",
  "calledMethod": "getCities"
}

  let options = {
    "method": "POST",
    "headers": {
      "content-type": "application/json",
      "apiKey": apiKey
    },
    "async": true,
    "crossDomain": true,
    "processData": false,
    "payload": JSON.stringify(data)
  }
  let response = UrlFetchApp.fetch(url, options);
  let dataParse = JSON.parse(response.getContentText());
  let maxCities=dataParse.data.length;

  let cities=[];
  let citiesRef=[];

  for(let i=0; i<maxCities; i++)
  {
    cities.push([dataParse.data[i].Description]);
    citiesRef.push([dataParse.data[i].Ref])
  }
  sheetCities.getRange(2,1,maxCities).setValues(cities);
  sheetCities.getRange(2,2,maxCities).setValues(citiesRef);

 console.log([dataParse.data[1]]);
}


function getWarehouseSender() {
let CityReference = sheetNP.getRange("B2").getValue();
let data={
  "modelName": "Address",
   "calledMethod": "getWarehouses",
   "methodProperties": {
   "CityRef" : CityReference
   }
}

  let options = {
    "method": "POST",
    "headers": {
      "content-type": "application/json",
      "apiKey": apiKey
    },
    "async": true,
    "crossDomain": true,
    "processData": false,
    "payload": JSON.stringify(data)
  }
  let response = UrlFetchApp.fetch(url, options);
  let dataParse = JSON.parse(response.getContentText());
  let maxCities=dataParse.data.length;

  let cities=[];
  let citiesRef=[];

  for(let i=0; i<maxCities; i++)
  {
    cities.push([dataParse.data[i].Description]);
    citiesRef.push([dataParse.data[i].Ref])
  }
 sheetWarehouses.getRange(2,1,maxCities).setValues(cities);
 sheetWarehouses.getRange(2,2,maxCities).setValues(citiesRef);
console.log(maxCities);
}

function getWarehouseReciever() {
let CityReference = sheetNP.getRange("F2").getValue();
let data={
  "modelName": "Address",
   "calledMethod": "getWarehouses",
   "methodProperties": {
   "CityRef" : CityReference
   }
}

  let options = {
    "method": "POST",
    "headers": {
      "content-type": "application/json",
      "apiKey": apiKey
    },
    "async": true,
    "crossDomain": true,
    "processData": false,
    "payload": JSON.stringify(data)
  }
  let response = UrlFetchApp.fetch(url, options);
  let dataParse = JSON.parse(response.getContentText());
  let maxCities=dataParse.data.length;

  let cities=[];
  let citiesRef=[];

  for(let i=0; i<maxCities; i++)
  {
    cities.push([dataParse.data[i].Description]);
    citiesRef.push([dataParse.data[i].Ref])
  }
 sheetWarehouses.getRange(2,4,maxCities).setValues(cities);
 sheetWarehouses.getRange(2,5,maxCities).setValues(citiesRef);
console.log(maxCities);
}

function getRecipientRef() {
sheetData.getRange("F2").setValue(" ");
let firstName = sheetNP.getRange("O2").getValue();
let middleName = sheetNP.getRange("P2").getValue();
let lastName = sheetNP.getRange("N2").getValue();
let phone = sheetNP.getRange("Q2").getValue();
let email = sheetNP.getRange("R2").getValue();
let data=
{
 "apiKey": "656965e036b7c746ed512ff08ef5e35a",
  "modelName": "Counterparty",
   "calledMethod": "save",
   "methodProperties": {
"FirstName" : firstName,
"MiddleName" : middleName,
"LastName" : lastName,
"Phone" : phone,
"Email" : email,
"CounterpartyType" : "PrivatePerson",
"CounterpartyProperty" : "Recipient"
   }
}

  let options = {
    "method": "POST",
    "headers": {
      "content-type": "application/json",
      "apiKey": apiKey
    },
    "async": true,
    "crossDomain": true,
    "processData": false,
    "payload": JSON.stringify(data)
  }
  let response = UrlFetchApp.fetch(url, options);
  let dataParse = JSON.parse(response.getContentText());

 sheetData.getRange("F2").setValue(dataParse.data[0].Ref);
}


function getContactRecipientRef() {
  
let ref= sheetData.getRange("F2").getValue();
let data=
{
 "apiKey": "656965e036b7c746ed512ff08ef5e35a",
  "modelName": "Counterparty",
   "calledMethod": "getCounterpartyContactPersons",
   "methodProperties": {
"Ref" : ref,
"Page" : "1"
   }
}


  let options = {
    "method": "POST",
    "headers": {
      "content-type": "application/json",
      "apiKey": apiKey
    },
    "async": true,
    "crossDomain": true,
    "processData": false,
    "payload": JSON.stringify(data)
  }
  let response = UrlFetchApp.fetch(url, options);
  let dataParse = JSON.parse(response.getContentText());

 
sheetData.getRange("G2").setValue(dataParse.data[0].Ref);
}



function createTTN() {
let date=sheetNP.getRange("S2").getDisplayValue();
let payerType= sheetNP.getRange("U2").getValue();
let paymentMethod= sheetNP.getRange("I2").getValue();
let volumeGeneral= sheetNP.getRange("T2").getValue();
let weight= sheetNP.getRange("J2").getValue();
let seatsAmount= sheetNP.getRange("K2").getValue();
let description= sheetNP.getRange("L2").getValue();
let cost= sheetNP.getRange("M2").getValue();
let citySender= sheetNP.getRange("B2").getValue();
let cityRecipient= sheetNP.getRange("F2").getValue();
let senderAddress= sheetNP.getRange("D2").getValue();
let recipientAddress= sheetNP.getRange("H2").getValue();
let sendersPhone= sheetData.getRange("C5").getValue();
let recipientsPhone= sheetNP.getRange("Q2").getValue();
let sender= sheetData.getRange("C2").getValue();
let contactSender= sheetData.getRange("D2").getValue();
let recipient= sheetData.getRange("F2").getValue();
let contactRecipient= sheetData.getRange("G2").getValue();
let data=
{
 "apiKey": "656965e036b7c746ed512ff08ef5e35a",
  "modelName": "InternetDocument",
   "calledMethod": "save",
   "methodProperties": {
"PayerType" : payerType,
"PaymentMethod" : paymentMethod,
"DateTime" : date,
"CargoType" : "Cargo",
"VolumeGeneral" : volumeGeneral,
"Weight" : weight,
"ServiceType" : "WarehouseWarehouse",
"SeatsAmount" : seatsAmount,
"Description" : description,
"Cost" : cost,
"CitySender" : citySender,
"Sender" : sender,
"SenderAddress" : senderAddress,
"ContactSender" : contactSender,
"SendersPhone" : sendersPhone,
"CityRecipient" : cityRecipient,
"Recipient" : recipient,
"RecipientAddress" : recipientAddress,
"ContactRecipient" : contactRecipient,
"RecipientsPhone" : recipientsPhone
   }
}

let validate= sheetNP.getRange("X2").getValue();
  let options = {
    "method": "POST",
    "headers": {
      "content-type": "application/json",
      "apiKey": apiKey
    },
    "async": true,
    "crossDomain": true,
    "processData": false,
    "payload": JSON.stringify(data)
  }
  let response = UrlFetchApp.fetch(url, options);
  let dataParse = JSON.parse(response.getContentText());

 if(validate==1)
 {
   sheetNP.getRange("Y2").setValue(dataParse.data[0].IntDocNumber);
 }
 else
 {
    sheetNP.getRange("Y2").clearContent();
 }
}
