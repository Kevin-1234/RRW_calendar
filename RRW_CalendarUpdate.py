import requests #sending soap request
import gzip #decompress the soap response
import xml.etree.ElementTree as ET #formatting data
import time #time interval for sending request
import datetime
#creating event on google calendar through google calendar API
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
import pickle
import pytz
from datetime import datetime, timedelta


class Orders():
    def __init__(self):
        self.orderId = ''
        self.orderStatus = ''
        self.billName = ''
        self.billPhone = ''
        self.deliveryDueDate = ''
        self.dateCreated = ''
        self.customerIds = []
        #number of orders with no delivery due date
        self.noDDD = 0
        #number of orders uploaded successfully
        self.success = 0
        self.timezone ="Australia/Sydney"
        self.currentTime = pytz.timezone(self.timezone).localize(datetime.now())

    def customerIdRquesting(self):
        self.customerIds = []
        customerId=''
        #select customers that are updated 3 days before current time
        lastUpdated = (self.currentTime - timedelta(days=3)).isoformat()
        url="https://v2wsisandbox.retailexpress.com.au/DOTNET/Admin/WebServices/v2/Webstore/Service.asmx"
        headers = {'Accept': 'text/html', 'content-Type': 'text/xml'}
        body =f"""<?xml version="1.0" encoding="UTF-8"?>
            <soap:Envelope xmlns:soap="http://www.w3.org/2003/05/soap-envelope" xmlns:ret="http://retailexpress.com.au/">
           <soap:Header>
              <ret:ClientHeader>
                 <!--Optional:-->
                 <ret:ClientID>*************</ret:ClientID>
                 <!--Optional:-->
                 <ret:UserName>*******</ret:UserName>
                 <!--Optional:-->
                 <ret:Password>**********</ret:Password>
              </ret:ClientHeader>
           </soap:Header>
           <soap:Body>
              <ret:CustomerGetBulkDetails>
                 <ret:LastUpdated>{lastUpdated}</ret:LastUpdated>
                 <ret:OnlyCustomersWithEmails>0</ret:OnlyCustomersWithEmails>
                 <ret:OnlyCustomersForExport>0</ret:OnlyCustomersForExport>
              </ret:CustomerGetBulkDetails>
           </soap:Body>
        </soap:Envelope>"""
        response = requests.post(url,data=body,headers=headers)

        #decompress the response (.xml.gz)
        data = gzip.decompress(response.content)
        #formatting and sorting data
        data=str(data, 'utf-8')
        root = ET.fromstring(data)
        for Customer in root[0].findall('Customer'):
            customerId = Customer.find('CustomerId').text
            self.customerIds.append(customerId)


    def orderInfoRequesting(self):
        response = ''
        data = ''
        root = ''
        for index in range(len(self.customerIds)):
            url="https://v2wsisandbox.retailexpress.com.au/DOTNET/Admin/WebServices/v2/Webstore/Service.asmx"
            headers = {'Accept': 'text/html', 'content-Type': 'text/xml'}
            body =f"""<?xml version="1.0" encoding="UTF-8"?>
                <soap:Envelope xmlns:soap="http://www.w3.org/2003/05/soap-envelope" xmlns:ret="http://retailexpress.com.au/">
               <soap:Header>
                  <ret:ClientHeader>
                     <!--Optional:-->
                     <ret:ClientID>***************</ret:ClientID>
                     <!--Optional:-->
                     <ret:UserName>*********</ret:UserName>
                     <!--Optional:-->
                     <ret:Password>*********</ret:Password>

                  </ret:ClientHeader>
               </soap:Header>
               <soap:Body>
                  <ret:OrdersGetHistoryByChannel>
                    <ret:CustomerId>{self.customerIds[index]}</ret:CustomerId>
                    <ret:WebOrdersOnly>0</ret:WebOrdersOnly>
                    <ret:ChannelId>1</ret:ChannelId>
                  </ret:OrdersGetHistoryByChannel>
               </soap:Body>
            </soap:Envelope>"""

            response = requests.post(url,data=body,headers=headers)

            #decompress the response (.xml.gz)
            data=str(response.content, 'utf-8')
            root = ET.fromstring(data)

            for Order in root[0][0][0][0].findall('Order'):
                self.orderStatus = Order.find('OrderStatus').text
                if self.orderStatus == 'Awaiting Payment':

                    SameId = 0
                    self.orderId = Order.find('OrderId').text
                    self.billName = Order.find('BillName').text
                    self.billPhone = Order.find('BillPhone').text
                    self.dateCreated = Order.find('DateCreated').text
                    #select orders that are created 3 days before current time
                    try:
                        phrasedTime = datetime.strptime(self.dateCreated, "%Y-%m-%dT%H:%M:%S.%f%z")
                    except:
                        phrasedTime = datetime.strptime(self.dateCreated, "%Y-%m-%dT%H:%M:%S%f%z")

                    validDate = (self.currentTime - timedelta(days=3)) <= phrasedTime
                    if validDate:
                        if root[0][0][0][0].findall('OrderDetail') != []:
                            for OrderDetail in root[0][0][0][0].findall('OrderDetail'):
                                orderId2 = OrderDetail.find('OrderId').text

                                if orderId2 == self.orderId:
                                    SameId+=1
                                    deliveryDueDate = OrderDetail.find('DeliveryDueDate')
                                    if deliveryDueDate == None:
                                        self.deliveryDueDate = 'None'
                                        self.noDDD += 1
                                        print('CustomerId:', self.customerIds[index])
                                        print ('OrderId:', self.orderId, 'DateCreated:', self.dateCreated, 'OrderStatus:', self.orderStatus, 'BillName:', self.billName, 'BillPhone:', self.billPhone, 'DeliveryDueDate: ',self.deliveryDueDate, "(Uploading failed! No delivery due date is found!)\n")

                                    else:
                                        self.deliveryDueDate = deliveryDueDate.text
                                        self.orderUploading()
                                        self.success += 1
                                        print('CustomerId:', self.customerIds[index])
                                        print ('OrderId:', self.orderId, 'DateCreated:', self.dateCreated, 'OrderStatus:', self.orderStatus, 'BillName:', self.billName, 'BillPhone:', self.billPhone, 'DeliveryDueDate:', self.deliveryDueDate, "(Order has been uploaded to the calendar!)\n")

                                break
                            if SameId==0:
                                self.noDDD += 1
                                print('CustomerId:', self.customerIds[index])
                                print ('OrderId:', self.orderId, 'DateCreated:', self.dateCreated, 'OrderStatus:', self.orderStatus, 'BillName:', self.billName, 'BillPhone:', self.billPhone, 'DeliveryDueDate: None', "(Uploading failed! No delivery due date is found!)\n")
            index += 1


    def orderUploading(self):
        #communicating with google calendar
        scopes = ['https://www.googleapis.com/auth/calendar']
        flow = InstalledAppFlow.from_client_secrets_file("client_secret.json", scopes=scopes)
        credentials = pickle.load(open("token.pkl", "rb"))
        service = build("calendar", "v3", credentials=credentials)
        result = service.calendarList().list().execute()
        calendar_id = result['items'][0]['id']
        description = f"""orderStatus:{self.orderStatus}\nbillName:{self.billName}\nbillPhone:{self.billPhone}"""
        events_result = service.events().list(calendarId='primary', timeMin=self.deliveryDueDate, orderBy='updated').execute()
        events = events_result.get('items', [])
        existedOrder = False
        #check if the order already existed in the calendar
        for event in events:
            if event['summary'] == self.orderId:
                existedOrder = True
                break

        if not existedOrder:
            event = {
              'summary': self.orderId,
              'OrderStatus': self.orderStatus,
              'description': description,
              'BillPhone': self.billPhone,
              'start': {
                'dateTime': self.deliveryDueDate,
                'timeZone':self.timezone,
              },
              'end': {
                'dateTime': self.deliveryDueDate,
                'timeZone': self.timezone,
              },
              'reminders': {
                'useDefault': False,
                'overrides': [
                  {'method': 'email', 'minutes': 24 * 60},
                  {'method': 'popup', 'minutes': 10},
                ],
              },
            }

            service.events().insert(calendarId=calendar_id, body=event).execute()




#sending request every 120 minutes
nexttime = time.time()
order = Orders()
while True:

    try:
        order.customerIdRquesting()
    except:
        print("\nThe API interface can only be accessed every 60 minutes, please try again later!\n")
        userInput = input("Press any key to close the window....")
        break
    print("\n\nUploading orders created in the last 3 days.....(Please do not close the window!)")
    print("Note: Only orders with the status of 'awaiting payment' and a valid delivery due date will be uploaded.\n\n")
    order.orderInfoRequesting()
    print("\nUploading process is done!")
    print(order.noDDD," orders were found with no delivery due date!")
    print(order.success, " orders were uploaded successfully!")
    userInput = input("Press any key to close the window or keep it open to refresh every 120 minutes...")
    break
    nexttime += 7200
    sleeptime = nexttime - time.time()
    if sleeptime > 0:
        time.sleep(sleeptime)
