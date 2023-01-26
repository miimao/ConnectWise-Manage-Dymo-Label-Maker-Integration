import requests
import base64
import json
from win32com.client import Dispatch
import sys
import time
from datetime import datetime
from datetime import timedelta
from datetime import date
import pythoncom
import logging
import tomllib
import pathlib

if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
    # running as bundle (aka frozen)
    exe_dir = pathlib.Path(sys.executable).parent
else:
    # running live
    exe_dir = pathlib.Path(__file__).parent

config_path = pathlib.Path.cwd() / exe_dir / "config.toml"
with config_path.open(mode="rb") as fp:
    config = tomllib.load(fp)

datetime_format = "%Y-%m-%dT%H:%M:%SZ"  # This is used to translate dates from CWManage API into datetime objects

product_ids_already_printed = []


def generate_cw_token(Company_ID, Public_Key, Private_Key):
    token = "{}+{}:{}".format(Company_ID, Public_Key, Private_Key)
    token = base64.b64encode(bytes(token, "utf-8"))
    token = token.decode("utf-8")
    return token


cw_manage_url = config["CONNECTWISEAPI"]["URL"]
company_id = config["CONNECTWISEAPI"]["COMPANYID"]
cw_manage_public = config["CONNECTWISEAPI"]["PUBLIC"]
cw_manage_private = config["CONNECTWISEAPI"]["PRIVATE"]
client_id = config["CONNECTWISEAPI"]["CLIENTID"]
cw_token = generate_cw_token(company_id, cw_manage_public, cw_manage_private)
headers_cw = {
    "Authorization": "Basic " + cw_token,
    "clientId": client_id,
    "Accept": "*/*",
    "Content-Type": "application/json",
}


def get_purchase_order_items(purchase_order_id):
    url = cw_manage_url + f"/procurement/purchaseorders/{purchase_order_id}/lineitems"
    params = {
        "pageSize": "1000",
    }
    items = requests.get(headers=headers_cw, url=url, params=params).json()

    logging.debug(f"Gathering Items for PurchaseOrder: {purchase_order_id}\n{items}")
    return items


def get_open_purchase_orders():
    params = {
        "conditions": "closedFlag=false",
        "pageSize": "1000",
    }
    url = cw_manage_url + f"/procurement/purchaseorders/"

    items = requests.get(headers=headers_cw, url=url, params=params).json()

    return items


def nearest(items, pivot):
    return min(items, key=lambda x: abs(x - pivot))

def get_client_site_from_product(product):
    if "ticket" in product:
        ticket = requests.get(headers=headers_cw, url=product['ticket']['_info']['ticket_href'],params={'fields':'site'}).json()
        logging.debug(f"Ticket Site Info\n{ticket}")
        return ticket['site']['name']
    elif "project" in product:
        project = requests.get(headers=headers_cw, url=product['project']['_info']['project_href'],params={'fields':'site'}).json()
        logging.debug(f"Project Site Info\n{project}")
        return project['site']['name']
    elif "salesOrder" in product:
        sales_order = requests.get(headers=headers_cw, url=product['salesOrder']['_info']['so_href'],params={'fields':'site'}).json()
        logging.debug(f"SalesOrder Site Info\n{sales_order}")
        return sales_order['site']['name']
    else:
        return None


def find_purchase_order_origin(purchaseorder_lineitem):
    logging.debug(f"Attempting to find origin of Line Item from purchase order")
    logging.debug(f"LineItem:\n{purchaseorder_lineitem}")
    # Because the connectwise api does not have anyway to tie a purchase order back to its origins
    # we will have to make a guess as to what it is using the information we can gather from the Purchase Order
    params = {
        "conditions": f"catalogItem/id={purchaseorder_lineitem['product']['id']}",
        "pageSize": "1000",
    }
    url = cw_manage_url + "/procurement/products/"

    products = requests.get(headers=headers_cw, url=url, params=params).json()

    if len(products) == 1:
        logging.debug(
            f"Only found one product:{products[0]['id']} with CatalogId:{purchaseorder_lineitem['product']['id']}. Using this product"
        )
        return products[0]

    purchaseorder_lineitem_entered_date = datetime.strptime(
        purchaseorder_lineitem["_info"]["dateEntered"], datetime_format
    )
    purchase_datetime = datetime.strptime(
        purchaseorder_lineitem["_info"]["lastUpdated"], datetime_format
    )
    product_dates = []
    for item in products:
        product_dates.append(
            datetime.strptime(item["_info"]["lastUpdated"], datetime_format)
        )
    logging.debug(
        f"Found multiple Products ({len(product_dates)}) with CatalogId:{purchaseorder_lineitem['product']['id']}"
    )
    nearest_date_found = nearest(product_dates, purchase_datetime)
    logging.debug(f"Attepting to find using lastUpdated")
    logging.debug(
        f"PurchaseOrder item lastUpdate: {purchase_datetime}, Closest date from found products is {nearest_date_found}"
    )
    if ((purchase_datetime - nearest_date_found).total_seconds()) < 10 and (
        (purchase_datetime - nearest_date_found).total_seconds()
    ) > -10:
        for item in products:
            if item["_info"]["lastUpdated"] == nearest_date_found.strftime(
                datetime_format
            ):
                logging.debug(
                    f"Found product:{item['id']} with CatalogId:{purchaseorder_lineitem['product']['id']} with [lastUpdated] within + or - 10 sec of [recieve date]"
                )
                return item
    logging.debug(f"Unable to find product based on LastUpdate")

    logging.debug(f"Attepting to find using dateEntered")
    nearest_date_found = nearest(product_dates, purchaseorder_lineitem_entered_date)
    logging.debug(
        f"PurchaseOrder item dateEntered: {purchaseorder_lineitem_entered_date}, Closest date from found products is {nearest_date_found}"
    )
    if ((purchaseorder_lineitem_entered_date - nearest_date_found)) < timedelta(
        days=1
    ) and ((purchaseorder_lineitem_entered_date - nearest_date_found)) > timedelta(
        days=-1
    ):
        logging.debug(
            f"Found product:{item['id']} with CatalogId:{purchaseorder_lineitem['product']['id']} with [lastUpdate] within + or - 1 day of PurchaseOrder Item [dateEntered]"
        )
        for product in products:
            if product["_info"]["lastUpdated"] == nearest_date_found.strftime(
                datetime_format
            ):
                return product
    logging.debug("Unable to find product based on dateEntered")

    logging.debug(
        f"Attepting to find using purchaseDate if avalaible (Not every product will have a purchase date)"
    )
    product_dates = []
    for item in products:
        if "purchaseDate" in item:
            product_dates.append(
                datetime.strptime(item["purchaseDate"], datetime_format)
            )
    nearest_date_found = nearest(product_dates, purchaseorder_lineitem_entered_date)
    logging.debug(
        f"PurchaseOrder PurchaseDate: {purchaseorder_lineitem_entered_date}, Closest date from found products is {nearest_date_found}"
    )
    if ((purchaseorder_lineitem_entered_date - nearest_date_found)) < timedelta(
        days=7
    ) and ((purchaseorder_lineitem_entered_date - nearest_date_found)) > timedelta(
        days=-7
    ):
        logging.debug(
            f"Found product:{item['id']} with CatalogId:{purchaseorder_lineitem['product']['id']} with [purchaseDate] within + or - 7 days of Creation date"
        )
        for product in products:
            if product["purchaseDate"] == nearest_date_found.strftime(datetime_format):
                return product
    logging.debug("Unable to find product based on [purchaseDate]")

    logging.debug(
        f"Cound not find product with CatalogId:{purchaseorder_lineitem['product']['id']} within + or - 10 sec of recieve date"
    )
    return None


def generate_lable(webhook_entity, item):
    printer_name = "DYMO LabelWriter 4XL"
    printer_com = Dispatch("Dymo.DymoAddIn", pythoncom.CoInitialize())
    printer_com.SelectPrinter(printer_name)
    label = Dispatch("Dymo.DymoLabels")
    label_template = pathlib.Path.cwd() / exe_dir / "template.label"
    isOpen = printer_com.Open(label_template)
    logging.debug(
        f"Print Info - PrinterName:{printer_name}, PrinterCom:{printer_com}, LabelTemplate:{label_template}"
    )
    purchaseorder_origin_info = find_purchase_order_origin(item)
    logging.debug(f"Product origin:\n{purchaseorder_origin_info}")
    if purchaseorder_origin_info != None:
        logging.debug(
            f"Product:{purchaseorder_origin_info['id']} was found adding infomation from product to label."
        )
        #set product name
        label.SetField("ProductName",f"{purchaseorder_origin_info['catalogItem']['identifier']}")

        #Set the Site of the Client
        label.SetField("ClientSite", get_client_site_from_product(purchaseorder_origin_info))

        if "company" in purchaseorder_origin_info:
            label.SetField(
                "ClientName", f'{purchaseorder_origin_info["company"]["name"]}'
            )
        else:
            label.SetField("ClientName", "Community Technology Services")
        if "ticket" in purchaseorder_origin_info:
            label.SetField(
                "TicketProjectSalesOrderNumber",
                f'Ticket Number: {purchaseorder_origin_info["ticket"]["id"]}',
            )
        elif "project" in purchaseorder_origin_info:
            label.SetField(
                "TicketProjectSalesOrderNumber",
                f'Project Number: {purchaseorder_origin_info["project"]["id"]}',
            )
        elif "salesOrder" in purchaseorder_origin_info:
            label.SetField(
                "TicketProjectSalesOrderNumber",
                f'Sales Order Number: {purchaseorder_origin_info["salesOrder"]["id"]}',
            )

    else:  # If we recieve no info from find_purchase_order_origin function set it to these defualts values
        label.SetField("ClientName", "Community Technology Services")
        label.SetField("ClientSite", "Community Technology Services")
        label.SetField("TicketProjectSalesOrderNumber", "N/A")
        label.SetField ('ProductName', item['product']['identifier'])
    label.SetField("PurchaseOrderNumber", item["packingSlip"])
    label.SetField(
        "VendorName", f'Vendor Name: {webhook_entity["vendorCompany"]["name"]}'
    )
    

    global product_ids_already_printed

    if (
        purchaseorder_origin_info == None
        or purchaseorder_origin_info["id"] not in product_ids_already_printed
    ):
        logging.info(
            f"""
        Printing Label/s with the folloing info:
        Client Name: {label.GetText("ClientName")}
        Client Site: {label.GetText("ClientSite")}
        Date Received: {date.today()}
        Product Name: {label.GetText("ProductName")}
        Ticket, Project, SalesOrder, Number: {label.GetText("TicketProjectSalesOrderNumber")}
        PurchaseOrder Number: {label.GetText("PurchaseOrderNumber")}
        Vendor Name: {label.GetText("VendorName")}
        """
        )
        if purchaseorder_origin_info != None:
            product_ids_already_printed.append(purchaseorder_origin_info["id"])
            #Get Quantitiy modifier if set (this is a custom field used to correct items that are sold by the unit but sold in collective sets Ex: 1000 ft of cable is sold by the foot but come in a roll of 1000ft so quanitity will be set to 1000)
            quantity_modifier = requests.get(url=purchaseorder_origin_info['catalogItem']["_info"]["catalog_href"], headers=headers_cw, params={'fields':'customFields/value','customFieldConditions':'caption="Lot Size" AND value!=0'}).json()
            print(bool(quantity_modifier))
            if len(quantity_modifier) != 0:
                quantity_modifier = int(quantity_modifier['customFields'][0]['value'])
                quantity_to_print = purchaseorder_origin_info['quantity']/quantity_modifier
            else:
                quantity_modifier = None
                quantity_to_print = int(purchaseorder_origin_info['quantity'])
            if quantity_to_print % 1 != 0:
                logging.warning(f"The quantity modifier has resulted in a non whole number: {quantity_to_print}. This should be corrected in ConnectWise manage. It will be automatically round to the nearest whole number.")
            quantity_to_print = round(quantity_to_print)
            logging.info(
                f"Quantity is set to: '{purchaseorder_origin_info['quantity']}', Quantity Modifier is set to: '{quantity_modifier}'. Printing '{quantity_to_print}' labels"
            )
            printer_com.StartPrintJob()
            printer_com.Print(quantity_to_print, False)
            printer_com.EndPrintJob()
            time.sleep(1)
        else:
            printer_com.StartPrintJob()
            printer_com.Print(1, False)
            printer_com.EndPrintJob()
            time.sleep(1)
    else:
        logging.info(
            f"Product ID:{purchaseorder_origin_info['id']} Has already had a label printed for it. Skipping"
        )


def proccess_request(webhook_json_data):
    webhook_entity_json_data = json.loads(webhook_json_data["Entity"])
    logging.debug(f"Entity json data from CW CallBack:\n{webhook_entity_json_data}")
    purchase_order_items = get_purchase_order_items(webhook_entity_json_data["id"])

    # get the item from the purchase order item list that is closest to the time stamp of the original webhook and make sure it has the statues "receivedStatus": "FullyReceived"
    webhook_datetime = datetime.strptime(
        webhook_entity_json_data["_info"]["lastUpdated"], datetime_format
    )
    logging.debug(f"CW CallBack Entity 'lastUpdated' DateTime: {webhook_datetime}")

    purchase_order_item_dates = []

    for item in purchase_order_items:
        purchase_order_item_dates.append(
            datetime.strptime(item["_info"]["lastUpdated"], datetime_format)
        )

    nearest_date_found = nearest(purchase_order_item_dates, webhook_datetime)
    logging.debug(
        f"From items assocated with PurchaseOrder closest match to Entity 'lastUpdated' is: {nearest_date_found}"
    )

    for item in purchase_order_items:
        if (
            item["_info"]["lastUpdated"] == nearest_date_found.strftime(datetime_format)
            and item["receivedStatus"] == "FullyReceived"
        ):
            logging.debug(
                f"Item:{item['id']} from PurchaseOrder:{webhook_entity_json_data['id']} is the closest match and is marked as 'FullyReceived'"
            )
            if ((webhook_datetime - nearest_date_found).total_seconds()) < 10 and (
                (webhook_datetime - nearest_date_found).total_seconds()
            ) > -10:
                logging.debug(
                    f"Item:{item['id']} from PurchaseOrder:{webhook_entity_json_data['id']} IS within + or - 10 sec of CallBack date. Adding to label Queue"
                )
                logging.debug(f"Item Found:\n{item}")
                generate_lable(webhook_entity_json_data, item)
            else:
                logging.debug(
                    f"Item:{item['id']} from PurchaseOrder:{webhook_entity_json_data['id']} NOT within + or - 10 sec of CallBack date. Skipping"
                )
