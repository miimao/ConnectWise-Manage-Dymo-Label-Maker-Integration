# CWManageLabelMaker
Simple webapp built to automatically print out a label from a Dymo Printer when we receive items in for a ticket project or sales order.
This Project is a very bare bones solution for Procurement in ConnectWise Manage. It allows the automatic creation of labels for Items after they have been received allowing for easy tracking and organizing of inventory.

Currently the following information is provided for received items.
<ul>
  <li>Company Name</li>
  <li>Site</li>
  <li>Date of receival</li>
  <li>The Ticket/Project/SalesOrder Number</li>
  <li>PurchaseOrder Number</li>
  <li>Product Name</li>
  <li>Vendor Name</li>
</ul>

# Configuration
In the install directory there is a configuration file.
Please ensure all the information required to access your ConnectWise Manage API is provided in this file. For more information see https://developer.connectwise.com/Best_Practices/Getting_Started

Once this is complete you will need to set up a ConnectWise Callback that will alert the program when changes have been made to procurement items
Please see the following link for more information on how that is done.
https://developer.connectwise.com/Manage/Callbacks
Use the "PurchaseOrder" callback with the callback level set to "all"

# Setup
Requirements:
<ul>
  <li>Compatible Dymo Label Printer</li>
  <li>Dedicated windows machine</li>
  <li>DYMO Drivers (DLS8Setup8.7.4.exe) Newer versions WILL NOT work!</li>
  <li>Non-Sucking Service Manager</li>
</ul>

Start by making sure you have the correct Dymo Drivers installed as mentioned above

Once you have Downloaded the Provided Zip in the Release Section make a folder in you Program Files Directory and place the contents of the zip archinve in there.

Place the Executable for Non-Sucking Service Manager in the folder as well.

Now you will need to open an elevated PowerShell Prompt and run the following commands inside of the install directory.
```
.\nssm.exe install CWManageLabelMaker "C:\Program Files\CWManageLabelMaker\CWManageLabelMaker.exe"
```
```
.\nssm.exe set CWManageLabelMaker AppStderr "C:\Program Files\CWManageLabelMaker\service-error.log"
```
This will add the program as a windows service.
Inside of windows service manager please ensure the service is running and is set to start automatically on boot up.

You may need to allow whatever port you have specified in your configuration file through windows firewall.

# Building From Source

If you would like to Build this Project from source, you will need to install the required libraries via poetry and use pyinstaller.
A .spec file has been provided to make this easy
