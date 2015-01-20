# System-Info-VBScript
Polls system information from Windows Dell computers and inserts it into a remote database.

* Author:  Patrick Karjala https://github.com/pat-trick/System-Info-VBScript
* Date:  2015/01/20
* Version:  1.2
* History:
  * 1.0 Original release
  * 1.1 Added System Tag information, changed SQL query to update duplicate keys, changed SQL table to use SystemTag as primary key
  * 1.2 Cleaned up and migrated to GitHub

## Overview

Visual Basic script to query the following parameters on a Dell Windows computer:

* Dell System Tag
* Windows Computer Name
* Dell Computer Model
* Memory Size in Slot 1
* Memory Speed in Slot 1
* Memory Size in Slot 2
* Memory Speed in Slot 2
* Total Memory in System

Directly writes to a MySQL database to the table <table_name> using the following query:

`INSERT INTO <table_name> (ComputerName, ComputerModel, MemorySize1, MemorySpeed1, MemorySize2,
      MemorySpeed2, TotalMemory) VALUES ('ComputerName', 'ComputerModel', 'MemorySize1', 'MemorySpeed1',
      'MemorySize2', 'MemorySpeed2', 'TotalMemory')
ON DUPLICATE KEY UPDATE ComputerName ='ComputerName', ComputerModel = 'ComputerModel',
      MemorySize1 = 'MemorySize1', MemorySpeed1 = 'MemorySpeed1', MemorySize2 = 'MemorySize2',
      MemorySpeed2 = 'MemorySpeed2', TotalMemory = 'TotalMemory';`

This script may be expanded upon in future iterations to query systems for
additional hardware information as necessary.

## Requirements:

Windows XP, 7, or 8.

Systems this script is run on must have the MySQL ODBC driver version 5.2.x or later installed.
This can be downloaded at http://dev.mysql.com/downloads/connector/odbc/

Remotely accessible MySQL database with a table that has the following paramaters:

* (string, PK) SystemTag
* (string) ComputerName
* (string) ComputerModel
* (int) MemorySize1
* (int) MemorySpeed1
* (int) MemorySize2
* (int) MemorySpeed2
* (int) TotalMemory

## Notes: 

All items in <> should be substituted with actual values, such as database info, passwords, etc.

PLEASE NOTE THAT THIS FILE STORES DATABASE CONNECTION INFORMATION IN THE CLEAR.
If you need to deploy this script to end user systems and are concerned about security,
please either delete the script after use or compile it into an .exe file.  There
are likely also better ways to do this, and I am not a VBScript wizard.

This script was written to originally work specifically on Dell systems with Windows XP, 7, or 8 installed.

