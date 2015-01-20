' ==================================================================================
' System Info VBScript
' Filename:     system_info_vbscript.vbs
' Author:       Patrick Karjala https://github.com/pat-trick/System-Info-VBScript
' Date:         2015/01/20
' Version:      1.2
' ==================================================================================
 
 
Option Explicit
On Error Resume Next
 
Dim strComputer, objWMIService, compBios, biosObject, systemTag, compSettings, settingObject, _
        systemName, systemModel, compMemory, memoryObject, totalMemory, i, memoryCapacity(1), _
        memorySpeed(1), SQLChangestr, sqlConn, strSQLConn, sqlRS
 
 
' ------------------  Set initial rights to access registry ------------------
' Set computer name to . to indicate that the script should query this computer
strComputer = "."
 
Set objWMIService = GetObject ("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
 
 
' ------------------  Poll BIOS to get System Tag info ------------------
Set compBios = objWMIService.ExecQuery("Select * from Win32_systemenclosure")
 
' ------------------  Retrieve computer's System Tag ------------------
For Each biosObject in compBios
        systemTag = biosObject.SerialNumber
NEXT
 
 
' ------------------  Poll system data to get Computer Name ------------------
Set compSettings = objWMIService.ExecQuery("SELECT * FROM Win32_computerSystem")
 
' ------------------  Retrieve computer's Name and Model ------------------
For Each settingObject in compSettings
        systemName = settingObject.Name
        systemModel = RTrim(settingObject.Model)
Next
 
' Uncomment this to see the results of your query
' MsgBox "System Tag: " & systemTag & vbCRLF & "System Model: " & systemModel & vbCRLF & "System Name: " & systemName
 
 
' ------------------  Poll Physical RAM chips on system ------------------
Set compMemory = GetObject("winmgmts:").InstancesOf("Win32_PhysicalMemory")
i = 0
For Each memoryObject In compMemory
        memoryCapacity(i) = memoryObject.capacity / 1024 / 1024
        memorySpeed(i) = memoryObject.Speed
        ' Uncomment this to see the current memory results being queried
        ' MsgBox "Module " & i & ": " & memoryCapacity(i) & " MB" & vbCRLF & "Speed: " & memorySpeed(i)
        totalMemory = totalMemory + memoryCapacity(i)
        i = i + 1
Next
 
' Uncomment this to see the retuls of the total RAM on the system being checked
' MsgBox "Total RAM: " & totalMemory & " MB"
 
 
' ------------------ Connect to SQL Database and upload data ------------------
Set SQLConn = Wscript.CreateObject("ADODB.Connection")
 
' Windows - ODBC Driver load
strSQLConn = "Driver={MySQL ODBC 5.2w Driver};Server=<server_ip_or_domain>;Database=<database_name>;Uid=<username>;Pwd=<password>"
 
' Open SQL connection
SQLConn.Open strSQLConn
 
SQLChangestr="INSERT INTO <table_name> (SystemTag, ComputerName, ComputerModel, MemorySize1, MemorySpeed1, MemorySize2, MemorySpeed2, TotalMemory) VALUES (" & _
        chr(39) & systemTag & chr(39) & ", " & chr(39) & systemName & chr(39) & ", " & chr(39) & systemModel & chr(39) & ", " & chr(39) & memoryCapacity(0) & chr(39) & ", " & _
        chr(39) & memorySpeed(0) & chr(39) & ", " & chr(39) & memoryCapacity(1) & chr(39) & ", " & chr(39) & memorySpeed(1) & chr(39) & ", " & chr(39) & totalMemory & _
        chr(39) & ") ON DUPLICATE KEY UPDATE ComputerName = " & chr(39) & systemName & chr(39) & ", ComputerModel = " & chr(39) & systemModel & chr(39) & ", MemorySize1 = " & _
        chr(39) & memoryCapacity(0) & chr(39) & ", MemorySpeed1 = " & chr(39) & memorySpeed(0) & chr(39) & ", MemorySize2 = " & chr(39) & memoryCapacity(1) & chr(39) & _
        ", MemorySpeed2 = " & chr(39) & memorySpeed(1) & chr(39) & ", TotalMemory = " & chr(39) & totalMemory & chr(39) & ";"
       
' Uncomment this to see the string that is being submitted to the SQL database
' MsgBox SQLChangestr
 
Set sqlRS=SQLConn.Execute(SQLChangestr)
 
SQLConn.Close