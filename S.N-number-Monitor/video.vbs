Option Explicit
Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")
Dim strComputer, message

Dim intMonitorCount
Dim oRegistry, sBaseKey, sBaseKey2, sBaseKey3, skey, skey2, skey3
Dim sValue
dim i, iRC, iRC2, iRC3
Dim arSubKeys, arSubKeys2, arSubKeys3, arrintEDID
Dim strRawEDID
Dim ByteValue, strSerFind, strMdlFind
Dim intSerFoundAt, intMdlFoundAt, findit
Dim tmp, tmpser, tmpmdl, tmpctr
Dim batch, bHeader
batch = False

If WScript.Arguments.Count = 1 Then
strComputer = WScript.Arguments(0)
' batch = True
Else
strComputer = wshShell.ExpandEnvironmentStrings("")
'strComputer = InputBox("Check Monitor info for what PC","PC Name?",strComputer)
strComputer = "."
End If

If strcomputer = "" Then WScript.Quit
strComputer = UCase(strComputer)

If batch Then
Dim fso,logfile, appendout
logfile = wshShell.ExpandEnvironmentStrings("%userprofile%") & "\desktop\MonitorInfo.csv"

'setup Log
Const ForAppend = 8
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(logfile) Then bHeader = True
set appendout = fso.OpenTextFile(logfile, ForAppend, True)

If bHeader Then
appendout.writeline "Computer,Model,Serial #,Vendor ID,Manufacture Date,Messages"
End If
End If

Dim strarrRawEDID()
intMonitorCount=0
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
'get a handle to the WMI registry object
On Error Resume Next
Set oRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "/root/default:StdRegProv")

If Err <> 0 Then
If batch Then
EchoAndLog strComputer & ",,,,," & Err.Description
Else
MsgBox "Failed. " & Err.Description,vbCritical + vbOKOnly,strComputer
WScript.Quit
End If
End If


sBaseKey = "SYSTEM\CurrentControlSet\Enum\DISPLAY\"
'По этому пути отображаются все подключенные мониторы HKLM\SYSTEM\CurrentControlSet\Enum\DISPLAY\
iRC = oRegistry.EnumKey(HKLM, sBaseKey, arSubKeys)
	For Each sKey In arSubKeys
	'Переходим id VESA на уровень вниз
	'HKLM\SYSTEM\CurrentControlSet\Enum\DISPLAY\<VESA_Monitor_ID\
	'В ключе "HardwareID находится идентификатор железки"
	sBaseKey2 = sBaseKey & sKey & "\"
	iRC2 = oRegistry.EnumKey(HKLM, sBaseKey2, arSubKeys2)
		For Each sKey2 In arSubKeys2
		'HKLM\SYSTEM\CurrentControlSet\Enum\DISPLAY\<VESA_Monitor_ID\<PNP_ID>\
		'Проверяем значение HardvareID
		oRegistry.GetMultiStringValue HKLM, sBaseKey2 & sKey2 & "\", "HardwareID", sValue
			For tmpctr=0 to ubound(svalue)
				If lcase(left(svalue(tmpctr),8))="monitor\" then

				'Далее следует проверка на является ли монитор активным в системе
				sBaseKey3 = sBaseKey2 & sKey2 & "\"
				iRC3 = oRegistry.EnumKey(HKLM, sBaseKey3, arSubKeys3)
					For Each sKey3 In arSubKeys3
					strRawEDID = ""
						If skey3="Control" Then
						'Control sub-key Если данный ключ присутствует -то монитор является активным на данный момент в системе
						oRegistry.GetBinaryValue HKLM, sbasekey3 & "Device Parameters\", "EDID", arrintEDID

							If vartype(arrintedid) <> 8204 then
							strRawEDID="EDID Not Available"
							else
								For each bytevalue in arrintedid 'Переводим массив в строку для облегчения последующей обработки
								strRawEDID=strRawEDID & chr(bytevalue)
								Next
							End If

						'Берем строку и заносим ее в массив
						redim preserve strarrRawEDID(intMonitorCount)
						strarrRawEDID(intMonitorCount)=strRawEDID
						intMonitorCount=intMonitorCount+1
						End If
				Next
				End If
		Next
	Next
Next
'*****************************************************************************************
'Теперь вся информация о активных мониторах находится в массиве strarrRwEDID
'called arrMonitorInfo, the dimensions are as follows:
'0=VESA Mfg ID, 1=VESA Device ID, 2=MFG Date (M/YYYY),3=Serial Num (If available),4=Model Descriptor
'5=EDID Version
'*****************************************************************************************
On Error Resume Next
dim arrMonitorInfo()
redim arrMonitorInfo(intMonitorCount-1,5)
dim location(3)
for tmpctr=0 to intMonitorCount-1
If strarrRawEDID(tmpctr) <> "EDID Not Available" then
'*********************************************************************
'Теперь нам необходимо достать название модели и серийный номер (если он есть) из раздел Device Parameters ключ EDID
	'00 FF	-Адресс начала S\N
	'00 FC	-Адресс начала Modell
'Останавливаемся на адрессе H36+1 и записываем 18 байт в массив location
'Останавливаемся на адрессе H48+1 и записываем 18 байт в массив location
'Останавливаемся на адрессе H5a+1 и записываем 18 байт в массив location
'Останавливаемся на адрессе H6c+1 и записываем 18 байт в массив location
'*********************************************************************
location(0)=mid(strarrRawEDID(tmpctr),&H36+1,18)
location(1)=mid(strarrRawEDID(tmpctr),&H48+1,18)
location(2)=mid(strarrRawEDID(tmpctr),&H5a+1,18)
location(3)=mid(strarrRawEDID(tmpctr),&H6c+1,18)
'Записываем в переменную адрес начала S\N 		00 FF
strSerFind=chr(&H00) & chr(&H00) & chr(&H00) & chr(&Hff)
'Записываем в переменную адрес начала Modell	00 FC
strMdlFind=chr(&H00) & chr(&H00) & chr(&H00) & chr(&Hfc)

strSerFind strSerFind
intSerFoundAt=-1
intMdlFoundAt=-1

	for findit = 0 to 3
		If instr(location(findit),strSerFind)>0 then
		intSerFoundAt=findit
		End If
		If instr(location(findit),strMdlFind)>0 then
		intMdlFoundAt=findit
		End If
	Next
'Проверка есть -ли серийный номер
		If intSerFoundAt<>-1 then
		tmp=right(location(intSerFoundAt),14)
			If instr(tmp,chr(&H0a))>0 then
			tmpser=trim(left(tmp,instr(tmp,chr(&H0a))-1))
		Else
		tmpser=trim(tmp)'Удаление пробелов
		End If
	If left(tmpser,1)=chr(0) then tmpser=right(tmpser,len(tmpser)-1)
	else
	tmpser="Not Found"
	End If
	
msgbox tmpser'В данной переменной находится sn

If intMdlFoundAt<>-1 then
tmp=right(location(intMdlFoundAt),14)
If instr(tmp,chr(&H0a))>0 then
tmpmdl=trim(left(tmp,instr(tmp,chr(&H0a))-1))
else
tmpmdl=trim(tmp)
End If
'although it is not part of the edid spec it seems as though the
'serial number will frequently be preceeded by &H00, this
'compensates for that
If left(tmpmdl,1)=chr(0) then tmpmdl=right(tmpmdl,len(tmpmdl)-1)
else
tmpmdl="Not Found"
End If

'**************************************************************
'Next get the mfg date
'**************************************************************
Dim tmpmfgweek,tmpmfgyear,tmpmdt
'the week of manufacture is stored at EDID offset &H10
tmpmfgweek=asc(mid(strarrRawEDID(tmpctr),&H10+1,1))

'the year of manufacture is stored at EDID offset &H11
'and is the current year -1990
tmpmfgyear=(asc(mid(strarrRawEDID(tmpctr),&H11+1,1)))+1990

'store it in month/year format
tmpmdt=month(dateadd("ww",tmpmfgweek,datevalue("1/1/" & tmpmfgyear))) & "/" & tmpmfgyear

'**************************************************************
'Next get the edid version
'**************************************************************
'the version is at EDID offset &H12
Dim tmpEDIDMajorVer, tmpEDIDRev, tmpVer
tmpEDIDMajorVer=asc(mid(strarrRawEDID(tmpctr),&H12+1,1))

'the revision level is at EDID offset &H13
tmpEDIDRev=asc(mid(strarrRawEDID(tmpctr),&H13+1,1))

'store it in month/year format
tmpver=chr(48+tmpEDIDMajorVer) & "." & chr(48+tmpEDIDRev)

'**************************************************************
'Next get the mfg id
'**************************************************************
'the mfg id is 2 bytes starting at EDID offset &H08
'the id is three characters long. using 5 bits to represent
'each character. the bits are used so that 1=A 2=B etc..
'
'get the data
Dim tmpEDIDMfg, tmpMfg
dim Char1, Char2, Char3
Dim Byte1, Byte2
tmpEDIDMfg=mid(strarrRawEDID(tmpctr),&H08+1,2)
Char1=0 : Char2=0 : Char3=0
Byte1=asc(left(tmpEDIDMfg,1)) 'get the first half of the string
Byte2=asc(right(tmpEDIDMfg,1)) 'get the first half of the string
'now shift the bits
'shift the 64 bit to the 16 bit
If (Byte1 and 64) > 0 then Char1=Char1+16
'shift the 32 bit to the 8 bit
If (Byte1 and 32) > 0 then Char1=Char1+8
'etc....
If (Byte1 and 16) > 0 then Char1=Char1+4
If (Byte1 and 8) > 0 then Char1=Char1+2
If (Byte1 and 4) > 0 then Char1=Char1+1

'the 2nd character uses the 2 bit and the 1 bit of the 1st byte
If (Byte1 and 2) > 0 then Char2=Char2+16
If (Byte1 and 1) > 0 then Char2=Char2+8
'and the 128,64 and 32 bits of the 2nd byte
If (Byte2 and 128) > 0 then Char2=Char2+4
If (Byte2 and 64) > 0 then Char2=Char2+2
If (Byte2 and 32) > 0 then Char2=Char2+1

'the bits for the 3rd character don't need shifting
'we can use them as they are
Char3=Char3+(Byte2 and 16)
Char3=Char3+(Byte2 and 8)
Char3=Char3+(Byte2 and 4)
Char3=Char3+(Byte2 and 2)
Char3=Char3+(Byte2 and 1)
tmpmfg=chr(Char1+64) & chr(Char2+64) & chr(Char3+64)

'**************************************************************
'Next get the device id
'**************************************************************
'the device id is 2bytes starting at EDID offset &H0a
'the bytes are in reverse order.
'this code is not text. it is just a 2 byte code assigned
'by the manufacturer. they should be unique to a model
Dim tmpEDIDDev1, tmpEDIDDev2, tmpDev

tmpEDIDDev1=hex(asc(mid(strarrRawEDID(tmpctr),&H0a+1,1)))
tmpEDIDDev2=hex(asc(mid(strarrRawEDID(tmpctr),&H0b+1,1)))
If len(tmpEDIDDev1)=1 then tmpEDIDDev1="0" & tmpEDIDDev1
If len(tmpEDIDDev2)=1 then tmpEDIDDev2="0" & tmpEDIDDev2
tmpdev=tmpEDIDDev2 & tmpEDIDDev1

'**************************************************************
'finally store all the values into the array
'**************************************************************
'Kaplan adds code to avoid duplication...

If Not InArray(tmpser,arrMonitorInfo,3) Then
arrMonitorInfo(tmpctr,0)=tmpmfg
arrMonitorInfo(tmpctr,1)=tmpdev
arrMonitorInfo(tmpctr,2)=tmpmdt
arrMonitorInfo(tmpctr,3)=tmpser
arrMonitorInfo(tmpctr,4)=tmpmdl
arrMonitorInfo(tmpctr,5)=tmpVer
End If
End If
Next

'For now just a simple screen print will suffice for output.
'But you could take this output and write it to a database or a file
'and in that way use it for asset management.
i = 0
for tmpctr = 0 to intMonitorCount-1
If arrMonitorInfo(tmpctr,1) <> "" And arrMonitorInfo(tmpctr,0) <> "PNP" Then
If batch Then
EchoAndLog strComputer & "," & arrMonitorInfo(tmpctr,4) & "," & _
arrMonitorInfo(tmpctr,3)& "," & arrMonitorInfo(tmpctr,0) & "," & _
arrMonitorInfo(tmpctr,2)
Else
message = message & "Monitor " & chr(i+65) & ")" & VbCrLf & _
"Model Name: " & arrMonitorInfo(tmpctr,4) & VbCrLf & _
"Serial Number: " & arrMonitorInfo(tmpctr,3)& VbCrLf & _
"VESA Manufacturer ID: " & arrMonitorInfo(tmpctr,0) & VbCrLf & _
"Manufacture Date: " & arrMonitorInfo(tmpctr,2) & VbCrLf & VbCrLf
'wscript.echo ".........." & "Device ID: " & arrMonitorInfo(tmpctr,1)
'wscript.echo ".........." & "EDID Version: " & arrMonitorInfo(tmpctr,5)
i = i + 1
End If
End If
Next

If not batch Then
MsgBox message, vbInformation + vbOKOnly,strComputer & " Monitor Info"
End If

Function InArray(strValue,List,Col)
Dim i
For i = 0 to UBound(List)
If List(i,col) = cstr(strValue) Then
InArray = True
Exit Function
End If
Next
InArray = False
End Function

Sub EchoAndLog (message)
'Echo output and write to log
Wscript.Echo message
AppendOut.WriteLine message
End Sub
