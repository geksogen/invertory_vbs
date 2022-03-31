Option Explicit
Dim WshShell
Dim strComputer
Dim intMonitorCount
Dim oRegistry, sBaseKey, sBaseKey2, sBaseKey3, skey, skey2, skey3
Dim sValue
dim iRC, iRC2, iRC3
Dim arSubKeys, arSubKeys2, arSubKeys3, arrintEDID
Dim strRawEDID
Dim ByteValue, strSerFind, strMdlFind
Dim intSerFoundAt, intMdlFoundAt, findit
Dim tmp, tmpser, tmpmdl, tmpctr
Dim batch, bHeader
Set WshShell = WScript.CreateObject("WScript.Shell")

Dim strarrRawEDID()
intMonitorCount=0
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
strComputer = "."
On Error Resume Next
Set oRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "/root/default:StdRegProv")
'**************************************************************

'**************************************************************
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

On Error Resume Next
dim arrMonitorInfo()
redim arrMonitorInfo(intMonitorCount-1,5)
dim location(3)

	For tmpctr=0 to intMonitorCount-1
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
End If

'**************************************************************
'Получаем device name
'**************************************************************
	If intMdlFoundAt<>-1 then
	tmp=right(location(intMdlFoundAt),14)
		If instr(tmp,chr(&H0a))>0 then
		tmpmdl=trim(left(tmp,instr(tmp,chr(&H0a))-1))
		else
		tmpmdl=trim(tmp)
		End If
			If left(tmpmdl,1)=chr(0) then tmpmdl=right(tmpmdl,len(tmpmdl)-1)
			else
			tmpmdl="Not Found"
		End If
'**************************************************************
'Получаем mfg id
'**************************************************************
Dim tmpEDIDMfg, tmpMfg
Dim Char1, Char2, Char3
Dim Byte1, Byte2

tmpEDIDMfg=mid(strarrRawEDID(tmpctr),&H08+1,2)
Char1=0 : Char2=0 : Char3=0
Byte1=asc(left(tmpEDIDMfg,1))
Byte2=asc(right(tmpEDIDMfg,1))
	If (Byte1 and 64) 	> 0 then Char1=Char1+16
	If (Byte1 and 32) 	> 0 then Char1=Char1+8
	If (Byte1 and 16) 	> 0 then Char1=Char1+4
	If (Byte1 and 8) 	> 0 then Char1=Char1+2
	If (Byte1 and 4) 	> 0 then Char1=Char1+1
	If (Byte1 and 2) 	> 0 then Char2=Char2+16
	If (Byte1 and 1) 	> 0 then Char2=Char2+8
	If (Byte2 and 128) 	> 0 then Char2=Char2+4
	If (Byte2 and 64) 	> 0 then Char2=Char2+2
	If (Byte2 and 32) 	> 0 then Char2=Char2+1
	
Char3=Char3+(Byte2 and 16)
Char3=Char3+(Byte2 and 8)
Char3=Char3+(Byte2 and 4)
Char3=Char3+(Byte2 and 2)
Char3=Char3+(Byte2 and 1)
tmpmfg=chr(Char1+64) & chr(Char2+64) & chr(Char3+64)
'**************************************************************



If Not InArray(tmpser,arrMonitorInfo,3) Then
arrMonitorInfo(tmpctr,0)=tmpmfg 
arrMonitorInfo(tmpctr,1)=tmpdev
arrMonitorInfo(tmpctr,2)=tmpmdt
arrMonitorInfo(tmpctr,3)=tmpser
arrMonitorInfo(tmpctr,4)=tmpmdl
arrMonitorInfo(tmpctr,5)=tmpVer
End If


msgbox arrMonitorInfo(tmpctr,0) 
msgbox arrMonitorInfo(tmpctr,4)
msgbox arrMonitorInfo(tmpctr,3)

Next	
