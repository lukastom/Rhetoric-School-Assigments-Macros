REM  *****  BASIC  *****

sub Posli_email_VIDEO

	'získat datum
	datum = InputBox ("Zadaj dátum, ku ktorému chceš poslať zadanie VIDEA e-mailom (vo formátu dd.mm.yyyy).")
	Doc = ThisComponent
	
	'zapsat ho do S1
	Sheets = Doc.getSheets  'get the collection of Sheets
	Sheet = Sheets.GetByName("Rozvrh")  'get the Sheet
	Cell = Sheet.getCellByPosition(18,0)  'Get the Cell (S1)
	Cell.String = datum
	
	'skoč do listu "Rozvrh"
	Doc.CurrentController.ActiveSheet = Doc.Sheets.getByName("Rozvrh")

	'Získat číslo v buňce S2, dát jí do proměnné "cislo_riadku"
	cislo_riadku = ThisComponent.CurrentController.ActiveSheet.getCellRangeByName("S2").Value
	
	'Zjisti jmeno_1
	jmeno_1 = ThisComponent.CurrentController.ActiveSheet.getCellRangeByName("F" & cislo_riadku).String
	
	'Zjisti jmeno_2
	jmeno_2 = ThisComponent.CurrentController.ActiveSheet.getCellRangeByName("G" & cislo_riadku).String

	'Zapis jmeno_1 do U1
	Cell = Sheet.getCellByPosition(20,0)  'Get the Cell (U1)
	Cell.String = jmeno_1
	
	'Zjisti email_1
	email_1 = ThisComponent.CurrentController.ActiveSheet.getCellRangeByName("U2").String	
	
	'Zapis jmeno_2 do U1
	Cell = Sheet.getCellByPosition(20,0)  'Get the Cell (U1)
	Cell.String = jmeno_2
	
	'Zjisti email_2
	email_2 = ThisComponent.CurrentController.ActiveSheet.getCellRangeByName("U2").String	
	
	'Posli e-mail (funguje jen v Libreoffice, nefunguje v OpenOffice)

   	eMailer = createUNOService("com.sun.star.system.SimpleCommandMail")
   	eMailClient = eMailer.querySimpleMailClient()
   	eMessage = eMailer.createSimpleMailMessage()

   	eMessage.Recipient = email_1 & "," & email_2
   	eMessage.Subject = "Zhromaždenie ŽaS " & datum & " – zadanie predvedenia"
   	eMessage.Body = "Ahojte," & chr(10) & chr(10) & "rád by som Vás požiadal, abyste si pripravili predvedenie namiesto VIDEA na zhromaždenie Život a služba " & datum & "." & chr(10) & chr(10) & "• Úvod: " & jmeno_1 & chr(10) & chr(10) & "• Partnerka: " & jmeno_2 & chr(10) & chr(10) & "Prosím, pripravte si úvod na literatúru, ktorú ponúkame príslušný mesiac (na základe dokumentu Ponuka literatúry v SPJ - LitOff17-VSL)." & chr(10) & chr(10) & "Ďakujem za spoluprácu," & chr(10) & chr(10) & "Lukáš"
   	eMailer.sendSimpleMailMessage (eMessage, com.sun.star.system.SimpleMailClientFlags.NO_USER_INTERFACE)
   	
   	rem - define variables
	dim document   as object
	dim dispatcher as object
	
	rem - get access to the document
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	'dát kurzor na F a tu proměnnou "cislo_riadku"
	dim args1(0) as new com.sun.star.beans.PropertyValue
	args1(0).Name = "ToPoint"
	args1(0).Value = "$F$" & cislo_riadku	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args1())
	
	'odboldovat
	dim args2(0) as new com.sun.star.beans.PropertyValue
	args2(0).Name = "Bold"
	args2(0).Value = false	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args2())
	
	'dát kurzor na G a tu proměnnou "cislo_riadku"
	dim args3(0) as new com.sun.star.beans.PropertyValue
	args3(0).Name = "ToPoint"
	args3(0).Value = "$G$" & cislo_riadku	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args3())
	
	'odboldovat
	dim args4(0) as new com.sun.star.beans.PropertyValue
	args4(0).Name = "Bold"
	args4(0).Value = false	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args4())
	
	'dát kurzor na D a tu proměnnou "cislo_riadku"
	dim args5(0) as new com.sun.star.beans.PropertyValue
	args5(0).Name = "ToPoint"
	args5(0).Value = "$D$" & cislo_riadku	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args5())

end sub

sub Ukaz_deadliny
	
	'získej přístup k dokumentu
	dim document   as object
	dim dispatcher as object
	Doc = ThisComponent
	document   = Doc.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	'odkryj list
	dim args1(0) as new com.sun.star.beans.PropertyValue
	args1(0).Name = "aTableName"
	args1(0).Value = "(Deadliny)"
	dispatcher.executeDispatch(document, ".uno:Show", "", 0, args1())
	
	'skoč do listu "(Deadliny)"
	Doc.CurrentController.ActiveSheet = Doc.Sheets.getByName("(Deadliny)")
	
	'Získat věty a dát je do proměnných
	prvni_veta = ThisComponent.CurrentController.ActiveSheet.getCellRangeByName("B14").String
	druha_veta = ThisComponent.CurrentController.ActiveSheet.getCellRangeByName("B26").String
	treti_veta = ThisComponent.CurrentController.ActiveSheet.getCellRangeByName("B40").String
	
	If prvni_veta = "" Then prvni_veta = "" Else prvni_veta = prvni_veta & chr(10)
	If treti_veta = "" Then treti_veta = "" Else treti_veta = chr(10) & treti_veta
	
	MsgBox(prvni_veta & druha_veta & treti_veta,48,"Deadliny")
	
	'zakryj list
	dim args2(0) as new com.sun.star.beans.PropertyValue
	args2(0).Name = "aTableName"
	args2(0).Value = "(Deadliny)"
	dispatcher.executeDispatch(document, ".uno:Hide", "", 0, args2())
	
	'skoč do listu "Rozvrh"
	Doc.CurrentController.ActiveSheet = Doc.Sheets.getByName("Rozvrh")

end sub

sub Splneny_znak_vyber_noveho_postupny

	'vzhledem k tomu, jak se doplňují už udělané znaky (nikoliv za sebou), tak tohle makro vlastně nebude fungovat. Ale nechávám to tady.

	'zjištění jestli jde o čtení/predvedenie/jsme s kurzorem na špatném sloupci. Typ se uloží do "typ". Řádek se uloží do "riadok".
	If CurrentColumn = 34 Then typ = "ČÍTANIE"
	If CurrentColumn = 36 Then typ = "PREDVEDENIE"
	If CurrentColumn <> 36 AND CurrentColumn <> 34 Then
		MsgBox("Kurzor musí byť na čísle znaku, ktorý je splnený.")
		Exit Sub
	End If
	cislo_riadku = CurrentRow
	
	'získej přístup k dokumentu
	dim document   as object
	dim dispatcher as object
	Doc = ThisComponent
	document   = Doc.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

	'typ = "ČÍTANIE"
	If typ = "ČÍTANIE" Then
		
		'----------------zapsání hotového znaku--------------------
		'dej kurzor na připravený nový seznam hotových znaků
		dim args1(0) as new com.sun.star.beans.PropertyValue
		args1(0).Name = "ToPoint"
		args1(0).Value = "$BI$" & cislo_riadku	
		dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args1())
		'zkopíruj obsah
		dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
		'dej kurzor tam, kde chceme zapsat seznam hotových znaků
		dim args2(0) as new com.sun.star.beans.PropertyValue
		args2(0).Name = "ToPoint"
		args2(0).Value = "$AG$" & cislo_riadku	
		dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args2())
		'vlož jen text
		dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())
		
		'----------------vymyšlení nového znaku--------------------
		'dej kurzor na vymyšlený nový znak
		dim args3(0) as new com.sun.star.beans.PropertyValue
		args3(0).Name = "ToPoint"
		args3(0).Value = "$AM$" & cislo_riadku	
		dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args3())
		'zkopíruj obsah
		dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
		'dej kurzor tam, kde chceme zapsat vymyšlený znak
		dim args4(0) as new com.sun.star.beans.PropertyValue
		args4(0).Name = "ToPoint"
		args4(0).Value = "$AH$" & cislo_riadku	
		dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args4())
		'vlož jen číslo
		dispatcher.executeDispatch(document, ".uno:PasteOnlyValue", "", 0, Array())
		
		'--------------odznačení-----------------------------------	
		dim args5(1) as new com.sun.star.beans.PropertyValue
		args5(0).Name = "By"
		args5(0).Value = 1
		args5(1).Name = "Sel"
		args5(1).Value = false	
		dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args5())
		
		dim args6(1) as new com.sun.star.beans.PropertyValue
		args6(0).Name = "By"
		args6(0).Value = 1
		args6(1).Name = "Sel"
		args6(1).Value = false	
		dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, args6())
		
	End If
	
	'typ = "PREDVEDENIE"
	If typ = "PREDVEDENIE" Then
		
		'----------------zapsání hotového znaku--------------------
		'dej kurzor na připravený nový seznam hotových znaků
		dim args7(0) as new com.sun.star.beans.PropertyValue
		args7(0).Name = "ToPoint"
		args7(0).Value = "$BJ$" & cislo_riadku	
		dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args7())
		'zkopíruj obsah
		dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
		'dej kurzor tam, kde chceme zapsat seznam hotových znaků
		dim args8(0) as new com.sun.star.beans.PropertyValue
		args8(0).Name = "ToPoint"
		args8(0).Value = "$AG$" & cislo_riadku	
		dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args8())
		'vlož jen text
		dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())
		
		'----------------vymyšlení nového znaku--------------------
		'dej kurzor na vymyšlený nový znak
		dim args9(0) as new com.sun.star.beans.PropertyValue
		args9(0).Name = "ToPoint"
		args9(0).Value = "$BG$" & cislo_riadku	
		dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args9())
		'zkopíruj obsah
		dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
		'dej kurzor tam, kde chceme zapsat vymyšlený znak
		dim args10(0) as new com.sun.star.beans.PropertyValue
		args10(0).Name = "ToPoint"
		args10(0).Value = "$AJ$" & cislo_riadku	
		dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args10())
		'vlož jen číslo
		dispatcher.executeDispatch(document, ".uno:PasteOnlyValue", "", 0, Array())
		
		'--------------odznačení-----------------------------------	
		dim args11(1) as new com.sun.star.beans.PropertyValue
		args11(0).Name = "By"
		args11(0).Value = 1
		args11(1).Name = "Sel"
		args11(1).Value = false	
		dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args11())
		
		dim args12(1) as new com.sun.star.beans.PropertyValue
		args12(0).Name = "By"
		args12(0).Value = 1
		args12(1).Name = "Sel"
		args12(1).Value = false	
		dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, args12())
		
	End If
		
end sub


sub Splneny_znak_vyber_noveho_nahodny

	'zjištění jestli jde o čtení/predvedenie/jsme s kurzorem na špatném sloupci. Typ se uloží do "typ". Řádek se uloží do "riadok".
	If CurrentColumn = 34 Then typ = "ČÍTANIE"
	If CurrentColumn = 36 Then typ = "PREDVEDENIE"
	If CurrentColumn <> 36 AND CurrentColumn <> 34 Then
		MsgBox("Kurzor musí byť na čísle znaku, ktorý je splnený.")
		Exit Sub
	End If
	cislo_riadku = CurrentRow
	
	'získej přístup k dokumentu
	dim document   as object
	dim dispatcher as object
	Doc = ThisComponent
	document   = Doc.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

	'typ = "ČÍTANIE"
	If typ = "ČÍTANIE" Then
		
		'----------------zapsání hotového znaku--------------------
		'dej kurzor na připravený nový seznam hotových znaků
		dim args1(0) as new com.sun.star.beans.PropertyValue
		args1(0).Name = "ToPoint"
		args1(0).Value = "$BI$" & cislo_riadku	
		dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args1())
		'zkopíruj obsah
		dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
		'dej kurzor tam, kde chceme zapsat seznam hotových znaků
		dim args2(0) as new com.sun.star.beans.PropertyValue
		args2(0).Name = "ToPoint"
		args2(0).Value = "$AG$" & cislo_riadku	
		dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args2())
		'vlož jen text
		dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())
		
		'----------------vymyšlení nového znaku--------------------
		'dej kurzor na vymyšlený nový znak
		dim args3(0) as new com.sun.star.beans.PropertyValue
		args3(0).Name = "ToPoint"
		args3(0).Value = "$AT$" & cislo_riadku	
		dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args3())
		'zkopíruj obsah
		dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
		'dej kurzor tam, kde chceme zapsat vymyšlený znak
		dim args4(0) as new com.sun.star.beans.PropertyValue
		args4(0).Name = "ToPoint"
		args4(0).Value = "$AH$" & cislo_riadku	
		dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args4())
		'vlož jen číslo
		dispatcher.executeDispatch(document, ".uno:PasteOnlyValue", "", 0, Array())
		
		'--------------odznačení-----------------------------------	
		dim args5(1) as new com.sun.star.beans.PropertyValue
		args5(0).Name = "By"
		args5(0).Value = 1
		args5(1).Name = "Sel"
		args5(1).Value = false	
		dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args5())
		
		dim args6(1) as new com.sun.star.beans.PropertyValue
		args6(0).Name = "By"
		args6(0).Value = 1
		args6(1).Name = "Sel"
		args6(1).Value = false	
		dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, args6())
		
	End If
	
	'typ = "PREDVEDENIE"
	If typ = "PREDVEDENIE" Then
		
		'----------------zapsání hotového znaku--------------------
		'dej kurzor na připravený nový seznam hotových znaků
		dim args7(0) as new com.sun.star.beans.PropertyValue
		args7(0).Name = "ToPoint"
		args7(0).Value = "$BJ$" & cislo_riadku	
		dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args7())
		'zkopíruj obsah
		dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
		'dej kurzor tam, kde chceme zapsat seznam hotových znaků
		dim args8(0) as new com.sun.star.beans.PropertyValue
		args8(0).Name = "ToPoint"
		args8(0).Value = "$AG$" & cislo_riadku	
		dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args8())
		'vlož jen text
		dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())
		
		'----------------vymyšlení nového znaku--------------------
		'dej kurzor na vymyšlený nový znak
		dim args9(0) as new com.sun.star.beans.PropertyValue
		args9(0).Name = "ToPoint"
		args9(0).Value = "$BH$" & cislo_riadku	
		dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args9())
		'zkopíruj obsah
		dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
		'dej kurzor tam, kde chceme zapsat vymyšlený znak
		dim args10(0) as new com.sun.star.beans.PropertyValue
		args10(0).Name = "ToPoint"
		args10(0).Value = "$AJ$" & cislo_riadku	
		dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args10())
		'vlož jen číslo
		dispatcher.executeDispatch(document, ".uno:PasteOnlyValue", "", 0, Array())
		
		'--------------odznačení-----------------------------------	
		dim args11(1) as new com.sun.star.beans.PropertyValue
		args11(0).Name = "By"
		args11(0).Value = 1
		args11(1).Name = "Sel"
		args11(1).Value = false	
		dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args11())
		
		dim args12(1) as new com.sun.star.beans.PropertyValue
		args12(0).Name = "By"
		args12(0).Value = 1
		args12(1).Name = "Sel"
		args12(1).Value = false	
		dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, args12())
		
	End If
		
end sub


Function CurrentColumn() As Long
	'funkce zjišťuje současnou polohu kurzoru - sloupec (od 1,1)
	Dim ODoc As Object
	Dim OSel As Object
	oDoc = ThisComponent
	oSel = oDoc.GetCurrentSelection()
	If Not oSel.supportsService("com.sun.star.sheet.SheetCellRange") Then Exit Function
	If (oSel.Columns().Count() > 1) Then Exit Function
	CurrentColumn = oSel.CellAddress.Column()+1
End Function

Function CurrentRow() As Long
	'funkce zjišťuje současnou polohu kurzoru - řádek (od 1,1)
	Dim ODoc As Object
	Dim OSel As Object
	oDoc = ThisComponent
	oSel = oDoc.GetCurrentSelection()
	If Not oSel.supportsService("com.sun.star.sheet.SheetCellRange") Then Exit Function
	If (oSel.Rows().Count() > 1) Then Exit Function
	CurrentRow = oSel.CellAddress.Row()+1
End Function



sub Dotaz_predseda_splnil_nekdo_znak

	'zapsání data, které chybí, do buňky S1
	sText = InputBox ("Zadaj dátum, ku ktorému ti chýba od predsedajúceho správa o splnených znakoch (vo formátu dd.mm.yyyy).")
	Doc = ThisComponent
	Sheets = Doc.getSheets  'get the collection of Sheets
	Sheet = Sheets.GetByName("Rozvrh")  'get the Sheet
	Cell = Sheet.getCellByPosition(18,0)  'Get the Cell (S1)
	Cell.String = sText
	
	'funguje jen v Libreoffice, nefunguje v OpenOffice

   	eMailer = createUNOService("com.sun.star.system.SimpleCommandMail")

   	eMailClient = eMailer.querySimpleMailClient()

   	eMessage = eMailer.createSimpleMailMessage()

   	'predmet
   	eMessage.Subject = "Zhromaždenie Život a služba "  & sText & " – splnené znaky"
   	eMessage.Body = "Ahoj," & chr(10) & chr(10) & "prosím, ty si bol predsedajúci zhromaždenie Život a služba " & sText & ". Môžem sa spýtať, či niekto splnil svoj znak? Prosím, keby si mi napísal aj čísla znakov, lebo pri výmenách si niekedy vymieňajú aj znaky." & chr(10) & chr(10) & "Ďakujem," & chr(10) & chr(10) & "Lukáš"
   	eMailer.sendSimpleMailMessage (eMessage, com.sun.star.system.SimpleMailClientFlags.NO_USER_INTERFACE)

   	'skoč do listu "Rozvrh"
	Doc.CurrentController.ActiveSheet = Doc.Sheets.getByName("Rozvrh")

	'Získat číslo v buňce S2, dát jí do proměnné "cislo_riadku"
	cislo_riadku = ThisComponent.CurrentController.ActiveSheet.getCellRangeByName("S2").Value
   	
   	'zapsat že čekám na odpověď
   	Sheet = Sheets.GetByName("Rozvrh")  'get the Sheet
	Cell = Sheet.getCellByPosition(17,(cislo_riadku-1))  'Get the Cell (Rcislo_riadku-1)
	Cell.String = "(e-mail, čekám)"
	
	dim document   as object
	dim dispatcher as object

	'získej přístup k dokumentu
	document   = Doc.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
   	
	'dej kurzor na to místo
	dim args1(0) as new com.sun.star.beans.PropertyValue
	args1(0).Name = "ToPoint"
	args1(0).Value = "$R$" & cislo_riadku	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args1())

end sub

sub Tisk_S_89

	sText = InputBox ("Zadaj dátum, ku ktorému chceš vytlačiť S-89 (vo formátu dd.mm.yyyy).")
	Doc = ThisComponent
	Sheets = Doc.getSheets  'get the collection of Sheets
	Sheet = Sheets.GetByName("(TLAČ - S-89)")  'get the Sheet
	Cell = Sheet.getCellByPosition(0,0)  'Get the Cell (0,0=A1)
	Cell.String = sText
	
	dim document   as object
	dim dispatcher as object

	'získej přístup k dokumentu
	document   = Doc.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	'odkryj list
	dim args2(0) as new com.sun.star.beans.PropertyValue
	args2(0).Name = "aTableName"
	args2(0).Value = "(TLAČ - S-89)"
	dispatcher.executeDispatch(document, ".uno:Show", "", 0, args2())
	
	'skoč do listu "(TLAČ - S-89)"
	Doc.CurrentController.ActiveSheet = Doc.Sheets.getByName("(TLAČ - S-89)")
	
	'dej kurzor na začátek (=odznačení)
	dim args1(0) as new com.sun.star.beans.PropertyValue
	args1(0).Name = "ToPoint"
	args1(0).Value = "$A$1"	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args1())
	
	'vytiskni to
	dispatcher.executeDispatch(document, ".uno:Print", "", 0, Array())

	'trik, díky kterému to vytiskne pouze aktivní list	
	Wait 500
	
	'Získat číslo v buňce B1, dát jí do proměnné "riadok"
	riadok = ThisComponent.CurrentController.ActiveSheet.getCellRangeByName("B1").Value

	'zakryj list
	dim args3(0) as new com.sun.star.beans.PropertyValue
	args3(0).Name = "aTableName"
	args3(0).Value = "(TLAČ - S-89)"
	dispatcher.executeDispatch(document, ".uno:Hide", "", 0, args3())
	
	'skoč do listu "Rozvrh"
	Doc.CurrentController.ActiveSheet = Doc.Sheets.getByName("Rozvrh")
	
	'dát kurzor na D a tu proměnnou "riadok"
	dim args4(0) as new com.sun.star.beans.PropertyValue
	args4(0).Name = "ToPoint"
	args4(0).Value = "$D$" & riadok	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args4())
	
	'odboldovat řádek
	dim args5(0) as new com.sun.star.beans.PropertyValue
	args5(0).Name = "Bold"
	args5(0).Value = false	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args5())

	'1
	dim args6(1) as new com.sun.star.beans.PropertyValue
	args6(0).Name = "By"
	args6(0).Value = 1
	args6(1).Name = "Sel"
	args6(1).Value = false	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args6())
	
	dim args7(0) as new com.sun.star.beans.PropertyValue
	args7(0).Name = "Bold"
	args7(0).Value = false	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args7())

	'4 (přeskočili jsme VIDEO)
	dim args12(1) as new com.sun.star.beans.PropertyValue
	args12(0).Name = "By"
	args12(0).Value = 3
	args12(1).Name = "Sel"
	args12(1).Value = false	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args12())
	
	dim args13(0) as new com.sun.star.beans.PropertyValue
	args13(0).Name = "Bold"
	args13(0).Value = false	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args13())
	
	'5
	dim args14(1) as new com.sun.star.beans.PropertyValue
	args14(0).Name = "By"
	args14(0).Value = 1
	args14(1).Name = "Sel"
	args14(1).Value = false	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args14())
	
	dim args15(0) as new com.sun.star.beans.PropertyValue
	args15(0).Name = "Bold"
	args15(0).Value = false	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args15())
	
	'6
	dim args16(1) as new com.sun.star.beans.PropertyValue
	args16(0).Name = "By"
	args16(0).Value = 1
	args16(1).Name = "Sel"
	args16(1).Value = false	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args16())
	
	dim args17(0) as new com.sun.star.beans.PropertyValue
	args17(0).Name = "Bold"
	args17(0).Value = false	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args17())

	'7
	dim args18(1) as new com.sun.star.beans.PropertyValue
	args18(0).Name = "By"
	args18(0).Value = 1
	args18(1).Name = "Sel"
	args18(1).Value = false	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args18())
	
	dim args19(0) as new com.sun.star.beans.PropertyValue
	args19(0).Name = "Bold"
	args19(0).Value = false	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args19())
	
	'8
	dim args20(1) as new com.sun.star.beans.PropertyValue
	args20(0).Name = "By"
	args20(0).Value = 1
	args20(1).Name = "Sel"
	args20(1).Value = false	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args20())
	
	dim args21(0) as new com.sun.star.beans.PropertyValue
	args21(0).Name = "Bold"
	args21(0).Value = false	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args21())
	
	'9
	dim args22(1) as new com.sun.star.beans.PropertyValue
	args22(0).Name = "By"
	args22(0).Value = 1
	args22(1).Name = "Sel"
	args22(1).Value = false	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args22())
	
	dim args23(0) as new com.sun.star.beans.PropertyValue
	args23(0).Name = "Bold"
	args23(0).Value = false	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args23())

	'10
	dim args24(1) as new com.sun.star.beans.PropertyValue
	args24(0).Name = "By"
	args24(0).Value = 1
	args24(1).Name = "Sel"
	args24(1).Value = false	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args24())
	
	dim args25(0) as new com.sun.star.beans.PropertyValue
	args25(0).Name = "Bold"
	args25(0).Value = false	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args25())

	'11
	dim args26(1) as new com.sun.star.beans.PropertyValue
	args26(0).Name = "By"
	args26(0).Value = 1
	args26(1).Name = "Sel"
	args26(1).Value = false	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args26())
	
	dim args27(0) as new com.sun.star.beans.PropertyValue
	args27(0).Name = "Bold"
	args27(0).Value = false	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args27())

	'12
	dim args28(1) as new com.sun.star.beans.PropertyValue
	args28(0).Name = "By"
	args28(0).Value = 1
	args28(1).Name = "Sel"
	args28(1).Value = false	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args28())
	
	dim args29(0) as new com.sun.star.beans.PropertyValue
	args29(0).Name = "Bold"
	args29(0).Value = false	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args29())
	
	'zapsat x, že bylo vytisknuto
	dim args31(0) as new com.sun.star.beans.PropertyValue
	args31(0).Name = "ToPoint"
	args31(0).Value = "$T$" & riadok	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args31())
	
	dim args32(0) as new com.sun.star.beans.PropertyValue
	args32(0).Name = "StringName"
	args32(0).Value = "x"

	dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args32())
		
	'skočit s kurzorem na začátek
	dim args30(0) as new com.sun.star.beans.PropertyValue
	args30(0).Name = "ToPoint"
	args30(0).Value = "$D$" & riadok	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args30())
	
end sub


Sub Poslani_rozvrhu_koordinatorovi
	
	'-------------------1) ZBAVÍME SE VZORCŮ------------------------------
	dim document   as object
	dim dispatcher as object
	
	'získej přístup k dokumentu
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	'odkryj listy
	dim args8(0) as new com.sun.star.beans.PropertyValue
	args8(0).Name = "aTableName"
	args8(0).Value = "(Export pre koordinátora 1)"
	dispatcher.executeDispatch(document, ".uno:Show", "", 0, args8())
	dim args7(0) as new com.sun.star.beans.PropertyValue
	args7(0).Name = "aTableName"
	args7(0).Value = "(Export pre koordinátora 2)"
	dispatcher.executeDispatch(document, ".uno:Show", "", 0, args7())
	
	dim args11(0) as new com.sun.star.beans.PropertyValue
	args11(0).Name = "aTableName"
	args11(0).Value = "(Export pre koordinátora 3)"
	dispatcher.executeDispatch(document, ".uno:Show", "", 0, args11())

	'skoč do listu (Export pre koordinátora 1)
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("(Export pre koordinátora 1)")
	
	'zjisti Prvý den mesiaca, na ktorý je rozvrh
	prvy_den_mesiaca = ThisComponent.CurrentController.ActiveSheet.getCellRangeByName("P8").String	
	'zjisti Posledný riadok v (Export pre koordinátora 3)
	riadok_kam_dat_x = ThisComponent.CurrentController.ActiveSheet.getCellRangeByName("P10").String
	
	'vyber tabulku
	dim args1(0) as new com.sun.star.beans.PropertyValue
	args1(0).Name = "ToPoint"
	args1(0).Value = "$A$1:$N$7"	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args1())
	
	'zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	'skoč do listu (Export pre koordinátora 2)
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("(Export pre koordinátora 2)")
	
	'vlož obsah - pouze čísla, text, formát
	dim args4(5) as new com.sun.star.beans.PropertyValue
	args4(0).Name = "Flags"
	args4(0).Value = "SVDT"
	args4(1).Name = "FormulaCommand"
	args4(1).Value = 0
	args4(2).Name = "SkipEmptyCells"
	args4(2).Value = false
	args4(3).Name = "Transpose"
	args4(3).Value = false
	args4(4).Name = "AsLink"
	args4(4).Value = false
	args4(5).Name = "MoveMode"
	args4(5).Value = 4
	
	dispatcher.executeDispatch(document, ".uno:InsertContents", "", 0, args4())
	
	'dej kurzor na začátek (=odznačení)
	dim args5(0) as new com.sun.star.beans.PropertyValue
	args5(0).Name = "ToPoint"
	args5(0).Value = "$A$1"	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args5())
	
	'-------------------2) VYTVOŘÍME PŘÍLOHU DO MAILU ------------------------------	

	Dim doc1
	Dim doc2
	Dim FileN As String
	Dim storeParms(1) as new com.sun.star.beans.PropertyValue
  
	doc1 = ThisComponent
  
	'název listu, který se kopíruje do nového dokumentu
	doc1.getCurrentController.select(doc1.getSheets().getByName("(Export pre koordinátora 2)"))
	'vyber vše na daném listu
	dispatchURL(doc1,".uno:SelectAll")
	'zkopíruj
	dispatchURL(doc1,".uno:Copy")
	
	'vytvoř nový dokument
	doc2 = StarDesktop.loadComponentFromUrl("private:factory/scalc","_blank",0,dimArray())
	'vlož
	dispatchURL(doc2,".uno:Paste")
	
	'zjistíme současnou cestu k souboru (Cesta)
   	GlobalScope.BasicLibraries.LoadLibrary("Tools") ' Only for GetFileName
   	Cesta = DirectoryNameoutofPath(doc1.getURL(),"/")
    FileN = Cesta & "/Export pre koordinátora/Rozdelenie zhromaždenia ŽaS.xlsx"
	 
	'formát souboru
	storeParms(0).Name = "FilterName"
	storeParms(0).Value = "Calc MS Excel 2007 XML" 
	'pokud už tam soubor stejného jména je, tak ho přepiš 
   	storeParms(1).Name = "Overwrite"
   	storeParms(1).Value = True 
	'uložení souboru
	doc2.storeToURL(FileN,storeParms())
	'zavření souboru
	doc2.close(True)
	
	'smaž obsah dočasné tabulky (Export pre koordinátora 2)
	dispatcher.executeDispatch(document, ".uno:ClearContents", "", 0, Array())
	
	'dej kurzor na začátek (=odznačení)
	dim args6(0) as new com.sun.star.beans.PropertyValue
	args6(0).Name = "ToPoint"
	args6(0).Value = "$A$1"	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args6())
	
	'skoč do listu "(Export pre koordinátora 3)"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("(Export pre koordinátora 3)")
	
	'zapsat datum a x
   	Sheet = ThisComponent.getSheets.GetByName("(Export pre koordinátora 3)")  'get the Sheet
	Cell = Sheet.getCellByPosition(0,riadok_kam_dat_x)  'Get the Cell (Rcislo_riadku-1)
	Cell.String = prvy_den_mesiaca
	
	Cell = Sheet.getCellByPosition(1,riadok_kam_dat_x)  'Get the Cell (Rcislo_riadku-1)
	Cell.String = "x"	
	
	'zakryj listy
	dim args9(0) as new com.sun.star.beans.PropertyValue
	args9(0).Name = "aTableName"
	args9(0).Value = "(Export pre koordinátora 1)"
	dispatcher.executeDispatch(document, ".uno:Hide", "", 0, args9())
	dim args10(0) as new com.sun.star.beans.PropertyValue
	args10(0).Name = "aTableName"
	args10(0).Value = "(Export pre koordinátora 2)"
	dispatcher.executeDispatch(document, ".uno:Hide", "", 0, args10())	
	dim args12(0) as new com.sun.star.beans.PropertyValue
	args12(0).Name = "aTableName"
	args12(0).Value = "(Export pre koordinátora 3)"
	dispatcher.executeDispatch(document, ".uno:Hide", "", 0, args12())
	
	'skoč do listu "Rozvrh"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Rozvrh")
	
	'-------------------3) POŠLEME MAIL S PŘÍLOHOU ------------------------------
	
	'funguje jen v Libreoffice, nefunguje v OpenOffice

   	eMailer = createUNOService("com.sun.star.system.SimpleCommandMail")

   	eMailClient = eMailer.querySimpleMailClient()

   	eMessage = eMailer.createSimpleMailMessage()

	'e-mailová adresa koordinátora zboru
   	eMessage.Recipient = "posli.posli@email.cz"
   	'e-mailové adresy ostatních starších
   	eMessage.CcRecipient = Array("kakadu.bajak@gmail.com", "boris.merai@gmail.com", "milos.kotula@gmail.com", "igirazor@gmail.com")
   	'predmet
   	eMessage.Subject = "Rozdelenie zhromaždenia ŽaS"
   	eMessage.Body = "Ahoj Lukáš," & chr(10) & chr(10) & "posielam ti ďalšie rozdelenie zhromaždenia Život a služba." & chr(10) & chr(10) & "Ďakujem," & chr(10) & chr(10) & "Lukáš"
   	eMessage.Attachement = Split(FileN, chr(13))
    'eMessage.Attachement
   	eMailer.sendSimpleMailMessage (eMessage, com.sun.star.system.SimpleMailClientFlags.NO_USER_INTERFACE)
  
End Sub


sub Rychle_prideleni

	rem - define variables
	dim document   as object
	dim dispatcher as object
	
	rem - get access to the document
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	rem ------------------------ v tabulce Študenti nakopíruj hodnoty ze vzorců do plaintextu --------------------------------------
	
	'skoč do listu "Študenti"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Študenti")
	
	rem - jdi do buňky M3
	dim args1(0) as new com.sun.star.beans.PropertyValue
	args1(0).Name = "ToPoint"
	args1(0).Value = "$M$3"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args1())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	rem - jdi do buňky O3
	dim args2(0) as new com.sun.star.beans.PropertyValue
	args2(0).Name = "ToPoint"
	args2(0).Value = "$O$3"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args2())
	
	rem - vlož jen text	
	dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())
	
	rem - jdi do buňky N3
	dim args4(0) as new com.sun.star.beans.PropertyValue
	args4(0).Name = "ToPoint"
	args4(0).Value = "$N$3"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args4())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	rem - jdi do buňky P3
	dim args5(0) as new com.sun.star.beans.PropertyValue
	args5(0).Name = "ToPoint"
	args5(0).Value = "$P$3"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args5())
	
	rem - vlož jen číslo	
	dispatcher.executeDispatch(document, ".uno:PasteOnlyValue", "", 0, Array())
	
	rem - jdi do buňky M4
	dim args6(0) as new com.sun.star.beans.PropertyValue
	args6(0).Name = "ToPoint"
	args6(0).Value = "$M$4"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args6())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	rem - jdi do buňky O4
	dim args7(0) as new com.sun.star.beans.PropertyValue
	args7(0).Name = "ToPoint"
	args7(0).Value = "$O$4"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args7())
	
	rem - vlož jen text	
	dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())
	
	rem - jdi do buňky M5
	dim args8(0) as new com.sun.star.beans.PropertyValue
	args8(0).Name = "ToPoint"
	args8(0).Value = "$M$5"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args8())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	rem - jdi do buňky O5
	dim args9(0) as new com.sun.star.beans.PropertyValue
	args9(0).Name = "ToPoint"
	args9(0).Value = "$O$5"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args9())
	
	rem - vlož jen text	
	dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())
	
	rem - jdi do buňky M12
	dim args10(0) as new com.sun.star.beans.PropertyValue
	args10(0).Name = "ToPoint"
	args10(0).Value = "$M$12"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args10())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	rem - jdi do buňky O12
	dim args11(0) as new com.sun.star.beans.PropertyValue
	args11(0).Name = "ToPoint"
	args11(0).Value = "$O$12"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args11())
	
	rem - vlož jen text	
	dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())
	
	rem - jdi do buňky M13
	dim args12(0) as new com.sun.star.beans.PropertyValue
	args12(0).Name = "ToPoint"
	args12(0).Value = "$M$13"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args12())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	rem - jdi do buňky O13
	dim args13(0) as new com.sun.star.beans.PropertyValue
	args13(0).Name = "ToPoint"
	args13(0).Value = "$O$13"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args13())
	
	rem - vlož jen text	
	dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())
	
	rem - jdi do buňky M14
	dim args14(0) as new com.sun.star.beans.PropertyValue
	args14(0).Name = "ToPoint"
	args14(0).Value = "$M$14"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args14())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	rem - jdi do buňky O14
	dim args15(0) as new com.sun.star.beans.PropertyValue
	args15(0).Name = "ToPoint"
	args15(0).Value = "$O$14"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args15())
	
	rem - vlož jen text	
	dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())
	
	rem - jdi do buňky M15
	dim args16(0) as new com.sun.star.beans.PropertyValue
	args16(0).Name = "ToPoint"
	args16(0).Value = "$M$15"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args16())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	rem - jdi do buňky O15
	dim args17(0) as new com.sun.star.beans.PropertyValue
	args17(0).Name = "ToPoint"
	args17(0).Value = "$O$15"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args17())
	
	rem - vlož jen text	
	dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())
	
	rem - jdi do buňky M16
	dim args18(0) as new com.sun.star.beans.PropertyValue
	args18(0).Name = "ToPoint"
	args18(0).Value = "$M$16"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args18())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	rem - jdi do buňky O16
	dim args19(0) as new com.sun.star.beans.PropertyValue
	args19(0).Name = "ToPoint"
	args19(0).Value = "$O$16"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args19())
	
	rem - vlož jen text	
	dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())
	
	rem - jdi do buňky M17
	dim args20(0) as new com.sun.star.beans.PropertyValue
	args20(0).Name = "ToPoint"
	args20(0).Value = "$M$17"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args20())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	rem - jdi do buňky O17
	dim args21(0) as new com.sun.star.beans.PropertyValue
	args21(0).Name = "ToPoint"
	args21(0).Value = "$O$17"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args21())
	
	rem - vlož jen text	
	dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())

	rem - jdi do buňky N12
	dim args22(0) as new com.sun.star.beans.PropertyValue
	args22(0).Name = "ToPoint"
	args22(0).Value = "$N$12"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args22())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	rem - jdi do buňky P12
	dim args23(0) as new com.sun.star.beans.PropertyValue
	args23(0).Name = "ToPoint"
	args23(0).Value = "$P$12"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args23())
	
	rem - vlož jen číslo	
	dispatcher.executeDispatch(document, ".uno:PasteOnlyValue", "", 0, Array())	
	
	rem - jdi do buňky N14
	dim args24(0) as new com.sun.star.beans.PropertyValue
	args24(0).Name = "ToPoint"
	args24(0).Value = "$N$14"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args24())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	rem - jdi do buňky P14
	dim args25(0) as new com.sun.star.beans.PropertyValue
	args25(0).Name = "ToPoint"
	args25(0).Value = "$P$14"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args25())
	
	rem - vlož jen číslo	
	dispatcher.executeDispatch(document, ".uno:PasteOnlyValue", "", 0, Array())	
	
	rem - jdi do buňky N16
	dim args26(0) as new com.sun.star.beans.PropertyValue
	args26(0).Name = "ToPoint"
	args26(0).Value = "$N$16"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args26())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	rem - jdi do buňky P16
	dim args27(0) as new com.sun.star.beans.PropertyValue
	args27(0).Name = "ToPoint"
	args27(0).Value = "$P$16"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args27())
	
	rem - vlož jen číslo	
	dispatcher.executeDispatch(document, ".uno:PasteOnlyValue", "", 0, Array())	
	
	rem ------------------------ zkopíruj jména a znaky do tabulky Rozvrh --------------------------------------
	
	rem - jdi do buňky O3
	dim args28(0) as new com.sun.star.beans.PropertyValue
	args28(0).Name = "ToPoint"
	args28(0).Value = "$O$3"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args28())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	'skoč do listu "Rozvrh"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Rozvrh")
	
	rem - vlož jen text	
	dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())
	
	rem - dej písmo jako tučné
	dim args30(0) as new com.sun.star.beans.PropertyValue
	args30(0).Name = "Bold"
	args30(0).Value = true
	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args30())
	
	rem - šipka doprava
	dim args31(1) as new com.sun.star.beans.PropertyValue
	args31(0).Name = "By"
	args31(0).Value = 1
	args31(1).Name = "Sel"
	args31(1).Value = false
	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args31())
	
	
	
	
	'skoč do listu "Študenti"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Študenti")
	
	rem - jdi do buňky P3
	dim args33(0) as new com.sun.star.beans.PropertyValue
	args33(0).Name = "ToPoint"
	args33(0).Value = "$P$3"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args33())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	'skoč do listu "Rozvrh"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Rozvrh")
	
	rem - vlož jen číslo	
	dispatcher.executeDispatch(document, ".uno:PasteOnlyValue", "", 0, Array())
	
	rem - dej písmo jako tučné
	dim args35(0) as new com.sun.star.beans.PropertyValue
	args35(0).Name = "Bold"
	args35(0).Value = true
	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args35())
	
	rem - šipka doprava
	dim args36(1) as new com.sun.star.beans.PropertyValue
	args36(0).Name = "By"
	args36(0).Value = 1
	args36(1).Name = "Sel"
	args36(1).Value = false
	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args36())
	
	
	
	
	'skoč do listu "Študenti"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Študenti")
	
	rem - jdi do buňky O4
	dim args38(0) as new com.sun.star.beans.PropertyValue
	args38(0).Name = "ToPoint"
	args38(0).Value = "$O$4"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args38())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	'skoč do listu "Rozvrh"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Rozvrh")
	
	rem - vlož jen text	
	dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())
	
	rem - dej písmo jako tučné
	dim args40(0) as new com.sun.star.beans.PropertyValue
	args40(0).Name = "Bold"
	args40(0).Value = true
	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args40())
	
	rem - šipka doprava
	dim args41(1) as new com.sun.star.beans.PropertyValue
	args41(0).Name = "By"
	args41(0).Value = 1
	args41(1).Name = "Sel"
	args41(1).Value = false
	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args41())
		
	
	
	
	'skoč do listu "Študenti"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Študenti")
	
	rem - jdi do buňky O5
	dim args43(0) as new com.sun.star.beans.PropertyValue
	args43(0).Name = "ToPoint"
	args43(0).Value = "$O$5"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args43())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	'skoč do listu "Rozvrh"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Rozvrh")
	
	rem - vlož jen text	
	dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())
	
	rem - dej písmo jako tučné
	dim args45(0) as new com.sun.star.beans.PropertyValue
	args45(0).Name = "Bold"
	args45(0).Value = true
	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args45())
	
	rem - šipka doprava
	dim args46(1) as new com.sun.star.beans.PropertyValue
	args46(0).Name = "By"
	args46(0).Value = 1
	args46(1).Name = "Sel"
	args46(1).Value = false
	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args46())
			
	
	
	
	'skoč do listu "Študenti"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Študenti")
	
	rem - jdi do buňky O12
	dim args48(0) as new com.sun.star.beans.PropertyValue
	args48(0).Name = "ToPoint"
	args48(0).Value = "$O$12"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args48())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	'skoč do listu "Rozvrh"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Rozvrh")
	
	rem - vlož jen text	
	dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())
	
	rem - dej písmo jako tučné
	dim args50(0) as new com.sun.star.beans.PropertyValue
	args50(0).Name = "Bold"
	args50(0).Value = true
	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args50())
	
	rem - šipka doprava
	dim args51(1) as new com.sun.star.beans.PropertyValue
	args51(0).Name = "By"
	args51(0).Value = 1
	args51(1).Name = "Sel"
	args51(1).Value = false
	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args51())
				
	
	
	
	'skoč do listu "Študenti"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Študenti")
	
	rem - jdi do buňky P12
	dim args53(0) as new com.sun.star.beans.PropertyValue
	args53(0).Name = "ToPoint"
	args53(0).Value = "$P$12"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args53())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	'skoč do listu "Rozvrh"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Rozvrh")
	
	rem - vlož jen číslo
	dispatcher.executeDispatch(document, ".uno:PasteOnlyValue", "", 0, Array())
	
	rem - dej písmo jako tučné
	dim args55(0) as new com.sun.star.beans.PropertyValue
	args55(0).Name = "Bold"
	args55(0).Value = true
	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args55())
	
	rem - šipka doprava
	dim args56(1) as new com.sun.star.beans.PropertyValue
	args56(0).Name = "By"
	args56(0).Value = 1
	args56(1).Name = "Sel"
	args56(1).Value = false
	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args56())
					
	
	
	
	'skoč do listu "Študenti"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Študenti")
	
	rem - jdi do buňky O13
	dim args58(0) as new com.sun.star.beans.PropertyValue
	args58(0).Name = "ToPoint"
	args58(0).Value = "$O$13"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args58())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	'skoč do listu "Rozvrh"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Rozvrh")
	
	rem - vlož jen text
	dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())
	
	rem - dej písmo jako tučné
	dim args60(0) as new com.sun.star.beans.PropertyValue
	args60(0).Name = "Bold"
	args60(0).Value = true
	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args60())
	
	rem - šipka doprava
	dim args61(1) as new com.sun.star.beans.PropertyValue
	args61(0).Name = "By"
	args61(0).Value = 1
	args61(1).Name = "Sel"
	args61(1).Value = false
	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args61())
		
	'skoč do listu "Študenti"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Študenti")
	
	rem - jdi do buňky O14
	dim args63(0) as new com.sun.star.beans.PropertyValue
	args63(0).Name = "ToPoint"
	args63(0).Value = "$O$14"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args63())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	'skoč do listu "Rozvrh"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Rozvrh")
	
	rem - vlož jen text	
	dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())
	
	rem - dej písmo jako tučné
	dim args65(0) as new com.sun.star.beans.PropertyValue
	args65(0).Name = "Bold"
	args65(0).Value = true
	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args65())
	
	rem - šipka doprava
	dim args66(1) as new com.sun.star.beans.PropertyValue
	args66(0).Name = "By"
	args66(0).Value = 1
	args66(1).Name = "Sel"
	args66(1).Value = false
	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args66())
				
	
	
	
	'skoč do listu "Študenti"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Študenti")
	
	rem - jdi do buňky P14
	dim args68(0) as new com.sun.star.beans.PropertyValue
	args68(0).Name = "ToPoint"
	args68(0).Value = "$P$14"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args68())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	'skoč do listu "Rozvrh"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Rozvrh")
	
	rem - vlož jen číslo
	dispatcher.executeDispatch(document, ".uno:PasteOnlyValue", "", 0, Array())
	
	rem - dej písmo jako tučné
	dim args70(0) as new com.sun.star.beans.PropertyValue
	args70(0).Name = "Bold"
	args70(0).Value = true
	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args70())
	
	rem - šipka doprava
	dim args71(1) as new com.sun.star.beans.PropertyValue
	args71(0).Name = "By"
	args71(0).Value = 1
	args71(1).Name = "Sel"
	args71(1).Value = false
	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args71())
					
	
	
	
	'skoč do listu "Študenti"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Študenti")
	
	rem - jdi do buňky O15
	dim args73(0) as new com.sun.star.beans.PropertyValue
	args73(0).Name = "ToPoint"
	args73(0).Value = "$O$15"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args73())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	'skoč do listu "Rozvrh"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Rozvrh")
	
	rem - vlož jen text
	dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())
	
	rem - dej písmo jako tučné
	dim args75(0) as new com.sun.star.beans.PropertyValue
	args75(0).Name = "Bold"
	args75(0).Value = true
	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args75())
	
	rem - šipka doprava
	dim args76(1) as new com.sun.star.beans.PropertyValue
	args76(0).Name = "By"
	args76(0).Value = 1
	args76(1).Name = "Sel"
	args76(1).Value = false
	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args76())
				
	
	
	
	'skoč do listu "Študenti"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Študenti")
	
	rem - jdi do buňky O16
	dim args78(0) as new com.sun.star.beans.PropertyValue
	args78(0).Name = "ToPoint"
	args78(0).Value = "$O$16"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args78())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	'skoč do listu "Rozvrh"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Rozvrh")
	
	rem - vlož jen text	
	dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())
	
	rem - dej písmo jako tučné
	dim args80(0) as new com.sun.star.beans.PropertyValue
	args80(0).Name = "Bold"
	args80(0).Value = true
	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args80())
	
	rem - šipka doprava
	dim args81(1) as new com.sun.star.beans.PropertyValue
	args81(0).Name = "By"
	args81(0).Value = 1
	args81(1).Name = "Sel"
	args81(1).Value = false
	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args81())
				
	
	
	
	'skoč do listu "Študenti"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Študenti")
	
	rem - jdi do buňky P16
	dim args83(0) as new com.sun.star.beans.PropertyValue
	args83(0).Name = "ToPoint"
	args83(0).Value = "$P$16"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args83())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	'skoč do listu "Rozvrh"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Rozvrh")
	
	rem - vlož jen číslo
	dispatcher.executeDispatch(document, ".uno:PasteOnlyValue", "", 0, Array())
	
	rem - dej písmo jako tučné
	dim args85(0) as new com.sun.star.beans.PropertyValue
	args85(0).Name = "Bold"
	args85(0).Value = true
	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args85())
	
	rem - šipka doprava
	dim args86(1) as new com.sun.star.beans.PropertyValue
	args86(0).Name = "By"
	args86(0).Value = 1
	args86(1).Name = "Sel"
	args86(1).Value = false
	
	dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args86())
					
	
	
	
	'skoč do listu "Študenti"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Študenti")
	
	rem - jdi do buňky O17
	dim args88(0) as new com.sun.star.beans.PropertyValue
	args88(0).Name = "ToPoint"
	args88(0).Value = "$O$17"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args88())
	
	rem - zkopíruj obsah
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	
	'skoč do listu "Rozvrh"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Rozvrh")
	
	rem - vlož jen text
	dispatcher.executeDispatch(document, ".uno:PasteOnlyText", "", 0, Array())
	
	rem - dej písmo jako tučné
	dim args90(0) as new com.sun.star.beans.PropertyValue
	args90(0).Name = "Bold"
	args90(0).Value = true
	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args90())
	
	rem -------------------------------- odznačení a vrácení kurzoru tam kde byl v obou sešitech --------------------------------------
	
	rem - šipka doleva na začátek
	dim args91(1) as new com.sun.star.beans.PropertyValue
	args91(0).Name = "By"
	args91(0).Value = 12
	args91(1).Name = "Sel"
	args91(1).Value = false
	
	dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, args91())

	'skoč do listu "Študenti"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Študenti")
	
	rem - označení předtím zkopírovaných hodnot
	dim args93(0) as new com.sun.star.beans.PropertyValue
	args93(0).Name = "ToPoint"
	args93(0).Value = "$O$3:$P$17"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args93())
	
	rem - vymazání
	dispatcher.executeDispatch(document, ".uno:ClearContents", "", 0, Array())
	
	rem - dát kurzor do první buňky
	dim args94(0) as new com.sun.star.beans.PropertyValue
	args94(0).Name = "ToPoint"
	args94(0).Value = "$A$1"
	
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args94())
	
	'skoč do listu "Rozvrh"
	ThisComponent.CurrentController.ActiveSheet = ThisComponent.Sheets.getByName("Rozvrh")
	
end sub



Sub dispatchURL(oDoc, aURL)

	'makro potřebné pro funkci makra Poslani_rozvrhu_koordinatorovi
	Dim noProps()
	Dim URL As New com.sun.star.util.URL
	Dim frame
	Dim transf
	Dim disp

	frame = oDoc.getCurrentController().getFrame()
	URL.Complete = aURL
	transf = createUnoService("com.sun.star.util.URLTransformer")
	transf.parseStrict(URL)

	disp = frame.queryDispatch(URL, "", _
            com.sun.star.frame.FrameSearchFlag.SELF _
         OR com.sun.star.frame.FrameSearchFlag.CHILDREN)
	disp.dispatch(URL, noProps())
  
End Sub


sub Prepocitani_Calcu

	dim document   as object
	dim dispatcher as object
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

	dispatcher.executeDispatch(document, ".uno:CalculateHard", "", 0, Array())

end sub
