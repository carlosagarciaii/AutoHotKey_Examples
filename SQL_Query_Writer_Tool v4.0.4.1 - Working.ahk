#SingleInstance  Force

vShotAHK := Object()
vShotAHK[1] := 1
vWaitLoop := Object()
vTable1 := Object()
vTable1A := Object()
vTable2 := Object()
vTable2A := Object()
vReverseLinkDirection := Object()
vLeftJoin := Object()
vHasExtension := Object()
vFilter1A := Object()
vFilter1Code := Object()
vNoLock := Object()
vAppVersion := Object()
vAppVersion[1] := "4.0.4.1"
vAppUpdateDate := Object()
vAppUpdateDate[1] := "2019-10-17"




EnvGet, vSQLQuickINIPath, LOCALAPPDATA

vSQLQuickINIPath = %vSQLQuickINIPath%\SQLQuickFiles



IF !InStr(FileExist(vSQLQuickINIPath),"D") {
		FileCreateDir, %vSQLQuickINIPath%	
		}
		
vSQLQuickINIPath = %vSQLQuickINIPath%\SQLQuick.ini
		
IF !FileExist(vSQLQuickINIPath) {
	
			IniWrite, AsmtAcct , %vSQLQuickINIPath%, FirstTable, TableName
	}

try {
		Menu, Tray, Icon, SQLIcon.ico
	}
Gui, MainApp:New, +AlwaysOnTop


try {
		Gui, MainApp:Font,s10,Verdana
	}
catch {
		Gui, MainApp:Font,s10,Arial
		
	}

/*
	Tray Icon Behavior
*/
		Menu, Tray, Add, Hide/Show,HideShowHandler
		Menu, Tray, Default, Hide/Show
		Menu, Tray, Click, 1

/*
	Menu Bar
*/
		Menu, MenuBar, Add, Hide, HideShowHandler

		Menu, HelpMenu, Add, Help, HelpHandler
		Menu, HelpMenu, Add, About, AboutHandler
		Menu, MenuBar, Add, Help, :HelpMenu
			Gui, MainApp:Menu, MenuBar






/*
	General 
*/

		try {
				Gui, MainApp:Font,s14,Verdana
			}
		catch {
				Gui, MainApp:Font,s14,Arial
				
			}


		
		Gui, MainApp:Add, Text, x15 y475 w450 vCounter, 

		try {
				Gui, MainApp:Font,s10,Verdana
			}
		catch {
				Gui, MainApp:Font,s10,Arial
				
			}
		
		Gui, MainApp:Add, Button, x10 y510 w100 gCloseApp vCloseAppButton, Close App  


	Gui, MainApp:Add, Tab3,y10 x10 h450 w450 vTabList,Header|Tables 1 and 2|CLD LKP|XML Path		;

/*
	SQL Header
*/

		Gui,MainApp:Tab,1
			StringRight, Payable, A_YYYY,2
			AsmtYear :=  A_YYYY -1
			StringRight, AsmtYear, AsmtYear,2
			
			TaxCycle = %AsmtYear%%Payable%
			

			Gui, MainApp:Add, Text, x25 y45 ,Tax Cycle
			Gui, MainApp:Add, Edit, x25 y60 w125 vTaxCycle t1,%TaxCycle%
			Gui, MainApp:Add, Checkbox,x25 y90 vRollTableLimited t2, Limit to Roll Year
			Gui, MainApp:Add, Checkbox,x25 y110 vTmpTableRemove CHECKED t3, Remove Temp Tables
			
			Gui, MainApp:Add, ListBox,x200 y45 vTablesIncluded Multi Sort  h200 t4, TABLE1|TABLE2|TABLE3|TABLE4
			
			
			Gui, MainApp:Add, Button, x25 y245 w325 gSQLHeaderExecute t5, Write SQL Header
			
		
/*
	Tables 1 & 2
*/
		Gui,MainApp:Tab,2

			Gui, MainApp:Add, Text, x25  y50 ,First Table
			Gui, MainApp:Add, ComboBox, x25  y70 w125 vTable1 gSetTable1XmlLabel  , TABLE1|TABLE2|TABLE3|TABLE4
			Gui, MainApp:Add, Text, x25  y95 ,First Table Abbreviation
			Gui, MainApp:Add, Edit, x25  y115 w125 vTable1A  

			Gui, MainApp:Add, Checkbox,x25 y145  vHasExtension1 , Has Extension

			
			Gui, MainApp:Add, Text, x260 y50 ,Second Table
			Gui, MainApp:Add, ComboBox, x260 y70 w125 vTable2 gSetTable2XmlLabel ,   TABLE1|TABLE2|TABLE3|TABLE4
			Gui, MainApp:Add, Text, x260 y95 ,Second Table Abbreviation
			Gui, MainApp:Add, Edit, x260 y115 w125 vTable2A 
			
			Gui, MainApp:Add, Checkbox,x260 y145 vReverseLinkDirection , Reverse Link
			Gui, MainApp:Add, Checkbox,x260 y160  vLeftJoin , Left Join
			Gui, MainApp:Add, Checkbox,x260 y175  vHasExtension , Has Extension
			Gui, MainApp:Add, Checkbox,x260 y190  vHasLinkExtension , Has Link Extension
			

			Gui, MainApp:Add, Button, x25 y215 w130 gFirstTableExecute , First Table 
			Gui, MainApp:Add, Button, x260 y215 w130 gJoinTableOnly, Join Table Only

			
			Gui, MainApp:Add, GroupBox, x240 y250 w170 h85, Not In 
				Gui, MainApp:Add, Checkbox, x260 y270 vNotIn Checked , Not In 
				Gui, MainApp:Add, Button,  x260 y290 w130 gNotInTableExecute, Not In Table

			Gui, MainApp:Add, Button, x190 y80 w50 g2to1, <--
			Gui, MainApp:Add, Button, x190 y110 w50 gClear2, Clr->
			Gui, MainApp:Add, Button, x190 y140 w50 g1to2, -->
	
			



/*
	CLD Lookup Table
*/

		Gui,MainApp:Tab,3
		

			Gui, MainApp:Add, Text, ,Filter Abbreviation
			Gui, MainApp:Add, Edit, w100 vFilter1A
			Gui, MainApp:Add, Text, ,Filter
			Gui, MainApp:Add, ListBox, vFilter1Code w200 h200, Table1|Table2|Table3
	
			Gui, MainApp:Add, Button, w200 gCDLkpExecute, Run CLD
		


/*
	XML Path
*/

		Gui,MainApp:Tab,4
		
			Gui, MainApp:Add, Text,x25 y50 w100 ,Table 1
			Gui, MainApp:Add, Text,x25 y70  w100 vTable1XmlLabel ,			
			Gui, MainApp:Add, Text,x150 y50 w100 , Table 2
			Gui, MainApp:Add, Text,x150 y70  w100 vTable2XmlLabel  ,			
	/*		Gui, MainApp:Add, Text,x260 y50 w100 , T2 Sort Field
			Gui, MainApp:Add, Edit,x260 y70  w100  ,	
			Gui, MainApp:Add, CheckBox,x380 y70  Checked, Tmp Tbl				
		*/	
			
		
			Gui, MainApp:Add, Text,x25 y100 w100 ,T1 Join Field
			Gui, MainApp:Add, Edit,x25 y120  w100 vJoinField1 ,			
			Gui, MainApp:Add, Text,x150 y100 w100 , T2 Join Field
			Gui, MainApp:Add, Edit,x150 y120  w100 vJoinField2 ,			
			Gui, MainApp:Add, Text,x260 y100 w100 , T2 Sort Field
			Gui, MainApp:Add, Edit,x260 y120  w100 vSortField2 ,	
			Gui, MainApp:Add, CheckBox,x380 y120 vXMLTempTable Checked, Tmp Tbl				
			
			
			
			Gui, MainApp:Add, Text,x25 y150 w100 ,Field 1
			Gui, MainApp:Add, Edit,x25 y170  w100 vXMLItem1 ,			
			Gui, MainApp:Add, Text,x150 y150 w100 ,Data Type
			Gui, MainApp:Add, ComboBox,x150 y170 w100  vXMLDataType1 ,Money|Number|Text
			Gui, MainApp:Add, Text,x260 y150 w100 ,Delimiter
			Gui, MainApp:Add, ComboBox,x260 y170 w100  vXMLDelimiter1 , |-|`,|_|--		;`
			Gui, MainApp:Add, CheckBox,x380 y170 vXMLExtension1 , Ext		
			
			
			
			Gui, MainApp:Add, Text,x25 y200   w100 ,Field 2
			Gui, MainApp:Add, Edit,x25 y220   w100 vXMLItem2 ,
			Gui, MainApp:Add, Text,x150 y200 w100 ,Data Type
			Gui, MainApp:Add, ComboBox,x150 y220 w100  vXMLDataType2 ,Money|Number|Text
			Gui, MainApp:Add, Text,x260 y200 w100 ,Delimiter
			Gui, MainApp:Add, ComboBox,x260 y220 w100  vXMLDelimiter2 , |-|`,|_|--		;`
			Gui, MainApp:Add, CheckBox,x380 y220 vXMLExtension2 , Ext		
			
			
			Gui, MainApp:Add, Text,x25 y250 w100 ,Field 3
			Gui, MainApp:Add, Edit,x25 y270 w100 vXMLItem3 ,
			Gui, MainApp:Add, Text,x150 y250 w100 ,Data Type
			Gui, MainApp:Add, ComboBox,x150 y270 w100  vXMLDataType3 ,Money|Number|Text
			Gui, MainApp:Add, Text,x260 y250 w100 ,Delimiter
			Gui, MainApp:Add, ComboBox,x260 y270 w100  vXMLDelimiter3 , |-|`,|_|--		;`
			Gui, MainApp:Add, CheckBox,x380 y270 vXMLExtension3 , Ext		
			
			
			Gui, MainApp:Add, Text,x25 y300 w100 ,Field 4
			Gui, MainApp:Add, Edit,x25 y320 w100 vXMLItem4 ,
			Gui, MainApp:Add, Text,x150 y300 w100 ,Data Type
			Gui, MainApp:Add, ComboBox,x150 y320 w100  vXMLDataType4 ,Money|Number|Text
			Gui, MainApp:Add, Text,x260 y300 w200 ,Delimiter
			Gui, MainApp:Add, ComboBox,x260 y320 w100  vXMLDelimiter4 , |-|`,|_|--		;`
			Gui, MainApp:Add, CheckBox,x380 y320 vXMLExtension4 , Ext		
			
			
			Gui, MainApp:Add, Text,x25 y350 w100 ,Field 5
			Gui, MainApp:Add, Edit,x25 y370  w100 vXMLItem5 ,
			Gui, MainApp:Add, Text,x150 y350 w100 ,Data Type
			Gui, MainApp:Add, ComboBox,x150 y370 w100  vXMLDataType5 ,Money|Number|Text
			Gui, MainApp:Add, Text,x260 y350 w100 ,Delimiter
			Gui, MainApp:Add, ComboBox,x260 y370 w100  vXMLDelimiter5 , |-|`,|_|--		;`
			Gui, MainApp:Add, CheckBox,x380 y370 vXMLExtension5 , Ext		
			


			
			Gui, MainApp:Add, Button,x25 y400 w335 gXMLPathExecute, XML Path




Gui, MainApp:Margin , 5, 5

Gui, MainApp:Show,  Center ,SQL Shortcuts

WinSet, Style, -0x30000, A ; WS_MAXIMIZEBOX 0x10000 + WS_MINIMIZEBOX 0x20000
RETURN ;  ; End of auto-execute section. The script is idle until the user does something.


SetTable1XmlLabel:
{
	Gui, MainApp:Submit, NoHide
	GuiControl,MainApp:,Table1XmlLabel,% Table1
}
RETURN ;

SetTable2XmlLabel:
{
	Gui, MainApp:Submit, NoHide
	GuiControl,MainApp:,Table2XmlLabel,% Table2
}
RETURN ;

Clear2:
{
		
	GUI, MainApp:Submit, NoHide	
		GuiControl, MainApp:,Table2A,
		GuiControl, Choose,Table2, 0
		
		GuiControl,MainApp:,ReverseLinkDirection,0
		GuiControl,MainApp:,LeftJoin,0
		GuiControl,MainApp:,HasLinkExtension,0		
		GuiControl,MainApp:,HasExtension,0
		
}
RETURN ;


2to1:
{
		
	GUI, MainApp:Submit, NoHide	
		GuiControl, Choose,Table1,%Table2%
		GuiControl, MainApp:,Table1A,%Table2A%
		GuiControl, Choose,Table2, 0
		GuiControl, MainApp:,Table2A,
		
		If (HasExtension) {
			GuiControl,MainApp:,HasExtension1,1
			} ELSE {
				GuiControl,MainApp:,HasExtension1,0
				
					}
		GuiControl,MainApp:,HasExtension,0

}
RETURN ;

1to2:
{
	
	GUI, MainApp:Submit, NoHide	
		GuiControl, Choose,Table2,%Table1%
		GuiControl, MainApp:, Table2A, %Table1A%
		GuiControl, Choose,Table1, 0
		GuiControl, MainApp:,Table1A,
		
		If (HasExtension1) {
			GuiControl,MainApp:,HasExtension,1
			} ELSE {
				GuiControl,MainApp:,HasExtension,0
				
					}
		GuiControl,MainApp:,HasExtension1,0

}		
RETURN ;


AboutHandler:
{
	Global vAppVersion
	Global vAppUpdateDate
	vAppVersion = % vAppVersion[1]
	vAppUpdateDate = % vAppUpdateDate[1]
	msgbox, 0x40040, About SQL Shortcuts, Author:`tCarlos A Garcia II`n`tMyEmail@Mail.com `nVersion:`t%vAppVersion% `nPublished:`t%vAppUpdateDate% `t
}
RETURN ;

HelpHandler:
{
	
	Gui, HelpUI:New, +AlwaysOnTop
	Gui, HelpUI:Font, s11
	HelpText = "Help Text Was Removed"


	Gui, HelpUI:Add, Edit,h480 w625 ReadOnly,%HelpText%
	Gui, HelpUI:Show,h500,Help: Table References
	}
RETURN ;

HideShowHandler:
{
	Global vShotAHK
	if (vShotAHK[1] = 1) {
			GUI, MainApp:Hide
			vShotAHK[1] := 0
		}
	else{
			GUI, MainApp:Show
			vShotAHK[1] := 1
	}
	Gui, HelpUI:Destroy
			
}
RETURN ;

CloseApp:
	{
		ExitApp
	}
RETURN ;

/*
		BEGIN Function Block
*/

fCountDown() {
	i = 5
	loop %i% {
		GuiControl, MainApp: ,Counter,Starting in %i% 
		sleep, 650 
		i -= 1
		}
	GuiControl, MainApp:,Counter, Writing SQL
	
}
RETURN ;

ToolTipMouseFollow:
{
	
	Tooltip, Click Where to Type SQL , % MouseX -50 ,% MouseY - 20,1
	}
RETURN	;

fStartTyping() {
		GuiControl, MainApp:,Counter, Click Where to Type SQL
		Sleep, 1000
		Settimer, ToolTipMouseFollow, 125		; 86400000
		Tooltip, Click Where to Type SQL, % MouseX -50 ,% MouseY - 20,1
		KeyWait, LButton, T30 L D
		Settimer, ToolTipMouseFollow, Off
		Sleep, 500
		ToolTip,,,,
		IF ErrorLevel
			{
				GuiControl, MainApp:,Counter, Cancelled Due to Inaction
				Exit
			}
			
		GuiControl, MainApp:,Counter, Writing SQL
	
}
RETURN	;


fCompleteNotice() {
	GuiControl, MainApp:,Counter, Write Complete
	sleep, 10000
	GuiControl, MainApp:,Counter, Ready
	
	
	
}
RETURN ;



fHasExtension(vTable,vTableA,vHasExtension := False,vLeftJoin := False){


		if (vHasExtension) {
			if (vLeftJoin) {
				send {TAB}LEFT JOIN %vTable%_EXT1 %vTableA%E1 {ENTER}
			} ELSE {
				send {TAB}INNER JOIN %vTable%_EXT1 %vTableA%E1 {ENTER}
			}	

			send {TAB}ON %vTableA%.oid = %vTableA%E1.oid {ENTER}{TAB 2}
			send AND %vTableA%E1.EFFECTIVE_Date = ({ENTER}
			SEND {TAB 8}SELECT MAX(%vTableA%E1.EFFECTIVE_Date) {ENTER}
			SEND FROM %vTable%_EXT1 %vTableA%E1 {ENTER}
			SEND WHERE {ENTER}
			SEND {TAB}%vTableA%E1.OID = %vTableA%.OID {ENTER}
			SEND and {ENTER}
			SEND %vTableA%E1.EFFECTIVE_Date <= @ExpiryDate	{ENTER}
			
			setkeydelay, 50
			SEND {SHIFT DOWN}{TAB 3}{SHIFT UP}){ENTER 2}
			
			SEND {SHIFT DOWN}{TAB 6}{SHIFT UP}
			
			
		}
		
			
}
RETURN	;


fHasLinkExtension(vLinkTableName,vLinkTableNameA,vHasLinkExtension := False,vLeftJoin := False){

	IF (vHasLinkExtension) {

				
					
					if (vLeftJoin) {
						send {TAB}LEFT JOIN %vLinkTableName%_EXT1  %vLinkTableNameA%E1  {ENTER}
					} ELSE {
						send {TAB}INNER JOIN %vLinkTableName%_EXT1 %vLinkTableNameA%E1 {ENTER}
					}	

					send {TAB}ON %vLinkTableNameA%E1.oid = %vLinkTableNameA%E1.oid {ENTER}{TAB 2}
					send AND %vLinkTableNameA%E1.EFFECTIVE_Date = ({ENTER}
					SEND {TAB 8}SELECT MAX(%vLinkTableNameA%E1.EFFECTIVE_Date) {ENTER}
					SEND FROM %vLinkTableName%_EXT1 %vLinkTableNameA%E1 {ENTER}
					SEND WHERE {ENTER}
					SEND {TAB}%vLinkTableNameA%E1.OID = %vLinkTableNameA%.OID {ENTER}
					SEND and {ENTER}
					SEND %vLinkTableNameA%E1.EFFECTIVE_Date <= @ExpiryDate	{ENTER}
					
					setkeydelay, 50
					SEND {SHIFT DOWN}{TAB 3}{SHIFT UP}){ENTER 2}
					
					SEND {SHIFT DOWN}{TAB 6}{SHIFT UP}
			}
			
		
			
}
RETURN	;


fTableConditions(vTable,vTableA,vHasExtension := False){
	
	setkeydelay, 25
		SEND {ENTER 2}
	
	IF (vTable = "ASMTACCT") {
			SEND AND %vTableA%.SUBTYPE LIKE '`%' {+} @PropertyType {ENTER}
			IF (vHasExtension) {
						SEND AND %vTableA%E1.STATUS IN ('A'`,'ACT') 	;	`
				}
		}
	
	IF (vTable = "ROLL") {
		
			SEND AND R.ROLLTYPE LIKE '`%' {+} @RollType  {ENTER}
			SEND AND R.ROLLSTATUS IN ('A'`,'ACT') 	;	`
		}
	
	IF (vTable = "BILL") {
	
			SEND AND %vTableA%E1.STATUS IN ('A'`,'ACT') 		;	`
		}
	
	SEND {ENTER 2}
	
}
RETURN	;


/*
		END Function Block
*/


SQLHeaderExecute:
{
	
	GUI, MainApp:Submit, NoHide	
	fStartTyping()
	setkeydelay, 50
	
	send SET DEADLOCK_PRIORITY -10{ENTER}
	send GO {ENTER 2}
	
	setkeydelay, 42
	SEND DECLARE{ENTER 2}
	
	SEND -- CONFIGURE BELOW {ENTER 2}
	
	SEND {TAB 3}@TaxCycle{TAB 2}INT{TAB 3}={TAB}%TaxCycle%{ENTER}{SHIFT DOWN}{TAB}{SHIFT UP}

;  Config Options

	Loop, Parse, TablesIncluded, |
		{
			;	MsgBox Selection number %A_Index% is %A_LoopField%.
				IF (InStr(A_LoopField,"AsmtAcct")) {
						send {,}{TAB}@PropertyType{TAB 1}NVARCHAR(2){TAB}={TAB}'RP'{TAB 2}-- PropertyType RP, PP, MH{ENTER}
					
				}
				IF (InStr(A_LoopField,"Bill")) {
						send {,}{TAB}@BillingCode{TAB 1}NVARCHAR(1){TAB}={TAB}'D'{TAB 3}-- D = Direct Deposite, E = Escrow, I = Individual{ENTER}
						send {,}{TAB}@BalanceDue{TAB 2}Money{TAB 2}={TAB}0{TAB 3}-- Minimum Balance Due{ENTER}
					
				}
				IF (InStr(A_LoopField,"Roll")) {
						send {,}{TAB}@RollType{TAB 2}NVARCHAR(1){TAB}={TAB}'F'{TAB 3}-- RollType F,P,A,S,O{ENTER}
					
				}
				IF (InStr(A_LoopField,"UTA")) {
						send {,}{TAB}@UTAID{TAB 3}NVARCHAR(30){TAB}={TAB}''{TAB 3}-- UTA ID{ENTER}
					
				}


			
		}

	
	
	
	SEND {ENTER 2}{SHIFT DOWN}{TAB 20}{SHIFT UP}
	SEND -- END CONFIGURATION {ENTER 2}


	SEND -- DO NOT ALTER BELOW THIS LINE{ENTER}{TAB 2}
	send {,}{TAB}@ExpiryDate{TAB 2}DATETIME{TAB}={TAB} getDate(){ENTER}
	
	
	if (RollTableLimited) {
			SEND {,}{TAB}@EffectiveDate{TAB}DATETIME{ENTER}
			
			SEND {,}{TAB}@BillYear{TAB 2}INT{ENTER}
			setkeydelay, 50
			SEND {SHIFT DOWN}{TAB 20}{SHIFT UP}{ENTER 2}
			setkeydelay, 25
			
			SENDRAW /
			SENDRAW *
			SEND {ENTER}
			SEND SELECT {ENTER}
			SEND {TAB 3}@EffectiveDate =  min(LRMS_ROLLYEAR_MANAGEMENT.START_DATE) {ENTER}{SHIFT DOWN}{TAB}{SHIFT UP}
			SEND {,}{TAB}@ExpiryDate = max(LRMS_ROLLYEAR_MANAGEMENT.END_DATE) {ENTER}
			SEND {,}{TAB}@BillYear = max(ROLL_YEAR_LKP.BILLYEAR) {ENTER}{SHIFT DOWN}{TAB}{SHIFT UP}
			SEND from ROLL_YEAR_LKP {ENTER}
			SEND {TAB}inner join LRMS_ROLLYEAR_MANAGEMENT {ENTER}
			SEND {TAB 2}ON LRMS_ROLLYEAR_MANAGEMENT.ROLLYEAR = ROLL_YEAR_LKP.ROLLYEAR {ENTER}
			SEND {TAB}AND ROLL_YEAR_LKP.CODE = @TaxCycle {TAB}
			SENDRAW -- *
			SENDRAW /
			SEND {ENTER}--{TAB}Uncomment the above section to limit results to a specific Tax Cycle
			SEND {ENTER 3}{SHIFT DOWN}{TAB 15}{SHIFT UP}
	
		}

	if (TmpTableRemove){
				SENDRAW -- CLEAR TEMP TABLES
				SEND {ENTER 2}
				setkeydelay, 60
				SEND DECLARE{TAB 2}
				SENDRAW @tSQL
				SEND {TAB 2}NVARCHAR(MAX){TAB}={TAB}''{ENTER}
				SENDRAW SELECT @tSQL = @tSQL + 'DROP TABLE ' + QUOTENAME(NAME) + ';' + CHAR(13) + CHAR(10)
				SEND {ENTER 2}{TAB 1}FROM TEMPDB..SYSOBJECTS {ENTER}
				SEND {TAB 1}WHERE NAME LIKE 
				SENDRAW '#`%'
				SEND {TAB}
				SENDRAW AND OBJECT_ID('TEMPDB..' + QUOTENAME(NAME)) IS NOT NULL
				SEND {ENTER 2}
				SENDRAW IF @tSQL <> ''
				SEND {ENTER}BEGIN{ENTER}{TAB}
				SENDRAW EXEC SP_SQLEXEC @tSQL
				SEND {ENTER 2}{SHIFT DOWN}{TAB 20}{SHIFT UP}{TAB 2}END {ENTER 3}
	
		}
	
	
	SEND {ENTER 4}
	SEND {UP}{RIGHT}
	SEND SELECT DISTINCT * {ENTER 6} {UP 4}
	
	fCompleteNotice()
}
RETURN ;





NotInTableExecute:
{
	
	
	GUI, MainApp:Submit, NoHide	
	
	StringUpper, Table1, Table1
	StringUpper, Table2, Table2		
	StringUpper, Table1A, Table1A
	StringUpper, Table2A, Table2A
	

	
	if (ReverseLinkDirection) {
				LinkTableName = %Table2%_%Table1%_LINK
				LinkTableNameAbbr = %Table2A%%Table1A%L
			}	ELSE {
				LinkTableName = %Table1%_%Table2%_LINK
				LinkTableNameAbbr = %Table1A%%Table2A%L
			}
	
	fStartTyping()
	

	if (NotIn) {
			SEND -- Not In %Table2% {ENTER}
			SEND {TAB}%Table1A%.OID NOT IN ( {ENTER} 
		} ELSE {			
			SEND -- In %Table2% {ENTER}
			SEND {TAB}%Table1A%.OID IN ( {ENTER} 
		}
		
	if (ReverseLinkDirection) {
				SEND {TAB 5}SELECT %LinkTableNameAbbr%.RIGHTOID{ENTER}
		}	ELSE {
				SEND {TAB 5}SELECT %LinkTableNameAbbr%.LEFTOID{ENTER}
		}	
	

	SEND {TAB}FROM %LinkTableName% %LinkTableNameAbbr% {ENTER}
/*	
	
	
	if (ReverseLinkDirection) {
		send {TAB}ON %Table1A%.oid = %LinkTableNameAbbr%.RIGHTOID {ENTER}
	} ELSE {
		send {TAB}ON %Table1A%.oid = %LinkTableNameAbbr%.LEFTOID {ENTER}
	}

*/
	
	send {TAB 2}

	setkeydelay, 50
	send {SHIFT DOWN}{TAB 3}{SHIFT UP}
	setkeydelay, 25
	
	SEND {TAB 3}INNER JOIN %Table2%  %Table2A% {ENTER}
	
	if (ReverseLinkDirection) {
				send {TAB}ON %Table2A%.oid = %LinkTableNameAbbr%.LEFTOID {ENTER}
			} ELSE {
				send {TAB}ON %Table2A%.oid = %LinkTableNameAbbr%.RIGHTOID {ENTER}
			}	
	send {TAB 2}
	
	
;	SEND AND %LinkTableNameAbbr%.EFFECTIVE_Date >= @EffectiveDate {ENTER}
	send AND %LinkTableNameAbbr%.EFFECTIVE_Date <= @ExpiryDate {ENTER}
	send AND isNull(%LinkTableNameAbbr%.EXPIRY_DATE,@ExpiryDate) >= @ExpiryDate {ENTER 2}
	
	
;	SEND AND %Table2A%.EFFECTIVE_Date >= @EffectiveDate {ENTER}
	send AND %Table2A%.EFFECTIVE_Date <= @ExpiryDate {ENTER}
	send AND isNull(%Table2A%.EXPIRY_DATE,@ExpiryDate) >= @ExpiryDate {ENTER}
	
	setkeydelay, 50
	send {SHIFT DOWN}{TAB 4}{SHIFT UP}{)}{ENTER 4}
	setkeydelay, 25
	
	fCompleteNotice()
	
}
RETURN ;





FirstTableExecute:
{

	
	GUI, MainApp:Submit, NoHide	



	StringUpper, Table1, Table1
	if (StrLen(Table1) < 3){
		TrayTip, Empty Criteria, You did not supply a Table Name
		EXIT
		}
	StringUpper, Table1A, Table1A
	if (StrLen(Table1A) < 1){
		TrayTip, Empty Criteria, You did not supply an abbreviation
		EXIT
		}		
	
	fStartTyping()
	
	send FROM %Table1% %Table1A% {SPACE}
	if (NoLock = 1) {
			sendraw with (NOLOCK)
			
		}
	SEND {ENTER}

	fHasExtension(Table1,Table1A,HasExtension1)
	fTableConditions(Table1,Table1A,HasExtension1)

	setkeydelay, 50
	SEND {SHIFT DOWN}{TAB 54}{SHIFT UP}{SPACE}{ENTER 2}
	setkeydelay, 25



	IF (StrLen(Table2) > 2) {
		SEND {SPACE}{TAB 3}
		fJoinTableExecute(Table1,Table1A,Table2,Table2A,HasExtension,ReverseLinkDirection,HasLinkExtension,LeftJoin)
		
	}
	
	
	fCompleteNotice()
	
}
RETURN	;


JoinTableOnly:
{
	
	GUI, MainApp:Submit, NoHide	
		
	fStartTyping()
		fJoinTableExecute(Table1,Table1A,Table2,Table2A,HasExtension,ReverseLinkDirection,HasLinkExtension,LeftJoin)

	fCompleteNotice()

}
RETURN	;

fJoinTableExecute(Table1,Table1A,Table2,Table2A,HasExtension,ReverseLinkDirection,HasLinkExtension,LeftJoin)
{
	
	StringUpper, Table1, Table1
	StringUpper, Table2, Table2		
	StringUpper, Table1A, Table1A
	StringUpper, Table2A, Table2A
	

	if (ReverseLinkDirection) {
		LinkTableName = %Table2%_%Table1%_LINK
		LinkTableNameAbbr = %Table2A%%Table1A%L
	}	ELSE {
		LinkTableName = %Table1%_%Table2%_LINK
		LinkTableNameAbbr = %Table1A%%Table2A%L
	}
	
	
	send -- %Table2% {ENTER}
	if (LeftJoin) {
		send {TAB}LEFT JOIN %LinkTableName% %LinkTableNameAbbr% {ENTER}	
	} ELSE {
		send {TAB}INNER JOIN %LinkTableName% %LinkTableNameAbbr% {ENTER}
	}
	
	
	
	if (ReverseLinkDirection) {
		send {TAB}ON %Table1A%.oid = %LinkTableNameAbbr%.RIGHTOID {ENTER}
	} ELSE {
		send {TAB}ON %Table1A%.oid = %LinkTableNameAbbr%.LEFTOID {ENTER}
	}
	send {TAB 2}
;	SEND AND %LinkTableNameAbbr%.EFFECTIVE_Date >= @EffectiveDate {ENTER}
	send AND %LinkTableNameAbbr%.EFFECTIVE_Date <= @ExpiryDate {ENTER}
	send AND isNull(%LinkTableNameAbbr%.EXPIRY_DATE,@ExpiryDate) >= @ExpiryDate {ENTER}

	setkeydelay, 50
	send {SHIFT DOWN}{TAB 3}{SHIFT UP}
	setkeydelay, 25
	
	fHasLinkExtension(LinkTableName,LinkTableNameAbbr,HasLinkExtension,LeftJoin)
	
	
	if (LeftJoin) {
		send {TAB}LEFT JOIN %Table2%  %Table2A% {ENTER}	
	} ELSE {
		send {TAB}INNER JOIN %Table2%  %Table2A% {ENTER}
	}
	
	if (ReverseLinkDirection) {
		send {TAB}ON %Table2A%.oid = %LinkTableNameAbbr%.LEFTOID {ENTER}
	} ELSE {
		send {TAB}ON %Table2A%.oid = %LinkTableNameAbbr%.RIGHTOID {ENTER}
	}	
	send {TAB 2}
;	SEND AND %Table2A%.EFFECTIVE_Date >= @EffectiveDate {ENTER}
	send AND %Table2A%.EFFECTIVE_Date <= @ExpiryDate {ENTER}
	send AND isNull(%Table2A%.EXPIRY_DATE,@ExpiryDate) >= @ExpiryDate {ENTER}
	
	setkeydelay, 50
	send {SHIFT DOWN}{TAB 4}{SHIFT UP}
	setkeydelay, 25
	
	fHasExtension(Table2,Table2A,HasExtension,LeftJoin)
	fTableConditions(Table2,Table2A,HasExtension)

			setkeydelay, 25
			send {ENTER 2}
			

}
RETURN	;




; Needs work... may need complete overhaul




CDLkpExecute:
{
	
	Gui, MailApp:Submit, NoHide
	
;;	InputBox, vTable1A, First Table Abbreviation, What Abbreviation will be used for Source Table?`n`nIE: AsmtAcct_Ext1 might be AAE1
	if (StrLen(Table1A) < 1){
		TrayTip, Empty Criteria, You did not supply a First Table
		EXIT
		}		

;;	InputBox, vFilter1A, Linked Table Abbreviation, What Abbreviation will be used for %vFilter1A%?`n`nIE: If using DelinquencyStatus you might use DQS
	if (StrLen(Filter1A) < 1){
		TrayTip, Empty Criteria, You did not supply a Filter Abbreviation
		EXIT
		}
		
;;	InputBox, vFilter1Code, Lookup Code, What is the Code Being Used?`n`nIE: DelinquencyStatus`, AppealHearing`, AssessmentAccountRole
	if (StrLen(Filter1Code) < 3){
		TrayTip, Empty Criteria, You did not supply a Filter Code.
		EXIT
		}
	
	
	StringUpper, vFilter1A, Filter1A		
	StringUpper, vTable1A, Table1A
	StringUpper, vFilter1Code, Filter1Code
	
	
	
	fStartTyping()
	
	send -- %vFilter1A% {ENTER}
	
	send {TAB}INNER JOIN LRMI_LKP_CD_FLT1 %vFilter1A% {ENTER}
	send {TAB}ON %vTable1A%.STATUS = %vFilter1A%.CODE{ENTER}
	send {TAB 2}AND %vFilter1A%.FILTER1 LIKE '%vFilter1Code%'{ENTER}
	setkeydelay, 50
	send {SHIFT DOWN}{TAB 4}{SHIFT UP}
	setkeydelay, 25
	send {ENTER}{ENTER}
	sendraw	-- 
	send	{SPACE}%vFilter1A%.Description{SPACE}
	sendraw <-- Copy & Paste in Select 
	send {ENTER}{ENTER}

	fCompleteNotice()
}
RETURN	;


fXMLItem(vXMLItem,vXMLDataType,vXMLDelimiter,vTableChild,vTableChildA,vXMLExtension){
	
	IF (vXMLExtension){

		vTableChildA = %vTableChildA%E1
	}
	
	if (vXMLDataType = "Money"){
		SEND {SPACE}FORMAT{(}%vTableChildA%.%vXMLItem%,{'}{$}{#},{#}{#}0.00{'}{)} {+} {'} %vXMLDelimiter% {'}  
		
	}
	if (vXMLDataType = "Number"){
	SEND {SPACE}FORMAT{(}%vTableChildA%.%vXMLItem%,{'}{#},{#}{#}{#}.{#}{#}{'}{)} {+} {'} %vXMLDelimiter% {'}
		
	}	
	
	if (vXMLDataType = "Text"){
		SEND {SPACE}CAST{(}%vTableChildA%.%vXMLItem% AS VARCHAR{)} {+} {'} %vXMLDelimiter% {'}
		
	}	
	
	
}
RETURN	;


XMLPathExecute:
{
	
		Gui, MainApp:Submit, NoHide
		
/*		
			Gui, MainApp:Add, Text,x25 y250 w100 ,Field 5
			Gui, MainApp:Add, Edit,x25 y270  w100 vXMLItem5 ,
			Gui, MainApp:Add, Text,x150 y250 w100 ,Data Type
			Gui, MainApp:Add, ComboBox,x150 y270 w100  vXMLDataType5 ,Money|Number|Text
			Gui, MainApp:Add, Text,x260 y250 w100 ,Delimiter
			Gui, MainApp:Add, ComboBox,x260 y270 w100  vXMLDelimiter5 , |-|`,|_|--		;`
			*/
			
		if (StrLen(XMLItem1) > 3) {
			
			fStartTyping()
			
			SEND {,}{Tab}{(}{Enter}{Tab}Select {ENTER}{TAB}{(}SELECT {ENTER}{TAB}CAST{(}ROW_NUMBER{(}{)} OVER{(}ORDER BY %Table2A%.%SortField2% ASC{)} AS NVARCHAR{)} {+} {'}{)} {'} {+} {ENTER}{TAB}
			fXMLItem(XMLItem1,XMLDataType1,XMLDelimiter1,Table2,Table2A,vXMLExtension1)
			
			
		}
		ELSE {
			msgbox A Field was not specified. Please enter a field and try again.
			EXIT
		}


		if (StrLen(XMLItem2) > 3) {
				SEND {ENTER}{+}{SPACE}
				fXMLItem(XMLItem2,XMLDataType2,XMLDelimiter2,Table2,Table2A,vXMLExtension2)
			}


		if (StrLen(XMLItem3) > 3) {
				SEND {ENTER}{+}{SPACE}
				fXMLItem(XMLItem3,XMLDataType3,XMLDelimiter3,Table2,Table2A,vXMLExtension3)
			}


		if (StrLen(XMLItem4) > 3) {
				SEND {ENTER}{+}{SPACE}
				fXMLItem(XMLItem4,XMLDataType4,XMLDelimiter4,Table2,Table2A,vXMLExtension4)
			}


		if (StrLen(XMLItem5) > 3) {
				SEND {ENTER}{+}{SPACE}
				fXMLItem(XMLItem5,XMLDataType5,XMLDelimiter5,Table2,Table2A,vXMLExtension5)
			}

		
		SEND {ENTER}{SHIFT DOWN}{TAB 2}{SHIFT UP}FROM{SPACE}
		IF (XMLTempTable){
				SEND {#}
			}
		
		SEND %Table2% %Table2A%{ENTER}{TAB}where %Table1A%.%JoinField1% {=} %Table2A%.%JoinField2%{ENTER}{SHIFT DOWN}{TAB}{SHIFT UP}FOR XML PATH{(}{'}{'}{)}{ENTER}{)} z{ENTER}{SHIFT DOWN}{TAB 2}{SHIFT UP}{)}{ENTER 2}
	

	
			fCompleteNotice()
}
RETURN	;

