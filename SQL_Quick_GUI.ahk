vShotAHK := Object()
vShotAHK[1] := 1

Menu, Tray, Add, Hide/Show,HideShowHandler
Menu, Tray, Default, Hide/Show
try {
		Menu, Tray, Icon, SQLIcon.ico
	}
Menu, Tray, Click, 1

Gui, MainApp:New, +AlwaysOnTop



Menu, MenuBar, Add, Hide, HideShowHandler

Menu, HelpMenu, Add, About, AboutHandler
Menu, HelpMenu, Add, Help, HelpHandler
Menu, MenuBar, Add, Help, :HelpMenu

Gui, MainApp:Menu, MenuBar

Gui, MainApp:Add, Button, w100 gSQLHeader, SQL Header
Gui, MainApp:Add, Button, w100 gFirstTable, First Table 
Gui, MainApp:Add, Button, w100 gRollFirstTable, Roll First Table

Gui, MainApp:Add, Button, w100 gJoinSQL, Join Table
Gui, MainApp:Add, Button, w100 gAddCD1Lkp, Add CD1 Lookup  

Gui, MainApp:Add, Text,vCounter,

Gui, MainApp:Add, Button, w100 gCloseApp, Close App  

Gui, MainApp:Margin , 5, 5

Gui, MainApp:Show, w120 Center ,SQL Shortcuts

WinSet, Style, -0x30000, A ; WS_MAXIMIZEBOX 0x10000 + WS_MINIMIZEBOX 0x20000
RETURN ;  ; End of auto-execute section. The script is idle until the user does something.

AboutHandler:
{
	msgbox, 0x40040, About SQL Shortcuts, Written by Carlos A Garcia II`ncarlos.garcia2@conduent.com `n
}
RETURN ;

HelpHandler:
{
	
	Gui, HelpUI:New, +AlwaysOnTop
	Gui, HelpUI:Font, s11
	HelpText = "Help Text went here but has been removed"


	Gui, HelpUI:Add, Edit,h480 w550 ReadOnly,%HelpText%
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

fCountDown() {
	i = 5
	loop %i% {
		GuiControl, MainApp: ,Counter,%i% " sec"
		sleep, 650 
		i -= 1
		}
	GuiControl,,Counter,
	
}
RETURN ;

FirstTable: ;This will give the Standard Table with Extension
{
	setkeydelay, 25

	InputBox, Table1, Base Table, What is the Table
	StringUpper, Table1, Table1
	if (StrLen(Table1) < 3){
		TrayTip, Empty Criteria, You did not supply a Table Name
		EXIT
		}
	InputBox, Table1A, Base Table Abbreviation, What Abbreviation will be used for %Table1%?
	StringUpper, Table1A, Table1A
	if (StrLen(Table1A) < 1){
		TrayTip, Empty Criteria, You did not supply an abbreviation
		EXIT
		}		
	
	fCountDown()
	
	send FROM %Table1% %Table1A% {ENTER}
	
	send {TAB}INNER JOIN %Table1%_EXT1 %Table1A%E1 {ENTER}
	send {TAB}ON %Table1A%.oid = %Table1A%E1.oid {ENTER}
	send {TAB}{TAB}
;	send AND %Table1A%.EFFECTIVE_Date >= @EffectiveDate {ENTER}
	send AND %Table1A%.EFFECTIVE_Date <= @ExpiryDate {ENTER}
	send AND isNull(%Table1A%.EXPIRY_DATE,@ExpiryDate) >= @ExpiryDate {ENTER}{ENTER}


	send AND %Table1A%E1.EFFECTIVE_Date = ( {ENTER}
;	SELECT MAX(%Table1A%E1.EFFECTIVE_Date) FROM %Table1A%E1 %Table1A%E1 WHERE %Table1A%E1.OID = %Table1A%.OID and %Table1A%E1.EFFECTIVE_Date <= @ExpiryDate){ENTER}
	SEND {TAB 8}SELECT MAX(%Table1A%E1.EFFECTIVE_Date) {ENTER}
	SEND FROM %Table1%_EXT1 %Table1A%E1 {ENTER}
	SEND WHERE {ENTER}
	SEND {TAB}%Table1A%E1.OID = %Table1A%.OID {ENTER}
	SEND and {ENTER}
	SEND %Table1A%E1.EFFECTIVE_Date <= @ExpiryDate	{ENTER}
	setkeydelay, 50
	SEND {SHIFT DOWN}{TAB 3}{SHIFT UP}){ENTER}
	setkeydelay, 25
	send {ENTER}{ENTER}

	send {ENTER}{ENTER}
	
	}
RETURN ;

SQLHeader: ;This will give the Standards Options/Variables for Scripts
{
	fCountDown()
	setkeydelay, 50
	
	send SET DEADLOCK_PRIORITY -10{ENTER}
	send GO {ENTER 2}
	
	setkeydelay, 42
	SEND -- Setting Effective Dates and Tax Cycle{ENTER}
	
	SEND DECLARE{TAB}{TAB}@TaxCycle{TAB}INT{TAB}={TAB}1718{ENTER}
	send {TAB}{TAB},{TAB}@ExpiryDate{TAB}DATETIME{TAB}={TAB} getDate(){ENTER}
	SEND {,}{TAB}@EffectiveDate{TAB}DATETIME{ENTER}
	setkeydelay, 50
	SEND {SHIFT DOWN}{TAB 20}{SHIFT UP}{ENTER 2}
	setkeydelay, 25
	
	SENDRAW /
	SENDRAW *
	SEND {ENTER}
	SEND SELECT {ENTER}
	SEND {TAB}{TAB}{TAB}@EffectiveDate =  min(LRMS_ROLLYEAR_MANAGEMENT.START_DATE) {ENTER}{SHIFT DOWN}{TAB}{SHIFT UP}
	SEND {,}{TAB}@ExpiryDate = max(LRMS_ROLLYEAR_MANAGEMENT.END_DATE) {ENTER}{SHIFT DOWN}{TAB}{SHIFT UP}
	SEND from ROLL_YEAR_LKP {ENTER}
	SEND {TAB}inner join LRMS_ROLLYEAR_MANAGEMENT {ENTER}
	SEND {TAB}{TAB}ON LRMS_ROLLYEAR_MANAGEMENT.ROLLYEAR = ROLL_YEAR_LKP.ROLLYEAR {ENTER}
	SEND {TAB}AND ROLL_YEAR_LKP.CODE = @TaxCycle {TAB}
	SENDRAW -- *
	SENDRAW /
	SEND {ENTER}--{TAB}Uncomment the above section to limit results to a specific Tax Cycle
	SEND {ENTER 3}
	SEND SELECT DISTINCT * {ENTER}
	
	

	}
RETURN ;


RollFirstTable: ;This will give the Roll Query with basic criteria
{
	setkeydelay, 25

	
	
	fCountDown()
	
	send FROM roll r {ENTER 4}
	
	
	
	SEND {TAB}WHERE {ENTER}
	SEND {TAB}R.STATUS = 'A' {ENTER}
	SEND and{TAB}R.ROLLSTATUS IN ('A', 'AE','AS') {ENTER}
	SEND and{TAB}R.TYPE = 'F' {ENTER}
	SEND and{TAB}R.YEAR = @TaxCycle {ENTER}


	setkeydelay, 50
	SEND {SHIFT DOWN}{TAB 3}{SHIFT UP}{ENTER}
	setkeydelay, 25
	send {ENTER}{ENTER}

	
	}
RETURN ;

AddCD1Lkp: ;This will give the LRMI_LKP_CD_FLT1
{
	
	setkeydelay, 25
	TrayTip, Started, SQL AutoScript

	InputBox, Table1A, First Table Abbreviation, What Abbreviation will be used for Source Table?`n`nIE: AsmtAcct_Ext1 might be AAE1
	if (StrLen(Table1A) < 1){
		TrayTip, Empty Criteria, You did not supply a First Table
		EXIT
		}		
		
	InputBox, Code1, Lookup Code, What is the Code Being Used?`n`nIE: DelinquencyStatus`, AppealHearing`, AssessmentAccountRole
	if (StrLen(Code1) < 3){
		TrayTip, Empty Criteria, You did not supply a Table to be Linked.
		EXIT
		}
	InputBox, Code1A, Linked Table Abbreviation, What Abbreviation will be used for %Code1%?`n`nIE: If using DelinquencyStatus you might use DQS
	if (StrLen(Code1A) < 1){
		TrayTip, Empty Criteria, You did not supply a First Table
		EXIT
		}	
		
		
	StringUpper, Code1, Code1		
	StringUpper, Table1A, Table1A
	StringUpper, Code1A, Code1A
	
	
	
	fCountDown()
	
	send -- %Code1% {ENTER}
	
	send {TAB}INNER JOIN LRMI_LKP_CD_FLT1 %Code1A% {ENTER}
	send {TAB}ON %Table1A%.STATUS = %Code1A%.CODE{ENTER}
	send {TAB}{TAB}AND %Code1A%.FILTER1 LIKE '%Code1%'{ENTER}
	setkeydelay, 50
	send {SHIFT DOWN}{TAB 4}{SHIFT UP}
	setkeydelay, 25
	send {ENTER}{ENTER}
	sendraw --	select distinct filter1 from LRMI_LKP_CD_FLT1
	send {ENTER}{ENTER}
	
	
	}
RETURN ;


JoinSQL: ;This will give the standard linking I use in queries
{


	setkeydelay, 25
	TrayTip, Started, SQL AutoScript

	InputBox, Table1, First Table, What is the First Table
	if (StrLen(Table1) < 3){
		TrayTip, Empty Criteria, You did not supply a First Table
		EXIT
		}
	InputBox, Table1A, First Table Abbreviation, What Abbreviation will be used for %Table1%?
	if (StrLen(Table1A) < 1){
		TrayTip, Empty Criteria, You did not supply a First Table
		EXIT
		}		
		
	InputBox, Table2, Table to Link, What is the Table being Linked?
	if (StrLen(Table2) < 3){
		TrayTip, Empty Criteria, You did not supply a Table to be Linked.
		EXIT
		}
	InputBox, Table2A, Linked Table Abbreviation, What Abbreviation will be used for %Table2%?
	if (StrLen(Table2A) < 1){
		TrayTip, Empty Criteria, You did not supply a First Table
		EXIT
		}	
		
		
	StringUpper, Table1, Table1
	StringUpper, Table2, Table2		
	StringUpper, Table1A, Table1A
	StringUpper, Table2A, Table2A
	
	inputBox, LinkDirection, First Table, Which table is on the left side of the link?`n 1 = %Table1%_%Table2%_LINK (Default) `n 2 = %Table2%_%Table1%_LINK

	if (LinkDirection = 2) {
		LinkTableName = %Table2%_%Table1%_LINK
		LinkTableNameAbbr = %Table2A%%Table1A%L
	}	ELSE {
		LinkTableName = %Table1%_%Table2%_LINK
		LinkTableNameAbbr = %Table1A%%Table2A%L
	}
	
	
	inputBox, JoinType,	Join Type, What type of Join will be used?`n I = Inner Join (Default) `n L = Left Join`n
	
	StringUpper, JoinType, JoinType
	
	fCountDown()
	
	send -- %Table2% {ENTER}
	if (JoinType = "L") {
		send {TAB}LEFT JOIN %LinkTableName% %LinkTableNameAbbr% {ENTER}	
	} ELSE {
		send {TAB}INNER JOIN %LinkTableName% %LinkTableNameAbbr% {ENTER}
	}
	
	
	
	if (LinkDirection = 2) {
		send {TAB}ON %Table1A%.oid = %LinkTableNameAbbr%.RIGHTOID {ENTER}
	} ELSE {
		send {TAB}ON %Table1A%.oid = %LinkTableNameAbbr%.LEFTOID {ENTER}
	}
	send {TAB}{TAB}
;	SEND AND %LinkTableNameAbbr%.EFFECTIVE_Date >= @EffectiveDate {ENTER}
	send AND %LinkTableNameAbbr%.EFFECTIVE_Date <= @ExpiryDate {ENTER}
	send AND isNull(%LinkTableNameAbbr%.EXPIRY_DATE,@ExpiryDate) >= @ExpiryDate {ENTER}

	setkeydelay, 50
	send {SHIFT DOWN}{TAB 3}{SHIFT UP}
	setkeydelay, 25
	
	if (JoinType = "L") {
		send {TAB}LEFT JOIN %Table2%  %Table2A% {ENTER}	
	} ELSE {
		send {TAB}INNER JOIN %Table2%  %Table2A% {ENTER}
	}
	
	if (LinkDirection = 2) {
		send {TAB}ON %Table2A%.oid = %LinkTableNameAbbr%.LEFTOID {ENTER}
	} ELSE {
		send {TAB}ON %Table2A%.oid = %LinkTableNameAbbr%.RIGHTOID {ENTER}
	}	
	send {TAB}{TAB}
;	SEND AND %Table2A%.EFFECTIVE_Date >= @EffectiveDate {ENTER}
	send AND %Table2A%.EFFECTIVE_Date <= @ExpiryDate {ENTER}
	send AND isNull(%Table2A%.EXPIRY_DATE,@ExpiryDate) >= @ExpiryDate {ENTER}
	
	setkeydelay, 50
	send {SHIFT DOWN}{TAB 4}{SHIFT UP}
	setkeydelay, 25
	
	if (JoinType = "L") {
		send {TAB}LEFT JOIN %Table2%_EXT1 %Table2A%E1 {ENTER}
	} ELSE {
		send {TAB}INNER JOIN %Table2%_EXT1 %Table2A%E1 {ENTER}
	}	

	send {TAB}ON %Table2A%.oid = %Table2A%E1.oid {ENTER}
	send AND %Table2A%E1.EFFECTIVE_Date = ({ENTER}
	SEND {TAB 8}SELECT MAX(%Table2A%E1.EFFECTIVE_Date) {ENTER}
	SEND FROM %Table2%_EXT1 %Table2A%E1 {ENTER}
	SEND WHERE {ENTER}
	SEND {TAB}%Table2A%E1.OID = %Table2A%.OID {ENTER}
	SEND and {ENTER}
	SEND %Table2A%E1.EFFECTIVE_Date <= @ExpiryDate	{ENTER}
	setkeydelay, 50
	SEND {SHIFT DOWN}{TAB 3}{SHIFT UP}){ENTER}
	setkeydelay, 25
	send {ENTER}{ENTER}
	
	}
	
RETURN ;


