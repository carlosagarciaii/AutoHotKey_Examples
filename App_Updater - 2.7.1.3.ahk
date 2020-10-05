#SingleInstance  Force
uTime := "20191104"
uAppXNTargetPath = ""
vEnvironmentName := ""
vEnvironmentNameDefaultOption := ""
vresourceServerAppPath := ""
vresourceServerDMPath := ""
vIconName := ""
vTargetService := ""
				
GUITitleVersion = AppXN Updater - 2.7.1.3
EnvGet, vFileInstallPath, LOCALAPPDATA
EnvGet, vSplashImagePath, LOCALAPPDATA
vSplashImagePath = %vSplashImagePath%\AppXN_Updater
FileInstall, C:\Users\SomeUser\Documents\AutoHotKey Scripts\AppXN Updater\Hold_Up.png,%vSplashImagePath%\Hold_Up.png,0


/*
	Environment Setup
*/

		IF !InStr(FileExist(vFileInstallPath),"D") {
				FileCreateDir, %vFileInstallPath%	
				}
		vFileInstallPath = %vFileInstallPath%\_AppXN
				


		vProjectCatalogPath = %vFileInstallPath%\Export\projectCatalog.xml
		IF !FileExist(vProjectCatalogPath) {
				SplashImage , %vSplashImagePath%\Hold_Up.png, h329, w489,,,Loading Key Files 

				FileInstall, C:\Users\SomeUser\Documents\AutoHotKey Scripts\AppXN Updater\Export.zip,%vFileInstallPath%\Export.zip,0
				RunWait, Powershell.exe "expand-archive -LiteralPath '%vFileInstallPath%\Export.zip' -Destination '%vFileInstallPath%\'", ,Hide
			}

		SplashImage, Off
				

/*
	Set INI Path
*/

		EnvGet, vAppXNUpdaterCorePath, LOCALAPPDATA
		vAppXNUpdaterCorePath = %vAppXNUpdaterCorePath%\AppXN_Updater
		If !InStr(FileExist(vAppXNUpdaterCorePath),"D")	{
				FileCreateDir, %vAppXNUpdaterCorePath%
			}



/*
	Generate New INI File if 
*/
		IniFile = %vAppXNUpdaterCorePath%\AppXN_Updater.ini
		
		IniRead, tvEnvironmentCount,% IniFile, Environments,EnvironmentCount,0	
		if (tvEnvironmentCount = 0) {
				FileDelete, % IniFile
			}
			
		
Menu, Tray, Add, Hide/Show,HideShowHandler
	Menu, Tray, Default, Hide/Show


		FileInstall, C:\Users\SomeUser\Documents\AutoHotKey Scripts\AppXN Updater\AppXN_Updater.ico,%vAppXNUpdaterCorePath%\AppXN_Updater.ico,0
		
try {
		Menu, Tray, Icon, % vAppXNUpdaterCorePath "\AppXN_Updater.ico"
	}
	
Menu, Tray, Click, 1

Gui, MainApp:New, +AlwaysOnTop


Gui,MainApp:Add,Tab3,x5 y5 h370 vTabMaster, Update|Manage Environments

GuiControlGet, TabMaster, POS

/*
	Update Tab
*/

	

	Gui,MainApp:Tab, Update
	
	Loop 
		{
			if (	FileExist(IniFile)	)
				{
					
					ListBoxItems := "ALL|"

						IniRead, tEnvironmentCount,% IniFile, Environments,EnvironmentCount,0
						i=1
						while (i <= tEnvironmentCount)
						{
							IniRead,tEnvironmentName,% IniFile, % "EnvironmentNo" i, IconName
							ListBoxItems := ListBoxItems tEnvironmentName "|"
							i += 1
						}
						
					Gui, MainApp:Add, ListBox, % " x" TabMasterX + 10 " y" TabMasterY + 30 " w125 h180 vEnvironmentList altsubmit ", % ListBoxItems

						GuiControl,Choose, EnvironmentList, 2
					
					
					
					
					IniRead, vChkState, %IniFile%, Settings, GetZip , 1
					if (vChkState = 1)
						{
							Gui, MainApp:Add, CheckBox, % " x" TabMasterX + 150 " y" TabMasterY + 30 " vGetZip Checked", Get Zip (Faster)
						}
					Else
						{
							Gui, MainApp:Add, CheckBox,  % " x" TabMasterX + 150 " y" TabMasterY + 30 " vGetZip " , Get Zip (Faster)
						}
					IniRead, vChkState, %IniFile%, Settings, ClearCache , 1
					if (vChkState = 1)
						{
							Gui, MainApp:Add, CheckBox, vClearCache Checked, Clear Cache

						}
					Else
						{
							Gui, MainApp:Add, CheckBox, vClearCache, Clear Cache
						}
							
					IniRead, vChkState, %IniFile%, Settings, ClearOldFiles , 1
					if (vChkState = 1)
						{
							Gui, MainApp:Add, CheckBox, vClearOldFiles Checked, Clear Old Files			
						}
					Else
						{
							Gui, MainApp:Add, CheckBox, vClearOldFiles , Clear Old Files
						}
								
					IniRead, vChkState, %IniFile%, Settings, OpenLocation , 0
					if (vChkState = 1)
						{
							Gui, MainApp:Add, CheckBox, vOpenAppXNFolder Checked, Open AppXN Folder
						}
					Else
						{
							Gui, MainApp:Add, CheckBox, vOpenAppXNFolder , Open AppXN Folder
						}
				
					IniRead, vChkState, %IniFile%, Settings, LaunchAppXN , 1
					if (vChkState = 1)
						{
							Gui, MainApp:Add, CheckBox, vLaunchAppXN Checked, Launch AppXN
						}
					Else
						{
							Gui, MainApp:Add, CheckBox, vLaunchAppXN , Launch AppXN
						}
					
					IniRead, vChkState, %IniFile%, Settings, DesktopIcon , 0
					if (vChkState = 1)
						{
							Gui, MainApp:Add, CheckBox, vDesktopIcon Checked, Create Desktop Icon
						}
					Else
						{
							Gui, MainApp:Add, CheckBox, vDesktopIcon , Create Desktop Icon
						}
					IniRead, vChkState, %IniFile%, Settings, Install2User , 1
					if (vChkState = 1)
						{
							Gui, MainApp:Add, CheckBox, vInstall2User Checked, Install to AppData`n(Bypass Security)			
						}
					Else
						{
							Gui, MainApp:Add, CheckBox, vInstall2User , Install to AppData`n(Bypass Security)
						}
						
						BREAK
					
				}
			Else
				{
					IniWrite, 1, %IniFile%, Settings, GetZip
					IniWrite, 0, %IniFile%, Settings, ClearCache
					IniWrite, 1, %IniFile%, Settings, ClearOldFiles
					IniWrite, 0, %IniFile%, Settings, OpenLocation
					IniWrite, 1, %IniFile%, Settings, LaunchAppXN
					IniWrite, 0, %IniFile%, Settings, DesktopIcon
					IniWrite, 1, %IniFile%, Settings, Install2User
					IniWrite, % A_ScreenWidth - 325 , %IniFile%, Settings, XPosition
					IniWrite, 25, %IniFile%, Settings, YPosition
					IniWrite, B0B0FF, %IniFile%, Customize, ProgressBarColor
					
					IniWrite, 6 , %IniFile%,Environments,EnvironmentCount
					
						IniWrite, % "Client1" , %IniFile% , EnvironmentNo1,Name
						IniWrite, % "Client1" , %IniFile% , EnvironmentNo1,DefaultName
						IniWrite, % "\\resourceServerapp1\Client1_UI" , %IniFile% , EnvironmentNo1,AppPath
						IniWrite, % "\\resourceServerapp1\Client1_DataModel" , %IniFile% , EnvironmentNo1,DataModelPath
						IniWrite, % "Client1" , %IniFile% , EnvironmentNo1,IconName
						IniWrite, % "https://resourceServerapp1:44388/" , %IniFile% , EnvironmentNo1,TargetService
						IniWrite, % "" , %IniFile% , EnvironmentNo1,Client
						IniWrite, % "44905" , %IniFile% , EnvironmentNo1,Port


						IniWrite, % "AppXN_Client2" , %IniFile% , EnvironmentNo2,Name
						IniWrite, % "AppXN_Client2" , %IniFile% , EnvironmentNo2,DefaultName
						IniWrite, % "\\resourceServerapp1\AppXN_Client2_UI" , %IniFile% , EnvironmentNo2,AppPath
						IniWrite, % "\\resourceServerapp1\AppXN_Client2_DataModel" , %IniFile% , EnvironmentNo2,DataModelPath
						IniWrite, % "AppXN_Client2" , %IniFile% , EnvironmentNo2,IconName
						IniWrite, % "https://resourceServerapp1:44388/" , %IniFile% , EnvironmentNo2,TargetService
						IniWrite, % "" , %IniFile% , EnvironmentNo2,Client
						IniWrite, % "44906" , %IniFile% , EnvironmentNo2,Port




					Sleep, 2000

				}		
			}

		Gui, MainApp:Add, Text, % " x" TabMasterX + 10 " y" TabMasterY + 225  " w300 vMessageBox Center "



		if (	FileExist(IniFile)	)
			{ 
				IniRead, vProgressBarColor, %IniFile%, Customize, ProgressBarColor , B0B0FF
				if (vProgressBarColor = "B0B0FF") {
						IniWrite, B0B0FF, %IniFile%, Customize, ProgressBarColor
					}
				Gui, MainApp:Add, Progress,% " x" TabMasterX + 10 " y" TabMasterY + 250 " w300 h20 c" vProgressBarColor " -Smooth vMyProgress ", 0

			}
		else
			{	
				Gui, MainApp:Add, Progress, % " x" TabMasterX + 10 " y" TabMasterY + 250 " w300 h20  cB0B0FF -Smooth vMyProgress ", 0
			}
			


		Gui, MainApp:Add, Button, % " x" TabMasterX + 10 " y" TabMasterY + 275 " w300 gAppXNUpdate ", AppXN Update

		Gui, MainApp:Add, Button, % " x" TabMasterX + 10 " y" TabMasterY + 300 " w300 gCloseApp ", Save and Close  

/*
	Manage Environments
*/

		Gui,MainApp:Tab,Manage Environments
			Gui,MainApp:Add,Text,% " x" TabMasterX + 10 " y" TabMasterY + 30 ,Hello
							
			ListBoxItems := ""

				IniRead, tEnvironmentCount,% IniFile, Environments,EnvironmentCount,0
				i=1
				while (i <= tEnvironmentCount)
				{
					IniRead,tEnvironmentName,% IniFile, % "EnvironmentNo" i, IconName
					ListBoxItems := ListBoxItems tEnvironmentName "|"
					i += 1
				}
				
			Gui, MainApp:Add, ListBox, % " x" TabMasterX + 10 " y" TabMasterY + 30 " w125 h200 vEnvironmentListManage gEnvironmentListManage altsubmit ", % ListBoxItems

				GuiControl,Choose, EnvironmentListManage, 1

				tMngEnvCtrlWidth = 150
					
			Gui,MainApp:Add,Text, % " x" TabMasterX + 145 " y" TabMasterY + 30 , Environment Name
			Gui,MainApp:Add,Edit, % " x" TabMasterX + 145 " y" TabMasterY + 45 " w" tMngEnvCtrlWidth " vMngEnvName "
			
			Gui,MainApp:Add,Text, % " x" TabMasterX + 145 " y" TabMasterY + 70 , Default Name (Opt)
			Gui,MainApp:Add,Edit, % " x" TabMasterX + 145 " y" TabMasterY + 85 " w" tMngEnvCtrlWidth " vMngEnvDefaultName "
			
			Gui,MainApp:Add,Text, % " x" TabMasterX + 145 " y" TabMasterY + 110 , GUI Path
			Gui,MainApp:Add,Edit, % " x" TabMasterX + 145 " y" TabMasterY + 125 " w" tMngEnvCtrlWidth " vMngEnvGuiPath "
			
			Gui,MainApp:Add,Text, % " x" TabMasterX + 145 " y" TabMasterY + 150 , DataModel Path
			Gui,MainApp:Add,Edit, % " x" TabMasterX + 145 " y" TabMasterY + 165 " w" tMngEnvCtrlWidth " vMngEnvDMPath "

			Gui,MainApp:Add,Text, % " x" TabMasterX + 145 " y" TabMasterY + 190 , Icon Name
			Gui,MainApp:Add,Edit, % " x" TabMasterX + 145 " y" TabMasterY + 205 " w" tMngEnvCtrlWidth " vMngEnvIconName "

			Gui,MainApp:Add,Text, % " x" TabMasterX + 145 " y" TabMasterY + 230 , Service URL
			Gui,MainApp:Add,Edit, % " x" TabMasterX + 145 " y" TabMasterY + 245 " w" tMngEnvCtrlWidth " vMngEnvTargetSvc  ", 

			Gui,MainApp:Add,Text, % " x" TabMasterX + 145 " y" TabMasterY + 270 , Client (Opt)
			Gui,MainApp:Add,ComboBox, % " x" TabMasterX + 145 " y" TabMasterY + 285 " w" tMngEnvCtrlWidth " vMngEnvClient Limit ", % A_Space "|MCCC"

			Gui,MainApp:Add,Text, % " x" TabMasterX + 145 " y" TabMasterY + 310 , Port Number
			Gui,MainApp:Add,Edit, % " x" TabMasterX + 145 " y" TabMasterY + 325 " w" tMngEnvCtrlWidth " vMngEnvPort Number Limit5 "

			Gui,MainApp:Add,Button, % " x" TabMasterX + 10 " y" TabMasterY + 245 " w125 center gMngUpdateExisting " , Update
			Gui,MainApp:Add,Button, % " x" TabMasterX + 10 " y" TabMasterY + 285 " w125 center gMngAddNew " , Add New
			Gui,MainApp:Add,Button, % " x" TabMasterX + 10 " y" TabMasterY + 325 " w125 center gMngRemove " , Remove
			
			
/*
	Location/Placement
*/			
		Gui, MainApp:Margin , 5, 5



		if (	FileExist(IniFile)	)
			{ 
				IniRead, vXPosition, %IniFile%, Settings, XPosition , % "x" A_ScreenWidth - 325 
				IniRead, vYPosition, %IniFile%, Settings, YPosition , % " y" 25
				
				if (vXPosition >  A_ScreenWidth - 325 )
					{
						vXPosition = % "x" A_ScreenWidth - 325
					}		  
				if (vXPosition <  0 )
					{
						vXPosition = % "x" A_ScreenWidth - 325
					}
				  
				if (vYPosition >  A_ScreenHeight - 325 )
					{
						vYPosition = %  A_ScreenHeight - 325
					}		
				if (vYPosition < 0 )
					{
						vYPosition = %  A_ScreenHeight - 325
					}	
				
				
				
				Try
					{
						Gui, Show, % "x" vXPosition " y" vYPosition  x250, %GUITitleVersion%	;AppXN Updater
					}
				Catch
					{
						Gui, Show, % "x" A_ScreenWidth - 325 " y" 25 x250, %GUITitleVersion%	;AppXN Updater
					}
			}
		else
			{
				Gui, Show, % "x" A_ScreenWidth - 325 " y" 25 x250, %GUITitleVersion%	;AppXN Updater
			}

		




WinSet, Style, -0x30000, A ; WS_MAXIMIZEBOX 0x10000 + WS_MINIMIZEBOX 0x20000
RETURN ;  


/*
		Manage Environments Core
*/


			MngAddNew:
			{
				Global IniFile
				Gui,MainApp:Submit,NoHide
				
					StringLen, tLen, MngEnvName
					if (tLen < 3) {
							MsgBox, 48 , Warning, The Environment Name must be at least 3 Characters
							Exit
						}
					
					
					StringLen, tLen, MngEnvIconName
					If (tLen < 3) {
							MsgBox, 48 , Warning, The Icon Name must be at least 3 Characters
							Exit
						}
					
					
					IfNotInString, MngEnvTargetSvc, HTTP
					{
							MsgBox, 48 , Warning, % "The Environment Path must fit the following format `n`nHTTP://<ServerName>:<Port #>`n`tOR`nHTTP://<ServerName>:<Port #>"
							Exit
					}
				
				
				IniRead, tEnvCount ,% IniFile, Environments, EnvironmentCount
				tEnvCount += 1
				IniWrite, % tEnvCount  ,% IniFile, Environments, EnvironmentCount
				
					IniWrite, % MngEnvName			, %IniFile% , % "EnvironmentNo" tEnvCount , Name
					IniWrite, % MngEnvDefaultName 	, %IniFile% , % "EnvironmentNo" tEnvCount , DefaultName
					IniWrite, % MngEnvGuiPath		, %IniFile% , % "EnvironmentNo" tEnvCount , AppPath
					IniWrite, % MngEnvDMPath		, %IniFile% , % "EnvironmentNo" tEnvCount , DataModelPath
					IniWrite, % MngEnvIconName		, %IniFile% , % "EnvironmentNo" tEnvCount , IconName
					IniWrite, % MngEnvTargetSvc		, %IniFile% , % "EnvironmentNo" tEnvCount , TargetService
					IniWrite, % MngEnvClient		, %IniFile% , % "EnvironmentNo" tEnvCount , Client
					IniWrite, % MngEnvPort			, %IniFile% , % "EnvironmentNo" tEnvCount , Port

				fMngUpdateLists()
				
				


				
				
			}
			RETURN	;
			
			MngRemove:
			{
				Global IniFile
				Gui,MainApp:Submit,NoHide
				
				
				IniRead, tIconName	, %IniFile%	, % "EnvironmentNo" EnvironmentListManage , IconName
				
				msgBox,36,Remove Environment?,Are you sure you want to remove the AppXN Install from your list?`n`n%tIconName%
					IfMsgBox, Yes
						{
							
							IniRead, INIRecordCount,% IniFile, Environments,EnvironmentCount,0
							if (INIRecordCount = 1)
								{
									msgbox, 48, Ettiquette Violation,It is not polite to deprive me of the last enviornment.`nYou must have 2 or more environments to use the remove function ; `n
									EXIT
								}
							ELSE {
									
									i := EnvironmentListManage
									IniWrite,% INIRecordCount - 1  ,% IniFile, Environments,EnvironmentCount
									
									While (i <= INIRecordCount)
										
										{
											if (i = INIRecordCount)
												{
													IniDelete,% IniFile ,% "EnvironmentNo" i
												}
											ELSE
												{
													
													IniRead, tMngEnvName		, %IniFile% , % "EnvironmentNo" i + 1 , Name
													IniRead, tMngEnvDefaultName , %IniFile% , % "EnvironmentNo" i + 1 , DefaultName
													IniRead, tMngEnvGuiPath		, %IniFile% , % "EnvironmentNo" i + 1 , AppPath
													IniRead, tMngEnvDMPath		, %IniFile% , % "EnvironmentNo" i + 1 , DataModelPath
													IniRead, tMngEnvIconName	, %IniFile% , % "EnvironmentNo" i + 1 , IconName
													IniRead, tMngEnvTargetSvc	, %IniFile% , % "EnvironmentNo" i + 1 , TargetService
													IniRead, tMngEnvClient		, %IniFile% , % "EnvironmentNo" i + 1 , Client
													IniRead, tMngEnvPort		, %IniFile% , % "EnvironmentNo" i + 1 , Port
													
														IniWrite, % tMngEnvName			, %IniFile% , % "EnvironmentNo" i , Name
														IniWrite, % tMngEnvDefaultName 	, %IniFile% , % "EnvironmentNo" i , DefaultName
														IniWrite, % tMngEnvGuiPath		, %IniFile% , % "EnvironmentNo" i , AppPath
														IniWrite, % tMngEnvDMPath		, %IniFile% , % "EnvironmentNo" i , DataModelPath
														IniWrite, % tMngEnvIconName		, %IniFile% , % "EnvironmentNo" i , IconName
														IniWrite, % tMngEnvTargetSvc	, %IniFile% , % "EnvironmentNo" i , TargetService
														IniWrite, % tMngEnvClient		, %IniFile% , % "EnvironmentNo" i , Client
														IniWrite, % tMngEnvPort			, %IniFile% , % "EnvironmentNo" i , Port
													
													
													
												}
													i += 1
										}
										
										
								}
							
						}
					sleep, 2000
				
				
				
				
				
				
				fMngUpdateLists()
			}
			
			RETURN	;	

			MngUpdateExisting:
			{
					
				Global IniFile
				Gui,MainApp:Submit,NoHide
				
					msgBox, 20 , Please Confirm, Are you sure you want to overwrite these settings?
					IfMsgBox, No
					{
						EXIT
					}
				
				
					StringLen, tLen, MngEnvName
					if (tLen < 3) {
							MsgBox, 48 , Warning, The Environment Name must be at least 3 Characters
							Exit
						}
					
					
					StringLen, tLen, MngEnvIconName
					If (tLen < 3) {
							MsgBox, 48 , Warning, The Icon Name must be at least 3 Characters
							Exit
						}
					
					
					IfNotInString, MngEnvTargetSvc, HTTP
					{
							MsgBox, 48 , Warning, % "The Environment Path must fit the following format `n`nHTTP://<ServerName>:<Port #>`n`tOR`nHTTP://<ServerName>:<Port #>"
							Exit
					}
				
				
					IniWrite, % MngEnvName			, %IniFile% , % "EnvironmentNo" EnvironmentListManage , Name
					IniWrite, % MngEnvDefaultName 	, %IniFile% , % "EnvironmentNo" EnvironmentListManage , DefaultName
					IniWrite, % MngEnvGuiPath		, %IniFile% , % "EnvironmentNo" EnvironmentListManage , AppPath
					IniWrite, % MngEnvDMPath		, %IniFile% , % "EnvironmentNo" EnvironmentListManage , DataModelPath
					IniWrite, % MngEnvIconName		, %IniFile% , % "EnvironmentNo" EnvironmentListManage , IconName
					IniWrite, % MngEnvTargetSvc		, %IniFile% , % "EnvironmentNo" EnvironmentListManage , TargetService
					IniWrite, % MngEnvClient		, %IniFile% , % "EnvironmentNo" EnvironmentListManage , Client
					IniWrite, % MngEnvPort			, %IniFile% , % "EnvironmentNo" EnvironmentListManage , Port

				fMngUpdateLists()
				
			
			}
			RETURN	;


			fMngUpdateLists()
			{
				Global IniFile
					
					tListBoxItems := ""
					
					IniRead, tEnvironmentCount,% IniFile, Environments,EnvironmentCount,0
						i=1
						while (i <= tEnvironmentCount)
						{
							IniRead,tEnvironmentName,% IniFile, % "EnvironmentNo" i, IconName
							tListBoxItems := tListBoxItems tEnvironmentName "|"
							i += 1
						}


						GuiControl,MainApp:, EnvironmentListManage, % "|" tListBoxItems
						GuiControl,Choose, EnvironmentListManage, 1
						

						GuiControl,MainApp:, EnvironmentList, % "|All|" tListBoxItems
						GuiControl,Choose, EnvironmentList, 2
						
					
			}
			
			RETURN	;

			EnvironmentListManage:
			{
				Global IniFile
				Gui,MainApp:Submit,NoHide
				
					IniRead, tMngEnvName		, %IniFile% , % "EnvironmentNo" EnvironmentListManage , Name
					IniRead, tMngEnvDefaultName , %IniFile% , % "EnvironmentNo" EnvironmentListManage , DefaultName
					IniRead, tMngEnvGuiPath		, %IniFile% , % "EnvironmentNo" EnvironmentListManage , AppPath
					IniRead, tMngEnvDMPath		, %IniFile% , % "EnvironmentNo" EnvironmentListManage , DataModelPath
					IniRead, tMngEnvIconName	, %IniFile% , % "EnvironmentNo" EnvironmentListManage , IconName
					IniRead, tMngEnvTargetSvc	, %IniFile% , % "EnvironmentNo" EnvironmentListManage , TargetService
					IniRead, tMngEnvClient		, %IniFile% , % "EnvironmentNo" EnvironmentListManage , Client
					IniRead, tMngEnvPort		, %IniFile% , % "EnvironmentNo" EnvironmentListManage , Port
					
				GuiControl, MainApp:, MngEnvName		, % tMngEnvName
				GuiControl, MainApp:, MngEnvDefaultName	, % tMngEnvDefaultName
				GuiControl, MainApp:, MngEnvGuiPath		, % tMngEnvGuiPath
				GuiControl, MainApp:, MngEnvDMPath		, % tMngEnvDMPath
				GuiControl, MainApp:, MngEnvIconName	, % tMngEnvIconName
				GuiControl, MainApp:, MngEnvTargetSvc	, % tMngEnvTargetSvc ""
				
				StringLen, tLen,tMngEnvClient 
				If (tLen = 0)
				{
					GuiControl, Choose	, MngEnvClient		,  1
					
				}
				ELSE{
					GuiControl, Choose	, MngEnvClient		, % tMngEnvClient
				}
				GuiControl, MainApp:, MngEnvPort		, % tMngEnvPort
											
				
				
			}
			RETURN	;
			




/*
	AppXN Update Core
*/


			AppXNUpdate:
			{
				Global Time
				Global uAppXNTargetPath := ""
				FormatTime, Time,, yyyy MM dd 
				Global uTime :=  Time
				Global IniFile
				IniRead, tEnvironmentCount,% IniFile, Environments,EnvironmentCount,0
				GUI, MainApp:Submit, NoHide		
				
				
			;	Set Up AppXN Parameters
				

					if (EnvironmentList = 1)
					{

							i = 2
						Loop {


									fAppXNSetEnvironmentInfo(i)
										Gui,MainApp:Show,,% GUITitleVersion " - Updating: " vEnvironmentName
									fAppXNSetGlobalVariables(Install2User)
									fAppXNClearCache(ClearCache,Install2User)
									fAppXNPullUpdate(Install2User,GetZip,vEnvironmentName,vresourceServerAppPath,vresourceServerDMPath)
									fAppXNConfig(vEnvironmentName,vEnvironmentNameDefaultOption,vTargetService)
									fAppXNDesktopIcon(DesktopIcon,vEnvironmentName,vIconName)
									fAppXNStartMenuIcon(vEnvironmentName,vIconName)	
									fAppXNStartMenuParentFolderLink(uAppXNTargetPath)
									fAppXNCleanUpOldFiles(vEnvironmentName,ClearOldFiles)
									fAppXNOpenFolder(OpenAppXNFolder,uAppXNTargetPath,vEnvironmentName,uTime)
									i += 1
									
								if (i > tEnvironmentCount + 1)
								{
									BREAK
								}
									
							}
					}
					
					ELSE {
						fAppXNSetEnvironmentInfo(EnvironmentList)
							Gui,MainApp:Show,,% GUITitleVersion " - Updating: " vEnvironmentName
						fAppXNSetGlobalVariables(Install2User)
						fAppXNClearCache(ClearCache,Install2User)
						fAppXNPullUpdate(Install2User,GetZip,vEnvironmentName,vresourceServerAppPath,vresourceServerDMPath)
						fAppXNConfig(vEnvironmentName,vEnvironmentNameDefaultOption,vTargetService)
						fAppXNDesktopIcon(DesktopIcon,vEnvironmentName,vIconName)
						fAppXNStartMenuIcon(vEnvironmentName,vIconName)	
						fAppXNStartMenuParentFolderLink(uAppXNTargetPath)
						fAppXNCleanUpOldFiles(vEnvironmentName,ClearOldFiles)
						fAppXNOpenFolder(OpenAppXNFolder,uAppXNTargetPath,vEnvironmentName,uTime)
						fAppXNLaunchApp(LaunchAppXN,uAppXNTargetPath,vEnvironmentName, uTime)
						
					}


				fAppXNLaunchApp(LaunchAppXN,uAppXNTargetPath,vEnvironmentName, uTime)


				Gui,MainApp:Show,,% GUITitleVersion
				GuiControl, MainApp:, MyProgress, 100
				GuiControl, MainApp: ,MessageBox, Update Complete
				
				
			}	
				RETURN	;


			fAppXNSetEnvironmentInfo(EnvironmentList)
			{
				Global IniFile
				GuiControl, MainApp:, MyProgress, 0
				GuiControl, MainApp: ,MessageBox, Setting Environment
				
						TargetEnv = % EnvironmentList - 1
						IniRead, tINIEnvName	, %  IniFile  , % "EnvironmentNo" TargetEnv  ,	Name
						IniRead, tINIEnvDefName	, %  IniFile  , % "EnvironmentNo" TargetEnv  ,	DefaultName
						IniRead, tINIEnvAppPath	, %  IniFile  , % "EnvironmentNo" TargetEnv  ,	AppPath
						IniRead, tINIEnvDMPath	, %  IniFile  , % "EnvironmentNo" TargetEnv  ,	DataModelPath
						IniRead, tINIEnvIconName, %  IniFile  , % "EnvironmentNo" TargetEnv  ,	IconName
						IniRead, tINIEnvTarget	, %  IniFile  , % "EnvironmentNo" TargetEnv  ,	TargetService
						IniRead, tINIEnvClient	, %  IniFile  , % "EnvironmentNo" TargetEnv  ,	Client
						IniRead, tINIEnvPort	, %  IniFile  , % "EnvironmentNo" TargetEnv  ,	Port
							
						
							Global vEnvironmentName := % tINIEnvName
							Global vEnvironmentNameDefaultOption := % tINIEnvDefName
							Global vresourceServerAppPath := % tINIEnvAppPath
							Global vresourceServerDMPath := % tINIEnvDMPath
							Global vIconName := % tINIEnvIconName
							Global vTargetService := % tINIEnvTarget
							vClient	:=	% tINIEnvClient ""
							;	44905
							

						
					
				GuiControl, MainApp: ,MessageBox, Updating %vEnvironmentName%
				GuiControl, MainApp:, MyProgress, 0

			}
			RETURN	;
				


			fAppXNClearCache(ClearCache,Install2User)
				{
			; Require AppXN to be Closed
					Global uAppXNTargetPath
				
				GuiControl, MainApp:, MyProgress, +1
				GuiControl, MainApp: ,MessageBox, Clearing Cache
					
					if (ClearCache)
						{
							Loop
								{	
									Process, Exist, CoreAppGUI.exe
										{
											If ! errorLevel
												{
													Break
												}
											Else {
													msgbox , 3,Application Open ,It looks like an instance of AppXN is currently running.`nIf you do not close AppXN, you may have issues running it.`n`nYou Can:`n   - Click Yes for Me to Close AppXN.`n   - Click No to accept the risk.`n   - Click Cancel to stop this update.		 
													ifMsgBox Yes
														{
															Loop
																{
																	Process, Exist, CoreAppGUI.exe
																	{
																		If ! errorLevel
																			{
																				Break
																			}
																
																
																		Process, Close, CoreAppGUI.exe
																		sleep, 5000
																
																	}
																}
															
														}
													IfMsgBox No 
														{
															Break
														}
													IfMsgBox Cancel
														{
															EXIT
														}
													
													
												}
										
										}	
								}
							}
					; Clear AppXN Cache
						if (ClearCache)
							{
								GuiControl, MainApp: ,MessageBox, Clearing Cache			
								EnvGet,tLocalAppData, LOCALAPPDATA
								try {
										FileRemoveDir, %tLocalAppData%\Conduent, 1
									}
								try {
										FileRemoveDir, %tLocalAppData%\CPNY, 1
									}
								try {
										FileRemoveDir, %tLocalAppData%\CPNY_Dir1, 1
									}
								try {
										FileRemoveDir, %tLocalAppData%\CPNY_Dir2, 1
									}
								
							}
						
				GuiControl, MainApp:, MyProgress, +1
				GuiControl, MainApp: ,MessageBox, Cache	Cleared
						
										
								
								
				}
			RETURN	;



			fAppXNSetGlobalVariables(Install2User)
			{

				GuiControl, MainApp:, MyProgress, +1
				GuiControl, MainApp: ,MessageBox, Setting Global Variables

				Global uAppXNTargetPath 	
				
				IF (Install2User) {
						EnvGet, vAppXNTargetPath, LOCALAPPDATA	
						Global uAppXNTargetPath := % vAppXNTargetPath "\_AppXN"
					}
				ELSE {
						Global uAppXNTargetPath := "C:\AppXN"
					}
					
				GuiControl, MainApp:, MyProgress, +1
				GuiControl, MainApp: ,MessageBox, Global Variables Set
			}
			RETURN	;


			fAppXNCleanUpOldFiles(vEnvironmentName,ClearOldFiles)
			{
				
				GuiControl, MainApp:, MyProgress, +1
				GuiControl, MainApp: ,MessageBox, Cleaning Old Files

					Global uAppXNTargetPath
					vAppXNTargetPath = % uAppXNTargetPath
			; Clean Up Old Files
					If (ClearOldFiles) 
						{
							GuiControl, MainApp: ,MessageBox, Removing Old Files
							Array := []
							loop Files, %vAppXNTargetPath%\%vEnvironmentName%\*, D
								{
									tFileName = %FileList%%A_LoopFileName%
									Array.Push(tFileName)
								}
								
								ArrayTarget = % Array.MaxIndex() -2

							i = 1
							If (ArrayTarget > 2)
								{
									while ( i <= ArrayTarget)
										{
											TargetDir = % Array[i]
											TargetDir = %vAppXNTargetPath%\%vEnvironmentName%\%TargetDir%
											MouseGetPos,MouseX,MouseY,1
											ToolTip,Removing %TargetDir%,MouseX,MouseY
											FileRemoveDir,%TargetDir% , 1 
											i += 1
											ToolTip,
										}
								}
						}
				GuiControl, MainApp:, MyProgress, +3
				GuiControl, MainApp: ,MessageBox, Old Files Cleared
							
				
			}
			RETURN	;	

			fAppXNOpenFolder(OpenAppXNFolder,vAppXNTargetPath,vEnvironmentName,Time)
			{
				

				GuiControl, MainApp:, MyProgress, +1
				GuiControl, MainApp: ,MessageBox, Fetching AppXN Folder
				
			; Open AppXN Folder
					If (OpenAppXNFolder) 
						{
							targetFolder = %vAppXNTargetPath%\%vEnvironmentName%\%Time%\
							clipboard = %targetFolder%
							Try {
									Run, explore %targetFolder%
								}
						}
					
			;	Clean Up removed files moved to Recycle Bin	
				FileRecycleEmpty
				
				GuiControl, MainApp:, MyProgress, +1
				GuiControl, MainApp: ,MessageBox, AppXN Folder Opened
				
			}
			RETURN		;

			fAppXNLaunchApp(LaunchAppXN,vAppXNTargetPath,vEnvironmentName,Time)
			{
				
			; Launch AppXN		
					if (LaunchAppXN) {
						
							Try	{
								GuiControl, MainApp: ,MessageBox, Launching AppXN

								Run, %vAppXNTargetPath%\%vEnvironmentName%\%Time%\GUI_Console\CoreAppGUI.exe
								}
						}

				GuiControl, MainApp:, MyProgress, 100
				GuiControl, MainApp: ,MessageBox, Launching AppXN App			

			}
			RETURN	;


			fAppXNStartMenuParentFolderLink(vAppXNTargetPath)
			{
				
			; Start Menu Insert Link to Folder
					
					EnvGet, vAppDataRoamingPath, AppData
					vStartMenuPath = %vAppDataRoamingPath%\Microsoft\Windows\Start Menu\Programs\CPNY AppXN
					
					
					If !InStr(FileExist(vStartMenuPath),"D")	{
						FileCreateDir, %vStartMenuPath%
						}	
					
					vStartMenuPath = %vStartMenuPath%\_AppXN Install Folder.lnk
					If FileExist(vStartMenuPath)	{
							GuiControl, MainApp: ,MessageBox, Removing Previous Start Menu Shortcut For Folder
							FileDelete, %vStartMenuPath%
						}	
					

							GuiControl, MainApp: ,MessageBox, Creating New Start Menu Shortcut For Folder
							FileCreateShortcut, %vAppXNTargetPath%,%vStartMenuPath%,%vAppXNTargetPath%,,AppXN Folder,%A_WinDir%\System32\shell32.dll,, 68

					GuiControl, MainApp:, MyProgress, +1
					GuiControl, MainApp: ,MessageBox, Creating Parent Folder in Start Menu
						
						;	%AppData%\Microsoft\Windows\Start Menu\Programs\
						
				
			}
			RETURN	;


			fAppXNStartMenuIcon(vEnvironmentName,vIconName)	
			{
				
			; Start Menu Insert GUI

					GuiControl, MainApp:, MyProgress, +1
					GuiControl, MainApp: ,MessageBox, Creating Start Menu Icons
					Global uAppXNTargetPath
					vAppXNTargetPath = % uAppXNTargetPath
					Global uTime
					Time = % uTime
					
					EnvGet, vAppDataRoamingPath, AppData
					vStartMenuPath = %vAppDataRoamingPath%\Microsoft\Windows\Start Menu\Programs\CPNY AppXN
					
					
					If !InStr(FileExist(vStartMenuPath),"D")	{
						FileCreateDir, %vStartMenuPath%
						}	
					
					vStartMenuPath = %vStartMenuPath%\%vIconName% GUI.lnk
					If FileExist(vStartMenuPath)	{
							GuiControl, MainApp: ,MessageBox, Removing Previous Start Menu Shortcut
							FileDelete, %vStartMenuPath%
						}	
					

							GuiControl, MainApp: ,MessageBox, Creating New Start Menu Shortcut
							FileCreateShortcut, %vAppXNTargetPath%\%vEnvironmentName%\%Time%\GUI_Console\CoreAppGUI.exe,%vStartMenuPath%,%vAppXNTargetPath%\%vEnvironmentName%\%Time%\GUI_Console,,AppXN link for %vIconName%,%A_WinDir%\System32\shell32.dll,, 93

					GuiControl, MainApp:, MyProgress, +2
					GuiControl, MainApp: ,MessageBox, Start Menu Icons Created
						
						
						;	%AppData%\Microsoft\Windows\Start Menu\Programs\
						

			}
			RETURN	;


			fAppXNPullUpdate(Install2User,GetZip,vEnvironmentName,vresourceServerAppPath,vresourceServerDMPath)
			{
				Global uTime
				Global uAppXNTargetPath 	
						vAppXNTargetPath = %uAppXNTargetPath% 
						Time = % uTime

				
					GuiControl, MainApp: ,MessageBox, Updating %vEnvironmentName%
					GuiControl, MainApp:, MyProgress, +1
					


			GuiControl, MainApp: ,MessageBox, Starting File Dig
			GuiControl, MainApp:, MyProgress, +1

				Criteria1Path = %vAppXNTargetPath%\%vEnvironmentName%\%Time% 
				If (	InStr(FileExist(Criteria1Path),"D")	)
					{
						GuiControl, MainApp: ,MessageBox, User Input Needed
						msgBox, 4, Existing Update, You have already updated %vEnvironmentName% today. `nDo you wish to continue?`n
						ifMsgBox, No
							{
								GuiControl, MainApp: ,MessageBox, Update Canceled by User
								EXIT 
							}
						ifMsgBox, Yes
							{
								
								GuiControl, MainApp: ,MessageBox, Removing Existing Directory
								Try {
										FileRemoveDir, %vAppXNTargetPath%\%vEnvironmentName%\%Time%, 1
									}
								Catch {
										
										GuiControl, MainApp: ,MessageBox, Unable to Update AppXN for one of the following reasons `n - AppXN is open`n - AppXN Folder is Read Only `n - User does not have access to AppXN folder `n - AppXN Folder is hidden or does not exist
										EXIT
									}
								
							}
						
					}



			;	Create Path for Installation

			GuiControl, MainApp: ,MessageBox, Starting File Dig
			GuiControl, MainApp:, MyProgress, +2
					

				
					
					If !InStr(FileExist(vAppXNTargetPath),"D")	{
						FileCreateDir, %vAppXNTargetPath%
						}	
						
					vAppXNTestPath = %vAppXNTargetPath%\%vEnvironmentName%
					If !InStr(FileExist(vAppXNTestPath),"D")	{
						FileCreateDir, %vAppXNTestPath%
						}	
						
					vAppXNTestPath = %vAppXNTargetPath%\%vEnvironmentName%\%Time%
					If !InStr(FileExist(vAppXNTestPath),"D")	{
						FileCreateDir, %vAppXNTestPath%
						}	


					GuiControl, MainApp:, MyProgress, +2
						

			; Pull DataModel from Server		
					if FileExist(vresourceServerDMPath "\DataModel.zip") and (GetZip)
						{
							GuiControl, MainApp: ,MessageBox, Fetching DataModel Zip				
							FileCreateDir, %vAppXNTargetPath%\%vEnvironmentName%\%Time%\DataModel
							FileCopy, %vresourceServerDMPath%\DataModel.zip, %vAppXNTargetPath%\%vEnvironmentName%\%Time%,1
							GuiControl, MainApp:, MyProgress, +12.5	

							GuiControl, MainApp: ,MessageBox, Unzipping Data Model				
							RunWait, Powershell.exe "expand-archive -LiteralPath '%vAppXNTargetPath%\%vEnvironmentName%\%Time%\DataModel.zip' -Destination '%vAppXNTargetPath%\%vEnvironmentName%\%Time%\DataModel\'", , Hide
							GuiControl, MainApp:, MyProgress, +12.5
							
						}
					else {				
							GuiControl, MainApp: ,MessageBox, Downloading DataModel - Zip Not Available
							FileCopyDir, %vresourceServerDMPath%, %vAppXNTargetPath%\%vEnvironmentName%\%Time%,1
							GuiControl, MainApp:, MyProgress, +25			
						}
						

						Criteria1Path = %vAppXNTargetPath%\%vEnvironmentName%\%Time%\DataModels
						Criteria2Path = %vAppXNTargetPath%\%vEnvironmentName%\%Time%\DataModels\TargetURL.xml
						Criteria3Path = %vAppXNTargetPath%\%vEnvironmentName%\%Time%\DataModels\DataModel
						Criteria4Path = %vAppXNTargetPath%\%vEnvironmentName%\%Time%\DataModels\DataModel\TargetURL.xml
					if 	(InStr(FileExist(Criteria1Path),"D")  && FileExist(Criteria2Path) )
						{
							FileMoveDir, %vAppXNTargetPath%\%vEnvironmentName%\%Time%\DataModels, %vAppXNTargetPath%\%vEnvironmentName%\%Time%\DataModel, R 
						}
					Else If (InStr(FileExist(Criteria3Path),"D") && FileExist(Criteria4Path) )
						{
							
							FileMoveDir, %vAppXNTargetPath%\%vEnvironmentName%\%Time%\DataModels\DataModel, %vAppXNTargetPath%\%vEnvironmentName%\%Time%\DataModel, 1
							FileRemoveDir,  %vAppXNTargetPath%\%vEnvironmentName%\%Time%\DataModels , 0
						}




			; Pull UI Server
					
					FileCreateDir, %vAppXNTargetPath%\%vEnvironmentName%\%Time%\GUI_Console

					if FileExist(vresourceServerAppPath "\GUI_Console.zip") and (GetZip)
						{	
							GuiControl, MainApp: ,MessageBox, Fetching GUI_Console.zip			
							FileCopy, %vresourceServerAppPath%\GUI_Console.zip, %vAppXNTargetPath%\%vEnvironmentName%\%Time%,1
							GuiControl, MainApp:, MyProgress, +12.5			
						
							GuiControl, MainApp: ,MessageBox, Unzipping AppXN App				
							RunWait, Powershell.exe "expand-archive -LiteralPath '%vAppXNTargetPath%\%vEnvironmentName%\%Time%\GUI_Console.zip' -Destination '%vAppXNTargetPath%\%vEnvironmentName%\%Time%\GUI_Console\'", , Hide
							GuiControl, MainApp:, MyProgress, +12.5			
							
						}

					Else
						{
							
							GuiControl, MainApp: ,MessageBox, Getting Client UI - Zip Not Available
							FileCopyDir, %vresourceServerAppPath%\GUI_Console, %vAppXNTargetPath%\%vEnvironmentName%\%Time%\GUI_Console,1
							GuiControl, MainApp:, MyProgress, +15			
						}

						Criteria3Path = %vAppXNTargetPath%\%vEnvironmentName%\%Time%\GUI_Console\GUI_Console
						Criteria4Path = %vAppXNTargetPath%\%vEnvironmentName%\%Time%\GUI_Console\GUI_Console\CoreAppGUI.exe
					If (InStr(FileExist(Criteria3Path),"D") && FileExist(Criteria4Path) )
						{
							GuiControl, MainApp: ,MessageBox, Resolving Zip Nesting
							FileMoveDir, %vAppXNTargetPath%\%vEnvironmentName%\%Time%\GUI_Console\GUI_Console, %vAppXNTargetPath%\%vEnvironmentName%\%Time%\GUI_Console, 2
							FileRemoveDir,  %vAppXNTargetPath%\%vEnvironmentName%\%Time%\GUI_Console\GUI_Console , 0
						}

						
						
						
						

			; Pull MC Server
								
					FileCreateDir, %vAppXNTargetPath%\%vEnvironmentName%\%Time%\Admin_UI
						
					if FileExist(vresourceServerAppPath "\Admin_UI.zip") and (GetZip)
						{
							GuiControl, MainApp: ,MessageBox, Fetching Admin_UI.zip	
							FileCopy, %vresourceServerAppPath%\Admin_UI.zip, %vAppXNTargetPath%\%vEnvironmentName%\%Time%,1
							GuiControl, MainApp:, MyProgress, +12.5			
							
							GuiControl, MainApp: ,MessageBox, Unzipping Management Console			
							RunWait, Powershell.exe "expand-archive -LiteralPath '%vAppXNTargetPath%\%vEnvironmentName%\%Time%\Admin_UI.zip' -Destination '%vAppXNTargetPath%\%vEnvironmentName%\%Time%\Admin_UI\'", , Hide
							GuiControl, MainApp:, MyProgress, +12.5			
						
							
						}

					Else
						{
							GuiControl, MainApp: ,MessageBox, Getting MC - Zip Not Available			
							FileCopyDir, %vresourceServerAppPath%\Admin_UI, %vAppXNTargetPath%\%vEnvironmentName%\%Time%\Admin_UI,1
							GuiControl, MainApp:, MyProgress, +15			
						}	

						
						Criteria3Path = %vAppXNTargetPath%\%vEnvironmentName%\%Time%\Admin_UI\Admin_UI
						Criteria4Path = %vAppXNTargetPath%\%vEnvironmentName%\%Time%\Admin_UI\Admin_UI\LRAssetManagementConsole.exe
					If (InStr(FileExist(Criteria3Path),"D") && FileExist(Criteria4Path) )
						{
							GuiControl, MainApp: ,MessageBox, Resolving Zip Nesting
							FileMoveDir, %vAppXNTargetPath%\%vEnvironmentName%\%Time%\Admin_UI\Admin_UI, %vAppXNTargetPath%\%vEnvironmentName%\%Time%\Admin_UI, 2
							FileRemoveDir,  %vAppXNTargetPath%\%vEnvironmentName%\%Time%\Admin_UI\Admin_UI , 0
						}			
						

				GuiControl, MainApp:, MyProgress, +1
				GuiControl, MainApp: ,MessageBox, Application Unpacked
				
			}
			RETURN	;

			fAppXNConfig(vEnvironmentName,vEnvironmentNameDefaultOption,vTargetService)
			{
				
				Global uAppXNTargetPath
				vAppXNTargetPath = % uAppXNTargetPath
						
			; AppXN Config


					GuiControl, MainApp: ,MessageBox, Setting Config
					GuiControl, MainApp:, MyProgress, +1
					
					TempFile = %A_WorkingDir%\Temp.Config

					FileOpen(TempFile, "rw")

					SourceFile = %vAppXNTargetPath%\%vEnvironmentName%\%Time%\GUI_Console\CoreAppGUI.exe.config
					
					skip = 0
					Loop
						{						
							CurrentLine = %A_Index%
							NextLine = % A_Index + 1
							FileReadLine,vDerp,%SourceFile%,%CurrentLine%
							FileReadLine,vDerp2,%SourceFile%,%NextLine%

					;		msgbox %vDerp%`n%vDerp2%`n
							If (InStr(vDerp2,"</configuration>") > 0)
								{
									FileAppend, %vDerp% `n, %TempFile%
									FileAppend, %vDerp2% `n, %TempFile%	
										break				
								}
							
							If (InStr(vDerp,"RootPath") > 0 and InStr(vDerp2,"<Value") > 0 )
								{
									FileAppend, %vDerp% `n, %TempFile%	; `n
									FileAppend, <value>%vAppXNTargetPath%\%vEnvironmentName%\%Time%\</value> `n, %TempFile%		;	`n
									skip = 2
								}
								
							If (InStr(vDerp,"TargetURL") > 0 and InStr(vDerp2,"<Value") > 0 )
								{
									FileAppend, %vDerp% `n, %TempFile%	; `n
									FileAppend, <value>%vTargetService%</value> `n, %TempFile%		;	`n
									skip = 2
								}
							
								
							If (InStr(vDerp,"ClientPath") > 0 and InStr(vDerp2,"<Value") > 0  and vClient = "MCCC" )
								{
									FileAppend, %vDerp% `n, %TempFile%	; `n
									FileAppend, <value>%vClient%</value> `n, %TempFile%		;	`n
									skip = 2
								}
								
							If (InStr(vDerp,"FirstUse") > 0 and InStr(vDerp2,"<Value") > 0 ) 
								{
									FileAppend, %vDerp% `n, %TempFile%	; `n
									FileAppend, <value>False</value> `n, %TempFile%		;	`n
									skip = 2
								}			
								
							If (InStr(vDerp,"Environment") > 0 and InStr(vDerp,"value=") > 0 ) 
								{
									FileAppend, <add key="Environment" value="%vEnvironmentNameDefaultOption%" />  `n, %TempFile%	; `n
									skip = 1
								}			




			 ;   <add key="Environment" value="Client1" /> 
								
							
							if (skip = 0)
								{
									FileAppend, %vDerp% `n, %TempFile%	;	`n
								}
							else
								{
									skip -= 1
								}
							
						}
						FileCopy, %TempFile% , %SourceFile%,1
						FileDelete %TempFile%
					
					GuiControl, MainApp:, MyProgress, +2
					GuiControl, MainApp: ,MessageBox, App Configured			



				
				
			}
			RETURN		;


			fAppXNDesktopIcon(DesktopIcon,vEnvironmentName,vIconName)
			{
					Global uAppXNTargetPath
					vAppXNTargetPath = % uAppXNTargetPath
					Global uTime
					Time = % uTime
						
						
			; Desktop Icon Placement
					try {
							GuiControl, MainApp: ,MessageBox, Removing Previous Shortcut
							FileDelete, C:\Users\%A_UserName%\Desktop\%vIconName% GUI.lnk
						}

					if (DesktopIcon) {
										GuiControl, MainApp: ,MessageBox, Creating New Shortcut
										
										FileCreateShortcut, %vAppXNTargetPath%\%vEnvironmentName%\%Time%\GUI_Console\CoreAppGUI.exe,C:\Users\%A_UserName%\Desktop\%vIconName% GUI.lnk,%vAppXNTargetPath%\%vEnvironmentName%\%Time%\GUI_Console,,AppXN link for %vIconName%,%A_WinDir%\System32\shell32.dll,, 93
										}

					GuiControl, MainApp:, MyProgress, +2
					GuiControl, MainApp: ,MessageBox, Desktop Icon Created
					
			}
			RETURN	;




/*
		System/General
*/


			AboutHandler:
			{
				msgbox, 0x40040, About AppXN Updater, Author:`tCarlos A Garcia II`n`tuser@mail.com `nVersion:`t2.6.0.12 `nPublished:`t2019-03-04 `t
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
			;	Update INI Settings
				
				GUI, MainApp:Submit, NoHide	
				WinGetPos, X, Y
						IniWrite, %GetZip%, %IniFile%, Settings, GetZip
						IniWrite, %ClearCache%, %IniFile%, Settings, ClearCache
						IniWrite, %ClearOldFiles%, %IniFile%, Settings, ClearOldFiles				
						IniWrite, %OpenAppXNFolder%, %IniFile%, Settings, OpenLocation
						IniWrite, %LaunchAppXN%, %IniFile%, Settings, LaunchAppXN
						IniWrite, %DesktopIcon%, %IniFile%, Settings, DesktopIcon
						IniWrite, %Install2User%, %IniFile%, Settings, Install2User
						IniWrite, %X%, %IniFile%, Settings, XPosition
						IniWrite, %Y%, %IniFile%, Settings, YPosition
					ExitApp
			}
			RETURN ;

			MainAppGuiClose:
			{
			;	Update INI Settings
				
				GUI, MainApp:Submit, NoHide	
				WinGetPos, X, Y
						IniWrite, %GetZip%, %IniFile%, Settings, GetZip
						IniWrite, %ClearCache%, %IniFile%, Settings, ClearCache
						IniWrite, %ClearOldFiles%, %IniFile%, Settings, ClearOldFiles			
						IniWrite, %OpenAppXNFolder%, %IniFile%, Settings, OpenLocation			
						IniWrite, %LaunchAppXN%, %IniFile%, Settings, LaunchAppXN
						IniWrite, %DesktopIcon%, %IniFile%, Settings, DesktopIcon
						IniWrite, %Install2User%, %IniFile%, Settings, Install2User
						IniWrite, %X%, %IniFile%, Settings, XPosition
						IniWrite, %Y%, %IniFile%, Settings, YPosition
					ExitApp
			}
			RETURN ;



