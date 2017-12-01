#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#SingleInstance,Force
#Include Lib\tf.ahk
FileEncoding, UTF-8
CoordMode, Pixel, Relative
CoordMode, Mouse, Relative

global startTime := A_Now

global P_debug := false, P_path := "", P_out := "", P_startup := false, P_append := false

paramPrint := ""
Loop %0% {
	;parse the args given to the script at startup
	; --debug : skip the startup
	; --path 'pathFile.txt' : start the script at the end of the specified path in the path file specified
	; --out 'outFile.txt' : outputs in the specified file
	; --startup : forces the script to launch ADE
	; --append : open 'outFile.txt' in append mode
	
	param := %A_Index%
	paramIndex := A_Index
	
	If (param == "--path") {
		Loop %0% {
			If (A_Index == paramIndex + 1) {
				P_path := %A_Index%
			}
		}
		If (P_path == "") {
			MsgBox, , ArboScraper, Vous n'avez pas spécifié de path.
			Stop()
		} else {
			MsgBox, , ArboScraper, % "Path reçu :" . P_path
		}
		If (InStr(P_path, "/") or InStr(P_path, "\")) {
			MsgBox, , ArboScraper, % "Le path spécifié est incorrect. Il doit faire référence à un fichier se trouvant dans le dossier du script. " . P_path
			Stop()
		}
		
	} else if (param == "--out") {
		Loop %0% {
			If (A_Index == paramIndex + 1) {
				P_out := %A_Index%
			}
		}
		If (P_out == "") {
			MsgBox, , ArboScraper, Vous n'avez pas spécifié de fichier out.
			Stop()
		} else {
			MsgBox, , ArboScraper, % "out reçu :" . P_out
		}
		
	} else if (param == "--debug") {
		P_debug := true
	} else if (param == "--startup") {
		P_startup := true
	} else if (param == "--append") {
		P_append := true
	}
	
	paramPrint := paramPrint . " " . param
}

nbParams := %0%

If (nbParams == 0) {
	;if 0 parameters were passed to the script, start in debug mode
	P_debug := true
	P_path := "testFinArbo.txt"
}


If (P_out == "") {
	P_out := "arbo_out.txt"
}

If (P_append) {
	global out := FileOpen(P_out, "a")
	
} else if (P_out == "arbo_out.txt" and P_path != "") {
	MsgBox, 4, ArboScraper, Reset arbo_out ?
	IfMsgBox,Yes
	{
		global out := FileOpen(P_out, "w")
	} else {
		global out := FileOpen(P_out, "a")
	}
	
} else {
	global out := FileOpen(P_out, "w")
}
	

If !IsObject(out) {
	MsgBox, , % "Can't open " . P_out . "!"
	ExitApp
}

DebugPrint("", "Script started with parameters: " . paramPrint, true)

global indent := "", arboX := 0, arboY := 0, arboFinX := 0, arboFinY := 0, lineHeight := 0, lineWidth := 0, devToolsX := 0, devToolsY := 0, coinScrollX := 0, coinScrollY := 0

global pauseState := false, pauseX := 0, pauseY := 0



Sleep 3000

If (!P_debug or P_startup) {
	;on ouvre ADE que si on n'est pas en mode debug
	DebugPrint("", "Starting the script in debug mode and startup mode", true)
	Startup()
}

TestADEcrash()

If WinActive("ahk_class Chrome_WidgetWin_1") {
	
	Init()
	
	SetLineDims()
	
	SetDevToolsWinSize()
	
	Main()
	
	TrayTip, ArboScraper, Finished!, 1, 1
	MsgBox, , ArboScraper, Fini!
	Stop()
} else {
	MsgBox, 4, ArboScraper, Wrong window! Launch ADE ?
	IfMsgBox, Yes 
	{
		Startup()
		Reload
	}
	Stop()
}


Startup() {
	;ouvre ADE et le DevTools dans chrome, et les met à une taille correcte
	
	CoordMode, Pixel, Screen
	
	Run, chrome.exe, , Min
	
	Sleep 2000
	
	;focus la nouvelle fenêtre créée
	ControlFocus, , Nouvel onglet - Google Chrome, Chrome Legacy Window, , Chrome Legacy Window`nChrome Legacy Window
	
	;on change la position de la fenêtre avec le racourci Win+Droite jusqu'a la position voulue
	pixelTop := 0x000000
	while(pixelTop != 0xFFFFFF) {
		ControlFocus, , Nouvel onglet - Google Chrome, Chrome Legacy Window, , Chrome Legacy Window`nChrome Legacy Window	;focus la nouvelle fenêtre
		Sleep 25
		SendInput, #{Right}
		Sleep 200
		PixelGetColor, pixelTop, 5, 5
		
		If (A_Index == 5) {
			;la fenêtre est peut-être pas étendue jusqu'en haut
			ControlFocus, , Nouvel onglet - Google Chrome, Chrome Legacy Window, , Chrome Legacy Window`nChrome Legacy Window	;focus la nouvelle fenêtre
			Sleep 25
			SendInput, #{Up}
			Sleep 20
		}
		
		If (A_Index > 10) {
			MsgBox, , ArboScraper, Unable to get the window to the correct position!
			Stop()
		}
	}
	
	Sleep 1000
	ControlFocus, , Nouvel onglet - Google Chrome, Chrome Legacy Window, , Chrome Legacy Window`nChrome Legacy Window
	Sleep 1000
	SendInput,  https://planning.univ-rennes1.fr/direct/myplanning.jsp
	Sleep 50
	SendInput, {Enter}
	Sleep 5000
	
	CoordMode, Pixel, Relative
	
	;si ADE nous demande de se connecter
	ImageSearch, loginX, loginY, 0, 0, 750, 1000, Images\ecranConnection2.png
	If (ErrorLevel == 2) {
		MsgBox, , ArboScraper, Unable to search for ecranConnection2.png!
		Stop()
	} else if (ErrorLevel == 0) {
		SendInput, {Tab 2}
		Sleep 50
		SendInput, {Enter}
		Sleep 2000
	}
	
	;on attend qu'ADE charge
	pixelADE := 0x000000
	while(pixelADE != 0x4C829B) {
		Sleep 100
		PixelGetColor, pixelADE, 10, 100, RGB
		
		If (A_Index > 100) {
			MouseMove, 10, 100
			MsgBox, , ArboScraper, % "Timeout : ADE didn't load? pixel:" . pixelADE
			Stop()
		}
	}
	
	Sleep 500
	SendInput, {F12 down}
	Sleep 50
	SendInput, {F12 up}
	Sleep 1000
	
	;waiting for DevTools
	Loop {
		Sleep 100
		WinGetActiveTitle, title
		
		If(InStr(title, "Developer Tools")) {
			Break
		}
		
		If (A_Index > 50) {
			MsgBox, , ArboScraper, %  "Unable to open or find DevTools! active window :" . title
			Stop()
		}
	}
	
	Sleep 100
	WinMove, Developer Tools, , 10, 10, 800, 256
	Sleep 250
	
	;on s'assure que l'onglet 'Elements' est actif
	ImageSearch, elementsX, elementsY, 0, 0, 800, 256, Images\devToolsElements.png
	If (ErrorLevel == 2) {
		MsgBox, , ArboScraper, Unable to search devToolsElements.png!
		Stop()
	} else if (ErrorLevel == 1) {
		MsgBox, , ArboScraper, Couldn't find devToolsElements.png!
		Stop()
	} 
	MouseMove, elementsX + 5, elementsY + 5
	Click
	Sleep 500
	
	;on enlève l'affichage des 'Styles', pour qu'il ne ralentisse pas le scraping
	ImageSearch, styleSwitchX, styleSwitchY, 400, elementsY + 5, 800, 256, *50 Images\switchStyleDevTools.png
	If (ErrorLevel == 2) {
		MsgBox, , ArboScraper, Unable to search switchStyleDevTools.png!
		Stop()
	} else if (ErrorLevel == 1) {
		MsgBox, , ArboScraper, Couldn't find switchStyleDevTools.png!
		Stop()
	}
	MouseMove, styleSwitchX + 3, styleSwitchY + 3
	Click
	Sleep 250
	MouseMove, 0, 80, 5, R		;on selectionne 'DOM BreakPoints', car il n'affiche rien du tout
	Sleep 50
	Click
	Sleep 50
	Winset, Alwaysontop, ON, A	;d'une certaine manière cela permet aussi d'accélérer le processus
	Sleep 250
	
	;on revient sur ADE
	ControlFocus, , ADE - Default, , Developer Tools
	Sleep 250
	
	;on agrandit l'arborescence
	ImageSearch, resizeStartX, resizeStartY, 200, 500, 500, 800, Images\limiteListeADE.png
	If (ErrorLevel == 2) {
		MsgBox, , ArboScraper, Unable to search for limiteListeADE.png!
		Stop()
	} else if (ErrorLevel == 1) {
		MsgBox, , ArboScraper, Couldn't find limiteListeADE.png!
		Stop()
	}
	MouseMove, resizeStartX + 5, resizeStartY, 5
	Sleep 50
	Click down
	Sleep 100
	MouseMove, 800, 0, 7, R
	Sleep 100
	Click up
	Sleep 1000
}


Init() {
	; init for arbo dims
	
	ImageSearch, arboX, arboY, 0, 0, 1000, 1000, Images\nom_liste_dossiers.png
	
	If (ErrorLevel == 2) {
		MsgBox, , ArboScraper, Could not search for image : nom_liste_dossiers.png
		Stop()
	} else if (ErrorLevel == 1) {
		MsgBox, , ArboScraper, Image not found : nom_liste_dossiers.png
		Stop()
	}
	
	arboX := arboX - 10
	arboY := arboY + 10
	
	ImageSearch, arboFinX, arboFinY, arboX, arboY, 1000, 1000, Images\scrollBasOn.png
	
	If (ErrorLevel == 2) {
		MsgBox, , ArboScraper, Could not search for image : scrollBasOn.png
		Stop()
	} else if (ErrorLevel == 1) {
		
		ImageSearch, arboFinX, arboFinY, arboX, arboY, 1000, 1000, Images\scrollBasOff.png
		
		If (ErrorLevel == 2) {
			MsgBox, , ArboScraper, Could not search for image : scrollBasOff.png
			Stop()
		} else if (ErrorLevel == 1) {
			MsgBox, , ArboScraper, Image not found : scrollBasOff.png
			Stop()
		}
	}
	
	arboFinX := arboFinX - 10 
	arboFinY := arboFinY - 10
	
	;coord du coin en bas à droite de la liste
	ImageSearch, coinScrollX, coinScrollY, arboFinX - 5, arboFinY - 5, arboFinX + 30, arboFinY + 30, Images\coinScrollBas.png
	If (ErrorLevel == 2) {
		MsgBox, , ArboScraper, Could not search for image : coinScrollBas.png
		Stop()
	} else if (ErrorLevel == 1) {
		MsgBox, , ArboScraper, Could not find image : coinScrollBas.png
		Stop()
	}
	
	coinScrollX += 3
	coinScrollY += 3
}


SetLineDims() {
	ImageSearch, flecheX, flecheY, arboX, arboY, arboFinX, arboFinY, *TransBlack Images\flecheOff.png
	
	if (ErrorLevel == 2) {
		MsgBox, , ArboScraper, Could not search for image : flecheOff.png
		Stop()
	} else if (ErrorLevel == 1) {
		MsgBox, , ArboScraper, Unable to find flecheOff.png (2)
		Stop()
	}
	
	lineHeight := flecheY
	lineWidth := flecheX
	
	MouseMove, flecheX + 3, flecheY + 3, 0
	Sleep 25
	Click
	Sleep 1000
	
	Sleep 25
	ImageSearch, flecheX, flecheY, lineWidth + 10, lineHeight + 10, arboFinX, arboFinY, *TransBlack Images\flecheOff.png
	
	if (ErrorLevel == 2) {
		MsgBox, , ArboScraper, Could not search for image : flecheOff.png (2)
		Stop()
	} else if (ErrorLevel == 1) {
		MsgBox, , ArboScraper, Unable to find flecheOff.png after %lineWidth% and %lineHeight% (2)
		Stop()
	}
	
	lineHeight := flecheY - lineHeight
	lineWidth := flecheX - lineWidth
	
	if (lineHeight < 10) {
		MsgBox, , ArboScraper, Error : lineHeight is too small (%lineHeight%)
		Stop()
	}
	if (lineWidth < 10) {
		MsgBox, , ArboScraper, Error : lineWidth is too small (%lineWidth%)
		Stop()
	}
	
	Sleep 25
	MouseClick, Left
	Sleep 500
	MouseMove, -30, 0, 0, R
}


SetDevToolsWinSize() {
	
	SendInput {F12}
	Sleep 500
	
	If (WinActive("ahk_exe chrome.exe", , "ADE - Default")) {
		;dev tools is the active window
		WinGetPos, , , devToolsX, devToolsY 
	} else {
		MsgBox, , ArboScraper, Unable to find the Developer Tools window
		Stop()
	}
	
	ControlFocus, , ADE - Default, , Developer Tools
	
	Sleep 200
}



Main() {	
	
	ControlFocus, , ADE - Default, , Developer Tools
	
	MouseMove, -5, -50, 0, R
	
	Sleep 500
	
	If (P_path != "") {
		UpdateIndent(0)
		
		pathLength := 0
		
		y := StartAtPath(pathLength)
		y := PreciseLine(y + lineHeight, 5)
		
		DebugPrint("Main", "indent : '" . indent . "' soit " . countIndents(indent) . " indentations.", true)
		
		pathLength -= 1
		
		x := arboX
		Loop %pathLength% {
			x += lineWidth
			UpdateIndent(x)
		}
		
		DebugPrint("Main", "indent final : " . countIndents(indent) . " pour un path de " . pathLength . " de longueur.", true)
		
	} else {
		y := GetFirstLine(5)
	}
	
	x := arboX
	
	progression := 0 		;0 pour début, 1 pour plus de scroll possible, 2 pour plus de lignes (donc fini)
	currentLineError := 0
	nbLignes := 0
	currentLine := ""
	siDossierEnFinDeLigne := false	;true si l'on se trouve tout à la fin de l'arborescence, chaque scroll va alors être vérifié, car ADE bug souvent
	siDossierJusteAvant := false		;si on a un fichier juste après un sous-dossier vide dans un dossier, l'indent sera foiré
	while (currentLineError < 5) {
		
		If (pauseState) {
			;on pause le script
			PauseScript()
		}
		
		isFolder := IsFolderAt(y, x, 5) 	;also updates the pos of x
		
		MouseMove, x + 3, y + 10, 0
		
		If (currentLineError == 0) {
			;on n'update l'indent et enregistre le nom du fichier qu'une seule fois
			UpdateIndent(x)
			
			currentLine := GetName(5, isFolder)
			out.WriteLine(indent . currentLine)		;get the name of the file and append it with the correct indent to arbo_out
		}
		
		If (isFolder) {
			MouseMove, -16, 0, 0, R
			Sleep 50
			Click
			
			If (!WaitFolderLoad(y)) {
				;restart this line
				DebugPrint("Main", "restarting line", true)
				y := PreciseLine(y, 5)
				currentLineError += 1
				Continue
			}
			
			If (y - arboY > (arboFinY - arboY) * 0.9) {
				;si on a un dossier sur la dernière ligne, quand on va l'ouvrir le prochain scroll sera foiré
				ImageSearch, , , arboX, y + lineHeight, arboX + 5, y + lineHeight + 3, Images\ADEblue.png
				If (ErrorLevel == 2) {
					MsgBox, , ArboScraper, ERROR : Unable to search ADEblue.png
					Stop()
				} else if (ErrorLevel == 0) {
					DebugPrint("Main", "On a un dossier en fin de liste", true)
					siDossierEnFinDeLigne := true
				}
			}
		}
		
		If (y - arboY > (arboFinY - arboY) * 0.9) {
			;need to scroll
			If (!ScrollDown()) {
				DebugPrint("Main", "no more scrolling possible", true)
				progression := 1
			} else {
				Sleep 100
				If (progression == 1) {
					progression := 0
					DebugPrint("Main", "scrolling possible! Continuing...", true)
				}
				
				y := ManageUnusualScrolls(y, currentLine, siDossierEnFinDeLigne)
			}
		}
		
		nbLignes += 1
		;debug : 20, real value : 100
		If (nbLignes > 100 and isFolder) {			
			path := getPathToRemember()
			
			WriteThePath(path)
			
			closeAllTheFolders()
			
			y := ReadThroughThePath(path)
			
			nbLignes := 0
			path :=
		}
		
		y := PreciseLine(y + lineHeight, 5)	;passe à la ligne suivante, puis recentre y sur une ligne
		
		If (progression == 1) {
			ImageSearch, , , arboX, y, arboX + 5, y + 3, Images\ADEblue.png
			If (ErrorLevel == 2) {
				MsgBox, , ArboScraper, ERROR : Unable to search ADEblue.png
				Stop()
			} else if (ErrorLevel == 0) {
				DebugPrint("Main", "ADE blue found! ENDING scraping", true)
				Break 	;fin de l'arborescence
			}
		}
		
		currentLineError := 0
		siDossierJusteAvant := isFolder
	}
	DebugPrint("Main", "nbLignes=" . nbLignes, true)
	out.WriteLine("END")
}


GetFirstLine(errNb) {
	
	ControlFocus, , ADE - Default, , Developer Tools
	
	MouseMove, arboX + 50, arboY + 50, 0
	
	errMessage := ""
	If !(WinActive("ahk_class Chrome_WidgetWin_1", , "Developer Tools")) {
		errMessage := errMessage . " - Wrong Window!"
	}
	
	ImageSearch, , y, arboX - 5, arboY - 5, arboX + lineWidth * 4, arboY + lineHeight, Images\firstLine2.png 	;init y
	If (ErrorLevel > 0 and errNb > 0) {
		;try again
		DebugPrint("GetFirstLine", "couldn't find the first line! " . errNb . errMessage, false)
		Sleep 250
		
		If (errNb == 1) {
			DebugPrint("GetFirstLine", "giving up! Returning :" . (arboY + 5), false)
			return arboY + 5
		}
		
		return GetFirstLine(errNb -1)
		
	} else if (ErrorLevel == 2) {
		MsgBox, , ArboScraper, % "ERROR : Unable to search firstLine2.png with : (" arboX - 5 ", " arboY - 5 ", " arboX + lineWidth * 4 ", " arboY + lineHeight ")"
		ControlFocus, , ADE - Default, , Developer Tools
		Sleep 1000
		DebugImageSearch(arboX - 5, arboY - 5, arboX + lineWidth * 4, arboY + lineHeight)
		Stop()
		
	} else if (ErrorLevel == 1) {
		MsgBox, , ArboScraper, % "ERROR : Couldn't find firstLinewith : (" arboX - 5 ", " arboY - 5 ", " arboX + lineWidth * 4 ", " arboY + lineHeight ")"
		ControlFocus, , ADE - Default, , Developer Tools
		Sleep 1000
		DebugImageSearch(arboX - 5, arboY - 5, arboX + lineWidth * 4, arboY + lineHeight)
		Stop()
	}
	
	return y
}


ScrollDown() {
	;si la flèche est noire, on peut scroller, sinon, non.
	
	ImageSearch, , , arboFinX, arboFinY, arboFinX + 20, arboFinY + 20, Images\scrollBasOn.png
	
	If (ErrorLevel == 2) {
		MsgBox, , ArboScraper, Could not search scrollBasOn.png
		Stop()
	} else if (ErrorLevel == 0) {
		SendInput, {WheelDown}
		return true
	} else {
		
		;on cherche si on ne trouve pas la couleur de la flèche du scroll, pour être sûr
		PixelSearch, , , arboFinX, arboFinY, coinScrollX, coinScrollY, 0x505050, 0, Fast
		If (ErrorLevel == 2) {
			MsgBox, , ArboScraper, Could not search for pixels!
			Stop()
		} else if (ErrorLevel == 1) {
			return false
		}
		SendInput, {WheelDown}
		return true
	}
}


IsFolderAt(y, ByRef x, errNb, folderCanBeOpened:=0) {
	;true if it is a folder, false if it is a file, 2 if it is an open folder
	
	;only used in ManageUnusualScrolls
	If (folderCanBeOpened) {
		ImageSearch, x, , arboX, y, arboFinX - 10, y + lineHeight, *TransBlack Images\flecheOn.png
		
		If (ErrorLevel == 2) {
			MsgBox, , ArboScraper, Could not search for image : flecheOn.png
			Stop()
		} else if (ErrorLevel == 0) {
			x += 14
			return 2
		}
	}
	
	ImageSearch, x, , arboX, y, arboFinX - 10, y + lineHeight, *TransBlack Images\flecheOff.png
	
	If (ErrorLevel == 2) {
		MsgBox, , ArboScraper, Could not search for image : flecheOff.png
		Stop()
	} else if (ErrorLevel == 0) {
		x += 14	;offset pour localiser l'icône du dossier, et donc sa position horizontale dans l'arborescence
		return true
	}
	
	ImageSearch, x, , arboX, y, arboFinX - 10, y + lineHeight, *TransBlack Images\fichierADE.png
	
	If (ErrorLevel == 2) {
		MsgBox, , ArboScraper, Could not search for image : fichierADE.png
		Stop()
	} else if (ErrorLevel == 1) {
		If (errNb != 0) {
			MouseMove, 0, -15, 0, R
			return IsFolderAt(y, x, errNb - 1) 	;retry
		} else {
			; debug
			DebugPrint("IsFolderAt", "unable to analyse line " . y . " because there is nothing to see here i swear", false)
			MsgBox, , ArboScraper, ERROR : nothing to analyse at line : y=%y% - ymax= %lineHeight%
			Stop()
		}
	}
	
	return false
}

;to remove
findStartingLineX(y) {
	y += 6
	x := arboX
	
	while (x < arboFinX) {
		x++
		PixelGetColor, pixelColor, x, y, RGB
		
		If (pixelColor != 0xFFFFFF) {
			;DebugPrint("findStartingLineX", "found line starting point at x=" . x, true)
			return x
		}
	}
	
	DebugPrint("findStartingLineX", "Couldn't find starting line point at y=" . y, false)
	MsgBox, , ArboScraper, % "Couldn't find starting line point at y=" . y
	Stop()
}


ManageUnusualScrolls(y, lineToFind, siDossierEnFinDeLigne) {
	;si la flèche de scroll en bas est noire, alors le scroll était complet puisque l'on peut toujours scroller
	;sinon, le scroll est incomplet, alors on doit rechercher la ligne où on était avant de scroller
	
	Sleep 250
	
	y := PreciseLine(y, 5)
	
	ImageSearch, , , arboFinX, arboFinY, arboFinX + 20, arboFinY + 20, Images\scrollBasOn.png
	
	If (ErrorLevel == 2) {
		MsgBox, , ArboScraper, Could not search scrollBasOn.png
		Stop()
	} else if (ErrorLevel == 0 and !siDossierEnFinDeLigne) {
		;on a monté de 3 lignes
		MouseMove, 0, - lineHeight * 3, 0, R
		y := PreciseLine(y - lineHeight * 3, 5)
		return y
	}
	
	;le scroll était incomplet, on recherche dans les 10 lignes au-dessus la ligne où l'on était
	
	x := 0
	Loop 10 {
		isFolder := IsFolderAt(y, x, 5, true)
		If (isFolder == 2)
			isFolder := true
		
		MouseMove, x + 3, y + 10
		Sleep 125
		lineName := GetName(5, isFolder)
		
		If (lineName == lineToFind) {
			y := PreciseLine(y, 5)
			MouseMove, arboX, y
			return y
		} else {
			MouseMove, 0, lineHeight, 0, R
			y := PreciseLine(y - lineHeight, 5)
		}
	}
	
	DebugPrint("ManageUnusualScrolls", "Unable to find currentLine after scrolling!", false)
	MsgBox, , ArboScraper, Unable to find lineToFind after scrolling! `n %lineToFind%
	Stop()
}



UpdateIndent(x) {
	static previousX := 1000			;preventing the first indent to be 4 spaces
	
	If (x > previousX + 5) {
		indent := indent . "    " 	;add 4 spaces at the end of the indentation, can only be one indent
		
	} else if (x < previousX) {
		nbUp := Floor((previousX - x) / lineWidth) ; nb of step ups
		
		indent := SubStr(indent, 1, StrLen(indent) - (4 * nbUp))	;removes 4 spaces of the indentation by step up
	}
	
	previousX := x
}


GetName(errNb, isFolder) {
	
	If (errNb < 0) {
		MsgBox, , ArboScraper, ERROR : unable to retrieve the name for this element...
		Stop()
	}
	
	WinGetActiveTitle, title
	If (InStr(title, "ADE - Default") == 0) {
		;wrong window!
		DebugPrint("PreciseLine", "wrong window! : " . title, false)
		ControlFocus, , ADE - Default, , Developer Tools
	}
	
	Clipboard =
	
	Sleep 25
	SendInput, ^+c			;inspect on
	Sleep 125				;prevent some lag issues
	
	MouseMove, 20, 4, 7, R	;hover file name
	Sleep 100
	
	If !(isPixelBlue(2)) {
		;pixel isn't blue, selection failed
		SendInput, ^+c		;inspect off
		Sleep 200
		MouseMove, -20, -4, 7, R
		Sleep 100
		
		; si l'inspection est inversée, on clique à côté pour la désactiver
		MouseMove, arboFinX + 50, 0, 5, R
		Sleep 100
		Click
		Sleep 100
		MouseMove, -arboFinX - 50, 0, 5, R
		Sleep 100
		
		If (errNb < 4) {
			MouseMove, 20, 0, 0, R	;on se décale, si jamais 'x' n'était pas bien positionné
		}
		
		return GetName(errNb -1, isFolder)
	}
	
	Sleep 50
	Click				;to get the element in dev tools (+ inspect off)
	Sleep 50
	
	If (WinActive("ahk_exe chrome.exe", , "ADE - Default")) {
		;dev tools is the active window
		
		while (true) {
			ImageSearch, , , 20, 20, 50, devToolsY - 30, Images\devToolsBlueSelected.png
			
			If (ErrorLevel == 2) {
				MsgBox, , ArboScraper, Unable to find devToolsBlueSelected.png!
				Stop()
			} else if (ErrorLevel == 0) {
				Break
			}
			
			Sleep 25
			
			If (A_Index > 120) {	;3 sec
				MsgBox, , ArboScraper, ERROR : Timeout while waiting for dev tools to select the element!
				Stop()
			}
		}
	}
	
	Sleep  60
	SendInput, ^c			;take the element
	ClipWait, 0.5
	
	if (ErrorLevel) {
		; the clipboard never got filled, reset and try again
		MouseMove, -20, -4, 0, R
		Sleep 25
		ControlFocus, , ADE - Default, , Developer Tools
		Sleep 200
		return GetName(errNb -1, isFolder)
	}
	
	name := Clipboard
	name := RegExReplace(name, "&amp;", "&")					;replaces the '&' in html to a normal '&'
	name := RegExReplace(name, "(<.+\"">)|(<\/span>)|(<\/div>)")	;removes the span nodes and add to the output with an indentation (remainder : " is the escape for ")
	
	MouseMove, -20, -4, 0, R	;return to initial pos
	Sleep 10
	ControlFocus, , ADE - Default, , Developer Tools				;switch to the chrome window that isn't the devtool one, aka ADE
	Sleep 10
	
	If (isFolder) {
		return name
	} else {
		return "__" . name	;add '__' before the name to mark it as a file
	}
}


WaitFolderLoad(y) {
	MouseMove, -30, 0, 0, R
	
	Sleep 100
	
	he_protec := 0
	while(true) {
		ImageSearch, , , arboX, y - 5, arboFinX, y + lineHeight, *TransBlack Images\flecheOff.png
		
		if (ErrorLevel == 2) {
			MsgBox, , ArboScraper, Unable to search flecheOff.png
			Stop()
		} else if (ErrorLevel == 1) {
			return true
		} else {
			Sleep 10
		}
		
		he_protec += 1
		if (he_protec > 100) {
			DebugPrint("WaitFolderLoad", "timeout", false)
			return false
		}
	}
}


isPixelBlue(errNb) {
	;if the pixel is quite blue -> click, else move the cursor around to force and wait for the update of the inspect tool
	
	MouseGetPos, curPosX, curPosY
	PixelGetColor, nameColor, curPosX - 2, curPosY, RGB
	
	If (ErrorLevel == 1) {
		MsgBox, , ArboScraper, Problem with pixelGetColor!
		DebugPrint("isPixelBlue", "pixelGetColor failed at " . curPosX . " - " . curPosY, false)
		Stop()
	}
	
	If (RegExMatch(nameColor, "0x(?<color>[[:xdigit:]]{2})(\k<color>){2}") == 0) {
		;pixel color isn't grey (regex didn't matched), but does it have a high value of blue ?
		;black color with blue filter : 0x496F91
		;white color with blue filter : 0xA0C6E8
		;all shades of grey and 'yellow selected' are included in this range
		
		If (nameColor >= 0x496F91 and nameColor <= 0xA0C6E8) {
			return true
		}
	}
	;else
	;pixel color is grey  (blue = red = green), and so not selected, or is not a valid shade of blue
	
	Sleep 25
	MouseMove, 50, 0, 5, R
	Sleep 25
	MouseMove, -50, 0, 5, R
	Sleep 50
	
	If (errNb > 0) {
		DebugPrint("isPixelBlue", "text wasn't blue :( -> " . errNb, false)
		return isPixelBlue(errNb -1)
	} else {
		DebugPrint("isPixelBlue", "couldn't make the text blue!", false)
		return false
	}
}


PreciseLine(y, errNb) {
	
	WinGetActiveTitle, title
	If (InStr(title, "ADE - Default") == 0) {
		;wrong window!
		DebugPrint("PreciseLine", "wrong window! : " . title, false)
		ControlFocus, , ADE - Default, , Developer Tools
	}
	
	;first errNb is always 5
	ImageSearch, , new_y, arboX, y - errNb, arboX + errNb, y + errNb, *errNb*5 Images\lineBorder.png			;on précise y en cherchant la bordure de ligne exacte autour de y (mesure de sécurité)
	If (ErrorLevel == 2) {
		MsgBox, , ArboScraper, ERROR : Unable to search lineBorder.png
		Stop()
	} else if (ErrorLevel == 1) {
		
		ImageSearch, , new_y, arboX, y - errNb, arboX + errNb, y + errNb, Images\lineSelectedBorder.png	;la ligne a peut-être été sélectionnée
		If (ErrorLevel == 2) {
			MsgBox, , ArboScraper, ERROR : Unable to search lineSelectedBorder.png
			Stop()
		} else if (ErrorLevel == 1) {
			
			If (errNb < 10) {
				DebugPrint("PreciseLine", errNb, false)
				return PreciseLine(y, errNb + 1) 	;retry with a larger area
				
			} else {				
				DebugPrint("PreciseLine", "failed to find the line at " . y, false)
				
				If (!TestADEcrash()) {
					DebugPrint("PreciseLine", "ADE n'a pas crashé", true)
					MsgBox, , ArboScraper, ERROR : Couldn't find line at %y%
					Stop()
				}
			}
		}
	}
	return new_y
}



closeAllTheFolders() {
	;searches an impossible file name, closing all of the folders in the process
	
	;MsgBox, , ArboScraper, closeAllTheFolders start..., 1
	
	Sleep 2000
	
	ImageSearch, searchX, searchY, arboFinX - 125, arboY - 100, arboFinX - 30, arboY - 20, Images\searchBar.png
	
	If (ErrorLevel == 2) {
		MsgBox, , ArboScraper, Unable to search searchBar.png!
		Stop()
	} else if (ErrorLevel == 1) {
		MsgBox, , ArboScraper, Couldn't find searchBar.png!
		
		Stop()
	}
	
	MouseMove, searchX, searchY + 10, 1
	
	Sleep 500
	
	MouseMove, arboX, searchY + 10, 1	;move to search bar
	Sleep 25
	Click
	Sleep 50
	SendInput, chose impossible à trouver pour tout refermer
	Sleep 200
	SendInput, {Enter}
	Sleep 1000
	MouseClick, Left, 0, -3, 1, 1, , R	;select all the text and delete it
	Sleep 200
	MouseClick, Left
	Sleep 200
	MouseClick, Left
	Sleep 500
	SendInput, {Delete}
	
	;wait for all the folders to close (0 éléments trouvés toast ?)
	
	waiting := true
	while (waiting) {
		ImageSearch, , , arboX - 25, arboFinY - 40, arboFinX + 10, arboFinY + 50, Images\zeroElementTrouveRecherche.png
		
		If (ErrorLevel == 2) {
			MsgBox, , ArboScraper, Unable to search zeroElementTrouveRecherche.png!
			Stop()
		} else if (ErrorLevel == 0) {
			waiting := false
		}
		
		Sleep 500
		
		If (A_Index > 90) {  ;1 min
			DebugPrint("closeAllTheFolders", "timeout", false)
			MsgBox, , ArboScraper, ERROR : Timeout for closeAllTheFolders!
			Stop()
		}
	}
	
	MouseMove, arboX, arboY, 0
	
	Sleep 5000
	
	return true
}


getPathToRemember(outFile:="") {
	;get the path to the folder where the scraper left off
	;called before 'closeAllTheFolders'
	
	out.close()	;save the file
	
	If (outFile != "") {
		outTF := TF(outFile, "outTF")
	} else {
		outTF := TF(P_out, "outTF")	;create a TF object to iterate in reverse
	}
	
	out := FileOpen(P_out, "a")	; re-open the file
	
	If !IsObject(out) {
		MsgBox, , % "Cannot append to " . P_out . "!"
		ExitApp
	}
	
	DebugPrint("getPathToRemember", "saving memory by closing the unused files...", true)
	
	fileWhereWeStoppedAt := TF_Tail(outTF, 1, 1, 0)	;get the last line, ignores every blank line and doesn't add a newline character
	
	pathToFollow := Object()		;array for the path to follow, iterated in reverse
	i := 1				;the number of lines to until the target file is reached, that will be added to the array for each sub-directory
	
	nbOf4Spaces := countIndents(fileWhereWeStoppedAt)	;get the initial number of indents	
	
	pathToFollow.Push(RegExReplace(fileWhereWeStoppedAt, "( {4})"))	;removes the indents and add the target file at the end of the path, to check if we arrived at the right place
	
	nbOfLines := TF_CountLines(outTF)
	RIndex := 1 	;reverse index of the file
	while (RIndex -1 < nbOfLines) {	;while we're not at the start of the file
		RIndex++
		currentLine := TF_Tail(outTF, - RIndex, 1)		;read the 'RIndex' line from the end of the file
		
		If (InStr(currentLine, "ERROR_", true) or InStr(currentLine, "INFO_", true)) {
			;it is an error or an info line, so it must be ignored
			Continue
		}
		
		currentNbOf4Spaces := countIndents(currentLine)
		
		If (currentNbOf4Spaces == nbOf4Spaces) {
			;une ligne du dossier contenant le dossier où l'on se trouvait
			lastValidLine := currentLine
			i++
		} else if (currentNbOf4Spaces < nbOf4Spaces) {
			;on arrive dans un dossier plus bas dans l'arborescence, celui qui nous contenait
			;on se déplace donc avec 4 espaces de moins
			pathToFollow.Push(i)
			pathToFollow.Push(RegExReplace(currentLine, "( {4})"))
			nbOf4Spaces := currentNbOf4Spaces
			i := 1
		}
		
		If (A_Index > 10000) {
			DebugPrint("getPathToRemember", "Unable to create path! Stuck in while-loop!", false)
			MsgBox, , ArboScraper, ERROR : Unable to create path! Stuck in while-loop!
			Stop()
		}
	}
	
	;if the last line isn't a number, then add 0 to the end (refers to the first directory of the arborescence)
	lastPath := pathToFollow[pathToFollow.MaxIndex()]
	
	If lastPath is not digit
	{
		DebugPrint("getPathToRemember", "last line of path wasn't a number! '" . pathToFollow[pathToFollow.MaxIndex()] . "' added " . i - 1 . " to the path.", false)
		pathToFollow.Push(i - 1) 	;since the path starts at the first folder
	} else 
		DebugPrint("getPathToRemember", "last line is a number : '" . pathToFollow[pathToFollow.MaxIndex()] . "'", true)
	
	
	;checking path length
	If (Mod(pathToFollow.Length(), 2) != 0) {
		DebugPrint("getPathToRemember", "the path has an incorrect length : " . pathToFollow.Length(), false)
		MsgBox, , ArboScraper, the path has an incorrect length!, 3
	}
	
	/*
		disp_pathtoFollow := ""
		i := 1
		while(i < pathToFollow.MaxIndex()) {
			disp_pathtoFollow := "" . disp_pathtoFollow . "`n" . pathToFollow[i]
			i++
		}
		
		Clipboard := disp_pathtoFollow
		MsgBox, , ArboScraper, %disp_pathtoFollow%, 2
	*/
	
	return pathToFollow
}


countIndents(str) {
	i := -1
	pos := 1
	while (true) {
		i += 1
		pos := InStr(str, "    ", false, pos) + 4
		
		If (pos <= 4) 
			Break	;InStr returned 0
	}
	return i
}


ReadThroughThePath(path) {
	;parcourt le 'path' pour retourner à l'endroit où on s'est arrêté
	
	Sleep 500
	
	y := arboY + 5 
	
	MouseMove, arboX + 3, y + 10
	
	Sleep 500
	
	i := path.MaxIndex()
	Loop {
		
		If (pauseState) {
			;on pause le script
			PauseScript()
		}
		
		;move to next file
		Loop, % path[i] {
			; move n times and scroll if needed
			If (y - arboY > (arboFinY - arboY) * 0.9) {
				;need to scroll
				If (!ScrollDown()) {
					DebugPrint("ReadThroughThePath", "no more scrolling possible", true)
				} else {
					MouseMove, 0, - lineHeight * 3, 0, R
					y -= lineHeight * 3
				}
			}
			
			y := PreciseLine(y + lineHeight, 5)
			
			MouseMove, arboX, y, 0
			
			Sleep 50	;unuseful?????
		}
		
		;inspect the folder before continuing
		isFolder := IsFolderAt(y, x, 5) 	;also updates the pos of x
		
		MouseMove, x + 3, y + 10, 0
		folderName := GetName(5, isFolder)
		
		If (path[i -1] != folderName) {
			DebugPrint("ReadThroughThePath", "'" folderName "' isn't the wanted '" path[i -1] "' cannot continue!", false)
			MsgBox, , ArboScraper, % "ERROR : '" folderName "' isn't the wanted '" path[i -1] "' cannot continue!"
			Stop()
		}
		
		If !(isFolder) {
			DebugPrint("ReadThroughThePath", "'" folderName "' isn't a folder name! Expected : " path[i -1], false)
			MsgBox, , ArboScraper, % "ERROR : '" folderName "' isn't a folder name! Expected : " path[i -1]
			Stop()
		}
		
		MouseMove, -16, 0, 0, R
		Sleep 50
		
		Click	;click on the folder arrow
		
		WaitFolderLoad(y)
		
		MouseMove, arboX, y, 0
		
		;change to the next sub-folder
		i -= 2
		If (i <= 0)
			Break
	}
	
	DebugPrint("ReadThroughThePath", "finished", true)
	
	;MsgBox, , ArboScraper, % "Done resuming the scraping at " path[1], 1
	return y
}


WriteThePath(path, name:="path.txt") {
	;saves the path to a file, for debug purposes
	
	pathFile := FileOpen(name, "w")
	
	If !IsObject(pathFile) {
		MsgBox, , Can't open pathFile !
		ExitApp
	}
	
	for i, thing in path {
		pathFile.WriteLine("" . thing)
	}
	
	pathFile.close()
	
	DebugPrint("WriteThePath", "sucessfully written the path to " . name, true)
}


ReadThePath(nameOfPathFile) {
	;returns the path contained by 'nameOfPathFile'
	
	pathFile := FileOpen(nameOfPathFile, "r")
	
	path := Object()
	
	If !IsObject(pathFile) {
		MsgBox, , Can't open %nameOfPathFile% !
		ExitApp
	}
	
	while (!pathFile.AtEOF()) {
		path.Push(pathFile.ReadLine())
	}
	
	pathFile.close()
	
	;cleaning the path
	toRemove := Object()
	for i, thing in path {
		If (thing == "") {
			toRemove.Push(A_Index)
		}
		;removes all new line characters
		path[i] := StrReplace(thing, "`n")
		path[i] := StrReplace(path[i], "`r")
	}
	
	for i, thing in toRemove {
		DebugPrint("ReadThePath", "removed item n°" . i . " in path", true)
		path.RemoveAt(thing)
	}
	
	;checking if path is valid
	shouldBeAStr := false
	for i, thing in path {
		shouldBeAStr := !shouldBeAStr
		If (shouldBeAStr) {
			if thing is digit
			{
				DebugPrint("ReadThePath", "the path read from " . nameOfPathFile . " is wrong because of " . thing . " isn't a string.", false)
				MsgBox, , ArboScraper, % "ERROR_ReadThePath:the path read from " . nameOfPathFile . " is wrong because of " . thing . " isn't a string."
				Stop()
			}
		} else {
			if thing is not digit
			{
				DebugPrint("ReadThePath", "the path read from " . nameOfPathFile . " is wrong because of " . thing . " isn't a number.", false)
				MsgBox, , ArboScraper, % "ERROR_ReadThePath:the path read from " . nameOfPathFile . " is wrong because of " . thing . " isn't a number."
				Stop()
			} 
		}
	}
	
	If (Mod(path.Length(), 2) != 0) {
		;path isn't even.
		DebugPrint("ReadThePath", "the path is uneven!", false)
		MsgBox, , ArboScraper, ERROR_ReadThePath: the path is uneven!
		Stop()
	}
	
	;everything is fine
	DebugPrint("ReadThePath", "the path in " . nameOfPathFile . " is valid! Success!", true)
	
	return path
}



TestADEcrash() {
	;si le scrolling d'ADE a crashé, alors il n'y a plus que du blanc à la place de la liste
	;dans ce cas on redémarre ADE et on supprime toutes les lignes de fichier jusqu'au dernier dossier,
	;pour repartir à partir de celui-ci
	
	ImageSearch, , , arboX, arboY, arboFinX, arboFinY, Images\testADEcrash.png
	If (ErrorLevel == 2) {
		MsgBox, , ArboScraper, Unable to search for testADEcrash.png!
		Stop()
	} else if (ErrorLevel == 1) {
		;return false ;ADE n'a pas crash
	}
	
	DebugPrint("TestADEcrash", "ADE crashed! Deleting lines until last folder...", false)
	
	out.close()
	
	outTF := TF(P_out, "outTF")	;on overwrite out directement
	
	lastLine := ""
	RIndex := 1
	Loop {
		RIndex++
		lastLine := TF_Tail(outTF, - RIndex, 1)
		
		If (!(InStr(lastLine, "ERROR_") or InStr(lastLine, "INFO_") or InStr(lastLine, " __"))) {
			;c'est un dossier, on supprime toutes les lignes après
			outTF := TF_RemoveLines(outTF, - RIndex)
			outTF := TF_InsertSuffix(outTF, -1, , "`n")
			Break
		}
	}
	
	TF_Save(outTF, "outCrash.txt", 1)
	
	outTF :=
	
	out := FileOpen("outCrash.txt", "a")	; on ouvre le fichier out que l'on vient de modifier
	
	If !IsObject(out) {
		MsgBox, , % "Cannot append to outCrash.txt !"
		ExitApp
	}
	
	DebugPrint("TestADEcrash", "outCrash.txt have been created!", true)
	
	Sleep 1000
	
	path := getPathToRemember("outCrash.txt")
	
	WriteThePath(path, "pathCrash.txt")
	
	Sleep 1000
	
	DebugPrint("TestADEcrash", "pathCrash.txt written!", true)
	
	Sleep 1000
	
	WinKill, ADE - Default, , 5, Developer Tools
	
	Sleep 1000
	
	IfWinExist, ADE - Default, , Developer Tools
	{
		SendInput, {Enter}
		Sleep 1000
	}
	
	IfWinExist, ADE - Default, , Developer Tools
	{
		WinKill
		Sleep 2000
	}
	
	IfWinExist, ADE - Default, , Developer Tools
	{
		MsgBox, , ArboScraper, Cannot close ADE!
		Stop()
	}
	
	params := ""
	If (P_path != "")
		params := "--path " . P_path
	If (P_out != "")
		params := params . " --out " . P_out
	If (P_debug)
		params := params . " --debug"
	
	params := params . " --startup --append"
	
	DebugPrint("TestADEcrash", "Script will be restarted with the following parameters: " . params, true)
	
	out.close()
	
	MsgBox, , ArboScraper, Last message before restart!
	
	Run,"%A_AhkPath%" /restart "%A_ScriptFullPath%" %params%
}



StartAtPath(ByRef pathLength) {
	;fait commencer le script où le fichier path de 'P_path' nous mène	
	path := ReadThePath(P_path)
	
	ControlFocus, , ADE - Default, , Developer Tools
	
	pathLength := Floor(path.Length() / 2)
	
	return ReadThroughThePath(path)
}



PauseScript() {
	MouseGetPos, pauseX, pauseY
	
	out.close()
	
	Pause on
}



DebugImageSearch(x1, y1, x2, y2) {
	;debug
	
	TrayTip, ArboScraper, % "DebugImageSearch Start with " x1 ", " y1, 1, 1 
	;Sleep, 1000
	;TrayTip
	
	Sleep 1000
	
	MouseMove, x1, y1, 5
	Sleep 500
	MouseMove, x1, y2, 5
	Sleep 500
	MouseMove, x2, y2, 5
	Sleep 500
	MouseMove, x2, y1, 5
	Sleep 500
	MouseMove, x1, y1, 5
	Sleep 500
	MouseMove, -30, 0, 0, R
	Sleep 250
}


DebugPrint(function, msg, isINFO) {
	If (P_debug) {
		If (isINFO) {
			out.WriteLine("INFO_" . function . ": " . msg)
		} else {
			out.WriteLine("ERROR_" . function . ": " . msg)
		}
	}
}



Pause::
	pauseState := !pauseState
	;on remet le curseur là où il était si on repend
	if (!pauseState) {
		ControlFocus, , ADE - Default, , Developer Tools
		Sleep, 500
		MouseMove, pauseX, pauseY, 5
		Sleep 1000
		Pause off
		
		out := FileOpen(P_out, "a")	; re-open the file
		
		If !IsObject(out) {
			MsgBox, , % "Cannot append to " . P_out . "!"
			ExitApp
		}
	}
return

Escape::
Stop() 
return

Stop() {
	endTime := A_Now
	EnvSub, endTime, startTime, S
	out.WriteLine("INFO_time=" . endTime . "  - from " . startTime)
	out.Close()
	If (endTime > 15) {
		TrayTip, ArboScraper, Finished!, 1, 1
	}
	ExitApp
}