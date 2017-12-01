#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#SingleInstance,Force

IfWinActive, ahk_class Chrome_WidgetWin_1 {
	
	CoordMode, Pixel, Relative
	
	ImageSearch, imgX, imgY, 0, 0, 1000, 1000, nom_liste_dossiers.png
	
	If (ErrorLevel == 2) {
		MsgBox, , ArboScraper, Could not search for image : nom_liste_dossiers.png
	} else if (ErrorLevel == 1) {
		MsgBox, , ArboScraper, Image not found : nom_liste_dossiers.png
	} else {
		MsgBox, , ArboScraper, Image found at : %imgX% - %imgY%
	}

} else 
	MsgBox, , ArboScraper, Wrong window!
