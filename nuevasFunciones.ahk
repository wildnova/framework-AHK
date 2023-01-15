#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance, Force
#Include, lib\Chrome.ahk
#Include, lib\XL.ahk
#Include, lib\jsonToAhk.ahk


FileRead, config, config.json
;Formatea el archivo a json en ahk
jsonWorkFlow:= JsonToAHK(config)

maxRetries:= jsonWorkFlow.maxRetries
pathTxt:=jsonWorkFlow.pathTxt

MsgBox, % maxRetries . "   " . pathTxt


ExitApp