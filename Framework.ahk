#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance, Force
#Include, lib\Chrome.ahk
#Include, lib\XL.ahk
#Include, lib\jsonToAhk.ahk

;variables globales

global ss := 200
global ms := 1000
global ls := 3000
global loadSleep := 30000
global LogFile := "logs.txt"

GetFullPathName(path) {
    cc := DllCall("GetFullPathName", "str", path, "uint", 0, "ptr", 0, "ptr", 0, "uint")
    VarSetCapacity(buf, cc*(A_IsUnicode?2:1))
    DllCall("GetFullPathName", "str", path, "uint", cc, "str", buf, "ptr", 0, "uint")
    return buf
}

esperaDinamica(numIntentosMax:=5,tiempoEspera:=5000,elementoEncontrar="",chrome="")
{
	paginaCargada:=False
	numIntentos:=1
	;MsgBox, % "numIntentosMax  " numIntentosMax " `n " "tiempoEspera   "  tiempoEspera " `n elementoEncontrar:  " elementoEncontrar
	loop, %numIntentosMax%
	{
		sleep %tiempoEspera%
        ;msgbox en el bucle en el intento: %numIntentos%
		try{
			
			page := chrome.getPage()
			sleep %ms%
			elementoEnPagina := % page.Evaluate(elementoEncontrar).Value
			;MsgBox, % elementoEnPagina
			sleep %ms%
			paginaCargada:= true

			FileAppend, "`nPágina cargada", %LogFile%
			Break
		}
		catch e
		{
			FileAppend, "`nEsperando la carga de la página . Número de intentos: " %numIntentos%, %LogFile%
			
			if(numIntentos>=numIntentosMax)
			{
				FileAppend, "`nNo se puede detectar si la pagina se ha cargado.", %LogFile%
				Throw Exception("No se puede detectar si la pagina se ha cargado.")
			}
		}
        numIntentos++
	}
}

;__________________________________________________________________________ MAIN_______________________________________________________________________________________
try
{
	;Definir fichero de logs

	if FileExist(LogFile)
		FileDelete, %LogFile%
	FileAppend, "---------------------- Inicio del proceso ---------------------------", %LogFile%, UTF-8

	;Cerrar cualquier instancia de excel abierta.
	try
    {
	    ex := ComObjActive("Excel.Application")
	    ex.quit()

		FileAppend, % "`nSe ha detectado una instancia de excel abierta, procedemos a cerrarla y continuar con el proceso", %LogFile%
    }
    Catch
    {
	    FileAppend % "`nNo se ha detectado ninguna instancia de excel abierta, continuamos con el proceso.", %LogFile%
    }

	;Definir ruta de usuario 
    EnvGet, hdrive, Homedrive ;hdrive nombre del disco local
	EnvGet, hpath, Homepath ; hpath ruta del usuario
	rutaUsuario:= hdrive . hpath
    
	;Cerrar Chrome si está abierto
	IfWinExist, ahk_exe chrome.exe
        WinClose
	sleep %ls%

	;Lectura de fichero Config
	FileRead, config, config.json
	jsonWorkFlow:= JsonToAHK(config)
	maxRetries:= jsonWorkFlow.maxRetries

	;Fichero de colas y reintentos
	if(FileExist("workflow.json"))
	{
		FileRead, workflow, workflow.json
		jsonWorkFlow:= JsonToAHK(workflow)

		phase:= jsonWorkFlow.phase  ;Fase de ejecución
		retries:=jsonWorkFlow.retries ;reintento actual
	}
	else
	{
		phase:=0
		retries:=0
	}
	if(retries>=maxRetries)
	{
		FileAppend, "`nNúmero máximo de reintentos alcanzado", %LogFile%
		ExitApp
	}

	;actualización del fichero de colas y reintentos.
	if(FileExist("workflow.json"))
	{
		FileDelete, workflow.json
	}
	FileAppend, {`n   "phase":0`,`n   "retries":0`n}, workflow.json,UTF-8

	FileAppend, "`nAutomatismo finalizado correctamente", %LogFile%
	ExitApp
}
catch e
{
	FileAppend, % "`n`nERROR: `n" . e.message . "`nExcepción lanzada en la línea: " . e.line , %LogFile%
	sleep %ss%
	retries++

	if(FileExist("workflow.json"))
	{
		FileDelete, workflow.json
	}

	sleep %ls%
	if(retries>=maxRetries)
	{
		FileAppend, "`n`nNúmero máximo de reintentos alcanzado", %LogFile%, UTF-8
		FileAppend, {`n   "phase":%phase%`,`n   "retries":0`n}, workflow.json,UTF-8

		ExitApp ;Finaliza el script
	}
	else
	{
		FileAppend, {`n   "phase":%phase%`,`n   "retries":%retries%`n}, workflow.json,UTF-8
		Reload ;Reinicia el script desde cero si llega hasta aquí
	}
	;MsgBox,% "Debug-- Error: " . e.message . "`nExcepción lanzada en la línea: " . e.line  ;Activar en pruebas para que aparezca el mensaje de error en pantalla.
	
}
	
