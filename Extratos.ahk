F9::						; AUTOMATIZAÇÃO DE ENTRADA DE EXTRATOS CONTÁBEIS MENSAIS
StartTime := A_TickCount
CoordMode, Mouse, Window
IfWinExist, Gestão Empresarial
	{
		InputBox, LoteInput, Nº Lote, Entre com o Nº do lote:,,200,100
		if ErrorLevel 
			{
				MsgBox, Operação cancelada.
				return
			}
		WinActivate
		WinWait, Gestão Empresarial
		Send, {BACKSPACE}
		MouseClick, left, 166,94
		Send, %LoteInput%
		MouseClick, left,890, 92
		Sleep 500
		Send, {DOWN}
	}
Else 
	{
		MsgBox, Error, Senior not running.
		return
	}
InputBox, DataInput, Data, Por favor entrar com o mês e ano,,200,100
if ErrorLevel 
	{
		MsgBox, Operação cancelada.
		return
	}	
else if (StrLen(DataInput)<7)
		{
			MsgBox, Data Inválida.
			return
		}
else
	MsgBox, Mês e ano definido como: "%DataInput%"

FileCopy, C:\Users\Processos\Desktop\DADOSCSV.csv, C:\Users\Processos\Desktop\DADOS.csv
Sleep, 2000
CoordMode, Mouse, Window
SetTitleMatchMode, 2
IfWinNotExist, Notepad++
	Run C:\Program Files\Notepad++\notepad++.exe
WinWait, Notepad++
IfWinExist, Notepad++
	{
		WinActivate
		WinMaximize
	}
IfWinNotActive, DADOS.csv
	{
		Send, {CONTROL DOWN}{o}{CONTROL UP}
		WinWait, Abrir
		Send, C:\Users\Processos\Desktop\DADOS.csv
		MouseClick, Left, 508, 446
	}
MsgBox,4,Warning, Os dados estão corretos?
IfMsgBox No
	Return
Sleep, 30
Send, {CONTROL DOWN}{HOME}{CONTROL UP}
Send, {HOME}{SHIFT DOWN}{END}{SHIFT UP}{DEL}{DEL}				
Send, {CONTROL DOWN}{h}{CONTROL UP}
WinWait, Substituir
Send, `;;;;;;;
Send, {TAB}{BACKSPACE}
MouseClick, Left, 476, 134
Sleep, 100
Send, R${Click}
Sleep, 50
Send, .{Click}
Sleep,50
Send, {ESC}
MouseClick, Left, 86, 41
MouseMove, 230, 317
Sleep, 750
MouseMove, 505,0,,Relative
MouseClick, Left, 705, 473
Send, {CONTROL DOWN}{s}{CONTROL UP}
Loop, Read, C:\Users\Processos\Desktop\DADOS.csv
	{
		if (BreakLoop = 1)
			{
				break 
			}
		total_lines = %A_Index%
	}
Send, {CONTROL DOWN}{HOME}{CONTROL UP}
Send, {UP}
Loop, %total_lines%
	{
		if (BreakLoop = 1)
			{
				break 
			}
		Send, {CONTROL DOWN}{F9}{CONTROL UP}
	}
Send, {CONTROL DOWN}{HOME}{CONTROL UP}
Send, {DOWN 3}
clipboard = %DataInput% ; 										DIGITE NESTA LINHA O MÊS E ANO NO FORMATO XX/XXXX
Loop, %total_lines%
	{
		if (BreakLoop = 1)
			{
				break 
			}
		Send, {CONTROL DOWN}{F10}{CONTROL UP}
	}
Send, {CONTROL DOWN}{s}{CONTROL UP}
Array := []
ArrayCount := 0
SetTitleMatchMode, 2
Loop % total_lines * 5
	{
		if (BreakLoop = 1)
			{
				break 
			}
		Sleep, 15			
		if WinExist("DADOS.csv - Notepad++")
			{
				ArrayCount += 1
				SetKeyDelay 0,0
				WinActivate  ; Automatically uses the window found above.
				Send, {CONTROL DOWN}{Home}{CONTROL UP}
				Sleep, 10
				Send, {SHIFT DOWN}{END}{SHIFT UP}
				Sleep, 10
				Send, {CONTROL DOWN}{x}{CONTROL UP}
				Send, {Delete}
				Array[A_Index] := clipboard
			}
	}
counter := 0
Sleep, 200
if WinExist("Senior |")
	{
		SetKeyDelay 20,15
		WinActivate  ; Automatically uses the window found above.
		Sleep, 300
		WinWaitActive
		Loop, % total_lines
			{
				if (BreakLoop = 1)
					{
						break 
					}
				counter := % counter + 1
				Send, {Enter}
				Sleep, 50
				Send % Value := Array[counter]		;INSERÇÃO DA CONTA DEBITADA
				counter := % counter + 1
				Send, {Enter}
				Sleep, 50
				Send % Value := Array[counter]		;INSERÇÃO DA CONTA CREDITADA
				counter := % counter + 1
				Send, {Enter}
				Sleep, 100
				Send % Value := Array[counter]   ;INSERÇÃO DO VALOR
				counter := % counter + 1
				Send, {Enter}
				Sleep, 50
				Send % Value := Array[counter]			;INSERÇÃO DA DATA
				counter := % counter + 1
				Send, {Enter}
				Sleep, 20
				Send, {Enter}
				Sleep, 50
				Send % Value := Array[counter]		;INSERÇÃO  DO MOTIVO
				Sleep, 50
				Send,  {Enter 2}
				Sleep, 100
			}
	}

;
ElapsedTime := A_TickCount - StartTime
MsgBox,  %ElapsedTime% milliseconds have elapsed.
Return

#^R::
MsgBox, 4,Script, Reload?
IfMsgBox, No
	{
		MsgBox, Operation Cancelled.
	}
IfMsgBox, Yes
	{	
		FileDelete, C:\Users\Processos\Desktop\DADOS.csv
		MsgBox, Script Reloaded.
		Reload
		Return
	}
	
F10::
BreakLoop = 1                  ; PARA SAIR DO LOOP, MODIFICAR PARA A TECLA DE PREFERENCIA
return
