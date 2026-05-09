Dim xl, wb, fso, folder, plik
Set fso = CreateObject("Scripting.FileSystemObject")

folder = fso.GetParentFolderName(WScript.ScriptFullName)
plik = folder & "\FormularzPomiarowy.xlsm"

If Not fso.FileExists(plik) Then
    MsgBox "Nie znaleziono formularza: " & plik, vbCritical
    WScript.Quit
End If

Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.ScreenUpdating = False

On Error Resume Next
' Otwieramy plik, makro pokaże formularz. Skrypt czeka w tym miejscu.
Set wb = xl.Workbooks.Open(plik)

' Kiedy formularz zniknie, następuje twardy reset pamięci.
If Not wb Is Nothing Then
    wb.Saved = True
    wb.Close False
End If

' Wymuszamy zamknięcie od strony systemu
xl.Quit

' Odpinamy wszystkie obiekty, by Windows mógł usunąć proces z tła
Set wb = Nothing
Set xl = Nothing
Set fso = Nothing

WScript.Quit