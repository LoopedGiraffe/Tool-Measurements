' StartDodawanieNarzedziTYP3.vbs
' Uruchamia Excela, otwiera plik DodawanieDoBazyNarzedziTYP3.xlsm
' i odpala makro PokazFormularzDodajNarzedzieTYP3, które pokazuje formularz.

Option Explicit

Dim fso, folder, plik
Dim xlApp, wb

' --- ustalenie folderu, w którym leży ten skrypt ---
Set fso = CreateObject("Scripting.FileSystemObject")
folder = fso.GetParentFolderName(WScript.ScriptFullName)

' nazwa pliku Excela w tym samym folderze
plik = folder & "\DodawanieDoBazyNarzedziTYP3.xlsm"

' --- uruchomienie Excela ---
Set xlApp = CreateObject("Excel.Application")
If xlApp Is Nothing Then
    MsgBox "Nie można uruchomić Microsoft Excel.", vbCritical, "Błąd"
    WScript.Quit
End If

' Zmieniamy na False - aplikacja Excel nie będzie widoczna na pasku ani na ekranie
xlApp.Visible = False

On Error Resume Next
Set wb = xlApp.Workbooks.Open(plik)

If Err.Number <> 0 Then
    MsgBox "Nie można otworzyć pliku:" & vbCrLf & plik, vbCritical, "Błąd"
    xlApp.Quit
    WScript.Quit
End If
On Error GoTo 0

' --- uruchomienie makra pokazującego formularz ---
xlApp.Run "PokazFormularzDodajNarzedzieTYP3"

' Skrypt VBS kończy tu pracę, a proces excel.exe działa dalej w tle, 
' wyświetlając jedynie okno formularza VBA.
' Zamknięcie aplikacji Excel należy obsłużyć w samym VBA.

Set wb = Nothing
Set xlApp = Nothing
Set fso = Nothing