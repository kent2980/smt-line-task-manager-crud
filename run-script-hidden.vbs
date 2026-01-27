' run-script-hidden.vbs
' PowerShellスクリプトを非表示で実行するVBScriptラッパー

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' スクリプトのディレクトリを取得
strScriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
strMainScript = objFSO.BuildPath(strScriptPath, "Main.ps1")

' PowerShellのパス（PowerShell 7を優先、なければWindows PowerShell）
strPowerShell = "C:\Program Files\PowerShell\7\pwsh.exe"
If Not objFSO.FileExists(strPowerShell) Then
    strPowerShell = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
End If

' 作業ディレクトリをスクリプトのディレクトリに設定
objShell.CurrentDirectory = strScriptPath

' PowerShellスクリプトを非表示で実行
intReturnCode = objShell.Run("""" & strPowerShell & """ -ExecutionPolicy Bypass -NoProfile -File """ & strMainScript & """", 0, False)

' エラーレベルを返す
WScript.Quit(intReturnCode)

