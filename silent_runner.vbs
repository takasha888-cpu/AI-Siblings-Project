Set shell = CreateObject("WScript.Shell")
' 第2引数 0 はウィンドウを非表示にすることを意味します
shell.Run "pythonw.exe " & WScript.Arguments(0), 0, False
