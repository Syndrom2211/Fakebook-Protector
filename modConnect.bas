Attribute VB_Name = "ModConnect"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef _
lpSFlags As Long, ByVal dwReserved As Long) As Long
Const INTERNET_CONNECTION_MODEM = 1
Const INTERNET_CONNECTION_LAN = 2
Const INTERNET_CONNECTION_PROXY = 4
Const INTERNET_CONNECTION_MODEM_BUSY = 8

'Fungsi yang berguna untuk mengecek koneksi internet yang ada di komputer kita menggunakan fungsi API InternetGetConnectedState
Function CekInternet(Optional connectMode As Integer) As Boolean
Dim flags As Long
CekInternet = InternetGetConnectedState(flags, 0)
connectMode = flags
End Function
