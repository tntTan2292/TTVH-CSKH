Set WshShell = CreateObject("WScript.Shell")

' Cố định thư mục làm việc tuyệt đối
WshShell.CurrentDirectory = "d:\Antigravity - Project - TTVH\CSKH"

' Khởi động Backend API (Cổng 8010)
WshShell.Run "cmd /c python backend/main.py", 0, False

' Đợi 3 giây để Backend sẵn sàng
WScript.Sleep 3000

' Khởi động Frontend Dashboard (Cổng 8088)
WshShell.Run "cmd /c python serve_dashboard.py", 0, False
