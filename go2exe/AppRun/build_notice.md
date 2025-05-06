1. 编译脚本
go build -ldflags "-H windowsgui" -o ..\..\SpeakMyBook.exe main.go
2. 注意修改app.pyw，在开头添加以下代码，避免路径问题：
logfile = os.path.join(os.path.dirname(__file__), "app.log")
sys.stdout = open(logfile, "a", encoding="utf-8")
sys.stderr = open(logfile, "a", encoding="utf-8")