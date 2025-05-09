package main

import (
	"bufio"
	"bytes"
	"fmt"
	"io"
	"log"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"sync"
	"syscall"
	"time"
	"unsafe"
)

// Windows API 函数声明
var (
	user32             = syscall.NewLazyDLL("user32.dll")
	kernel32           = syscall.NewLazyDLL("kernel32.dll")
	messageBox         = user32.NewProc("MessageBoxW")
	createWindowEx     = user32.NewProc("CreateWindowExW")
	defWindowProc      = user32.NewProc("DefWindowProcW")
	registerClassEx    = user32.NewProc("RegisterClassExW")
	getModuleHandle    = kernel32.NewProc("GetModuleHandleW")
	postQuitMessage    = user32.NewProc("PostQuitMessage")
	showWindow         = user32.NewProc("ShowWindow")
	updateWindow       = user32.NewProc("UpdateWindow")
	destroyWindow      = user32.NewProc("DestroyWindow")
	getMessage         = user32.NewProc("GetMessageW")
	translateMessage   = user32.NewProc("TranslateMessage")
	dispatchMessage    = user32.NewProc("DispatchMessageW")
	setWindowText      = user32.NewProc("SetWindowTextW")
	createWindowExW    = user32.NewProc("CreateWindowExW")
	sendMessage        = user32.NewProc("SendMessageW")
	getWindowRect      = user32.NewProc("GetWindowRect")
	createFont         = user32.NewProc("CreateFontW")
	getStdHandle       = kernel32.NewProc("GetStdHandle")
	allocConsole       = kernel32.NewProc("AllocConsole")
	freeConsole        = kernel32.NewProc("FreeConsole")
	writeConsole       = kernel32.NewProc("WriteConsoleW")
	setConsoleTitle    = kernel32.NewProc("SetConsoleTitleW")
	MB_OK              = 0x00000000
	MB_ICONINFORMATION = 0x00000040
	MB_ICONEXCLAMATION = 0x00000030
	STD_OUTPUT_HANDLE  = -11
)

// 显示消息框
func showMessageBox(title, message string) {
	titlePtr, _ := syscall.UTF16PtrFromString(title)
	messagePtr, _ := syscall.UTF16PtrFromString(message)
	messageBox.Call(
		0,
		uintptr(unsafe.Pointer(messagePtr)),
		uintptr(unsafe.Pointer(titlePtr)),
		uintptr(MB_OK|MB_ICONINFORMATION),
	)
}

var (
	hwndProgressWindow uintptr
	outputText         []string
	outputMutex        sync.Mutex
	progressTitle      string
)

// 初始化控制台窗口
func initConsole() {
	allocConsole.Call()
	titlePtr, _ := syscall.UTF16PtrFromString("安装进度")
	setConsoleTitle.Call(uintptr(unsafe.Pointer(titlePtr)))
}

// 关闭控制台窗口
func closeConsole() {
	freeConsole.Call()
}

// 写入控制台
func writeToConsole(text string) {
	handle, _, _ := getStdHandle.Call(uintptr(STD_OUTPUT_HANDLE))
	textPtr, _ := syscall.UTF16FromString(text + "\r\n")
	var written uint32
	writeConsole.Call(
		handle,
		uintptr(unsafe.Pointer(&textPtr[0])),
		uintptr(len(textPtr)),
		uintptr(unsafe.Pointer(&written)),
		0,
	)
}

func init() {
	// 创建日志文件
	logFile, err := os.OpenFile("app.log", os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0666)
	if err != nil {
		// 如果无法创建日志文件，继续执行但不记录日志
		return
	}
	log.SetOutput(logFile)
}

// 检查是否安装了uv
func isUVInstalled() (bool, string) {
	// 执行 uv -V 命令
	cmd := exec.Command("uv", "-V")
	// 隐藏窗口
	cmd.SysProcAttr = &syscall.SysProcAttr{
		HideWindow: true,
	}

	var outBuf bytes.Buffer
	cmd.Stdout = &outBuf
	cmd.Stderr = &outBuf

	err := cmd.Run()

	// 将输出转换为字符串
	outputStr := strings.TrimSpace(outBuf.String())

	// 记录日志
	log.Printf("UV 检查结果: %v, 输出: %s", err == nil, outputStr)

	// 如果命令执行出错，或输出包含错误信息，说明未安装
	if err != nil || strings.Contains(outputStr, "不是内部或外部命令") {
		return false, outputStr
	}

	// 如果输出包含版本号，说明已安装
	if strings.Contains(outputStr, "uv") {
		return true, outputStr
	}

	return false, outputStr
}

// 检查是否安装了Python3.11.9
func isPython3119Installed() (bool, error) {
	// 执行 uv python list 命令，使用PowerShell
	cmd := exec.Command("powershell", "-NoProfile", "-NonInteractive", "-Command", "uv python list")
	// 隐藏窗口
	cmd.SysProcAttr = &syscall.SysProcAttr{
		HideWindow: true,
	}

	var outBuf bytes.Buffer
	cmd.Stdout = &outBuf
	cmd.Stderr = &outBuf

	err := cmd.Run()
	if err != nil {
		log.Printf("Python 检查命令失败: %v", err)
		return false, fmt.Errorf("执行命令失败: %v", err)
	}

	// 将输出转换为字符串并按行分割
	outputStr := outBuf.String()
	lines := strings.Split(outputStr, "\n")

	// 查找包含 Python 3.11.9 的行
	targetVersion := "cpython-3.11.9-windows-x86_64-none"
	for _, line := range lines {
		if strings.Contains(line, targetVersion) {
			// 检查行是否包含路径而不是 "<download available>"
			if !strings.Contains(line, "<download available>") {
				log.Printf("找到已安装的 Python 3.11.9")
				return true, nil
			}
		}
	}
	log.Printf("未找到已安装的 Python 3.11.9")
	return false, nil
}

// 处理命令输出的辅助函数
func processCommandOutput(reader io.ReadCloser, isError bool) {
	scanner := bufio.NewScanner(reader)
	prefix := "INFO: "
	if isError {
		prefix = "ERROR: "
	}

	for scanner.Scan() {
		line := scanner.Text()
		log.Println(prefix + line)
		addOutputText(prefix + line)
	}
}

// 添加输出文本
func addOutputText(text string) {
	outputMutex.Lock()
	defer outputMutex.Unlock()
	outputText = append(outputText, text)
	writeToConsole(text)
}

// 安装uv
func installUV(exeDir string) error {
	// 设置环境变量 INSTALLER_DOWNLOAD_URL
	os.Setenv("INSTALLER_DOWNLOAD_URL", filepath.Join(exeDir, "uv"))
	log.Printf("正在安装 UV，使用本地路径: %s", filepath.Join(exeDir, "uv"))
	addOutputText(fmt.Sprintf("正在安装 UV，使用本地路径: %s", filepath.Join(exeDir, "uv")))

	// 执行 uv-installer.ps1 脚本
	cmd := exec.Command("powershell", "-ExecutionPolicy", "ByPass", "-File", filepath.Join(exeDir, "uv", "uv-installer.ps1"))
	// 隐藏窗口
	cmd.SysProcAttr = &syscall.SysProcAttr{
		HideWindow: true,
	}

	// 获取标准输出和错误输出管道
	stdout, err := cmd.StdoutPipe()
	if err != nil {
		log.Printf("获取标准输出管道失败: %v", err)
		return err
	}

	stderr, err := cmd.StderrPipe()
	if err != nil {
		log.Printf("获取标准错误输出管道失败: %v", err)
		return err
	}

	// 启动命令
	if err := cmd.Start(); err != nil {
		log.Printf("启动UV安装命令失败: %v", err)
		return err
	}

	// 实时处理输出
	go processCommandOutput(stdout, false)
	go processCommandOutput(stderr, true)

	// 等待命令完成
	err = cmd.Wait()
	if err != nil {
		log.Printf("UV 安装失败: %v", err)
		addOutputText(fmt.Sprintf("UV 安装失败: %v", err))
	} else {
		addOutputText("UV 安装成功！")
		log.Printf("UV 安装成功")
	}
	return err
}

// 安装Python3.11.9
func installPython(exeDir string) error {
	localMirror := "file:///" + filepath.Join(exeDir, "python")
	log.Printf("正在安装 Python 3.11.9，使用本地镜像: %s", localMirror)
	addOutputText(fmt.Sprintf("正在安装 Python 3.11.9，使用本地镜像: %s", localMirror))

	cmd := exec.Command("powershell", "-NoProfile", "-NonInteractive", "-Command",
		fmt.Sprintf("uv python install 3.11.9 --mirror '%s'", localMirror))
	// 隐藏窗口
	cmd.SysProcAttr = &syscall.SysProcAttr{
		HideWindow: true,
	}

	// 获取标准输出和错误输出管道
	stdout, err := cmd.StdoutPipe()
	if err != nil {
		log.Printf("获取标准输出管道失败: %v", err)
		return err
	}

	stderr, err := cmd.StderrPipe()
	if err != nil {
		log.Printf("获取标准错误输出管道失败: %v", err)
		return err
	}

	// 启动命令
	if err := cmd.Start(); err != nil {
		log.Printf("启动Python安装命令失败: %v", err)
		return err
	}

	// 实时处理输出
	go processCommandOutput(stdout, false)
	go processCommandOutput(stderr, true)

	// 等待命令完成
	err = cmd.Wait()
	if err != nil {
		log.Printf("Python 安装失败: %v", err)
		addOutputText(fmt.Sprintf("Python 安装失败: %v", err))
	} else {
		addOutputText("Python 3.11.9 安装成功！")
		log.Printf("Python 安装成功")
	}
	return err
}

// 运行Python应用
func runPythonApp() error {
	// 进入当前目录中的python目录
	err := os.Chdir("python")
	if err != nil {
		log.Printf("无法进入python目录: %v", err)
		addOutputText(fmt.Sprintf("无法进入python目录: %v", err))
		return fmt.Errorf("无法进入python目录: %v", err)
	}

	log.Printf("正在启动 Python 应用")
	addOutputText("正在启动 Python 应用...")

	// 执行 uv sync 命令，配置清华源
	log.Printf("正在执行 uv sync 配置清华源...")
	addOutputText("正在执行 uv sync 配置清华源...")

	syncCmd := exec.Command("powershell", "-NoProfile", "-NonInteractive", "-Command",
		"uv sync --default-index 'https://pypi.tuna.tsinghua.edu.cn/simple'")
	// 隐藏窗口
	syncCmd.SysProcAttr = &syscall.SysProcAttr{
		HideWindow: true,
	}

	// 获取标准输出和错误输出管道
	stdout, err := syncCmd.StdoutPipe()
	if err != nil {
		log.Printf("获取标准输出管道失败: %v", err)
		return err
	}

	stderr, err := syncCmd.StderrPipe()
	if err != nil {
		log.Printf("获取标准错误输出管道失败: %v", err)
		return err
	}

	// 启动命令
	if err := syncCmd.Start(); err != nil {
		log.Printf("启动uv sync命令失败: %v", err)
		return err
	}

	// 实时处理输出
	go processCommandOutput(stdout, false)
	go processCommandOutput(stderr, true)

	// 等待命令完成
	err = syncCmd.Wait()
	if err != nil {
		log.Printf("uv sync 配置失败: %v", err)
		addOutputText(fmt.Sprintf("uv sync 配置失败: %v", err))
		// 尽管配置失败，仍然继续尝试启动应用
	} else {
		addOutputText("uv sync 配置成功！")
		log.Printf("uv sync 配置成功")
	}

	// 执行Python应用
	cmd := exec.Command("./.venv/Scripts/pythonw.exe", "app.pyw", "--default-index", "https://pypi.tuna.tsinghua.edu.cn/simple")
	// 这里不要隐藏窗口，因为是启动真正的应用程序

	// 获取输出以便记录可能的错误
	var outBuf bytes.Buffer
	cmd.Stdout = &outBuf
	cmd.Stderr = &outBuf

	err = cmd.Start()
	if err != nil {
		log.Printf("Python 应用启动失败: %v", err)
		addOutputText(fmt.Sprintf("Python 应用启动失败: %v", err))
		return err
	}
	// 应用成功启动，记录信息并关闭控制台
	log.Printf("Python 应用已启动")
	addOutputText("Python 应用已启动")
	closeConsole() // 主动关闭控制台
	// 不调用 cmd.Wait()，让 Python 应用独立运行
	return nil

}

func main() {
	// 获取可执行文件的完整路径
	exePath, err := os.Executable()
	if err != nil {
		log.Printf("无法获取可执行文件路径: %v", err)
		addOutputText(fmt.Sprintf("无法获取可执行文件路径: %v", err))
		return
	}
	exeDir := filepath.Dir(exePath)
	log.Printf("程序所在目录: %s", exeDir)
	addOutputText(fmt.Sprintf("程序所在目录: %s", exeDir))

	// 第一步：检查是否安装了uv
	uvInstalled, output := isUVInstalled()
	log.Printf("uv安装状态: %v, 输出: %s", uvInstalled, output)
	addOutputText(fmt.Sprintf("uv安装状态: %v", uvInstalled))

	// 如果未安装uv，则安装
	if !uvInstalled {
		// 首先弹出一个简单的消息框告知用户
		showMessageBox("环境安装", "即将开始安装必要的环境组件，请稍候...\n\n这个过程只在首次运行时执行，可能需要几分钟。")
		// 初始化控制台窗口
		initConsole()
		defer closeConsole()
		log.Printf("正在安装uv...")
		addOutputText("正在安装uv...")
		err = installUV(exeDir)
		if err != nil {
			log.Printf("安装uv失败: %v", err)
			addOutputText(fmt.Sprintf("安装uv失败: %v", err))
			time.Sleep(5 * time.Second) // 给用户时间查看错误信息
			return
		}
		log.Printf("uv安装完成")
		addOutputText("uv安装完成")

		// 重新检查uv安装状态
		uvInstalled, _ = isUVInstalled()
		if !uvInstalled {
			log.Printf("安装后仍无法检测到uv，请检查安装过程")
			addOutputText("安装后仍无法检测到uv，请检查安装过程")
			time.Sleep(5 * time.Second) // 给用户时间查看错误信息
			return
		}
	} else {
		log.Printf("uv已安装，跳过安装步骤")
		addOutputText("uv已安装，跳过安装步骤")
	}

	// 检查是否安装了Python3.11.9
	pythonInstalled, err := isPython3119Installed()
	if err != nil {
		log.Printf("检查Python安装状态失败: %v", err)
		addOutputText(fmt.Sprintf("检查Python安装状态失败: %v", err))
		time.Sleep(5 * time.Second) // 给用户时间查看错误信息
		return
	}

	// 如果未安装Python3.11.9，则安装
	if !pythonInstalled {
		log.Printf("正在安装Python 3.11.9...")
		addOutputText("正在安装Python 3.11.9...")
		err = installPython(exeDir)
		if err != nil {
			log.Printf("安装Python 3.11.9失败: %v", err)
			addOutputText(fmt.Sprintf("安装Python 3.11.9失败: %v", err))
			time.Sleep(5 * time.Second) // 给用户时间查看错误信息
			return
		}
		log.Printf("Python 3.11.9安装完成")
		addOutputText("Python 3.11.9安装完成")
	} else {
		log.Printf("Python 3.11.9已安装，跳过安装步骤")
		addOutputText("Python 3.11.9已安装，跳过安装步骤")
	}

	// 运行Python应用
	log.Printf("正在运行Python应用...")
	addOutputText("正在运行Python应用...")
	err = runPythonApp()
	if err != nil {
		log.Printf("运行Python应用失败: %v", err)
		addOutputText(fmt.Sprintf("运行Python应用失败: %v", err))
		time.Sleep(5 * time.Second) // 给用户时间查看错误信息
		return
	}
}
