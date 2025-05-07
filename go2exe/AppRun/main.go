package main

import (
	"bytes"
	"fmt"
	"log"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"syscall"
)

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
	// 执行 uv python list 命令
	cmd := exec.Command("cmd", "/c", "uv", "python", "list")
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

// 安装uv
func installUV(exeDir string) error {
	// 设置环境变量 INSTALLER_DOWNLOAD_URL
	os.Setenv("INSTALLER_DOWNLOAD_URL", filepath.Join(exeDir, "uv"))
	log.Printf("正在安装 UV，使用本地路径: %s", filepath.Join(exeDir, "uv"))

	// 执行 uv-installer.ps1 脚本
	cmd := exec.Command("powershell", "-ExecutionPolicy", "ByPass", "-File", filepath.Join(exeDir, "uv", "uv-installer.ps1"))
	// 隐藏窗口
	cmd.SysProcAttr = &syscall.SysProcAttr{
		HideWindow: true,
	}

	var outBuf bytes.Buffer
	cmd.Stdout = &outBuf
	cmd.Stderr = &outBuf

	err := cmd.Run()
	if err != nil {
		log.Printf("UV 安装失败: %v, 输出: %s", err, outBuf.String())
	} else {
		log.Printf("UV 安装成功，输出: %s", outBuf.String())
	}
	return err
}

// 安装Python3.11.9
func installPython(exeDir string) error {
	localMirror := "file:///" + filepath.Join(exeDir, "python")
	log.Printf("正在安装 Python 3.11.9，使用本地镜像: %s", localMirror)

	cmd := exec.Command("uv", "python", "install", "3.11.9", "--mirror", localMirror)
	// 隐藏窗口
	cmd.SysProcAttr = &syscall.SysProcAttr{
		HideWindow: true,
	}

	var outBuf bytes.Buffer
	cmd.Stdout = &outBuf
	cmd.Stderr = &outBuf

	err := cmd.Run()
	if err != nil {
		log.Printf("Python 安装失败: %v, 输出: %s", err, outBuf.String())
	} else {
		log.Printf("Python 安装成功，输出: %s", outBuf.String())
	}
	return err
}

// 运行Python应用
func runPythonApp() error {
	// 进入当前目录中的python目录
	err := os.Chdir("python")
	if err != nil {
		log.Printf("无法进入python目录: %v", err)
		return fmt.Errorf("无法进入python目录: %v", err)
	}

	log.Printf("正在启动 Python 应用")

	// 执行Python应用
	cmd := exec.Command(".\\.venv\\Scripts\\pythonw.exe", "app.pyw", "--default-index", "https://pypi.tuna.tsinghua.edu.cn/simple")
	// 这里不要隐藏窗口，因为是启动真正的应用程序

	var outBuf bytes.Buffer
	cmd.Stdout = &outBuf
	cmd.Stderr = &outBuf

	err = cmd.Run()
	if err != nil {
		log.Printf("Python 应用启动失败: %v, 输出: %s", err, outBuf.String())
	} else {
		log.Printf("Python 应用已启动")
	}
	return err
}

func main() {
	// 获取可执行文件的完整路径
	exePath, err := os.Executable()
	if err != nil {
		log.Printf("无法获取可执行文件路径: %v", err)
		return
	}
	exeDir := filepath.Dir(exePath)
	log.Printf("程序所在目录: %s", exeDir)

	// 第一步：检查是否安装了uv
	uvInstalled, output := isUVInstalled()
	log.Printf("uv安装状态: %v, 输出: %s", uvInstalled, output)

	// 如果未安装uv，则安装
	if !uvInstalled {
		log.Printf("正在安装uv...")
		err = installUV(exeDir)
		if err != nil {
			log.Printf("安装uv失败: %v", err)
			return
		}
		log.Printf("uv安装完成")

		// 重新检查uv安装状态
		uvInstalled, _ = isUVInstalled()
		if !uvInstalled {
			log.Printf("安装后仍无法检测到uv，请检查安装过程")
			return
		}
	} else {
		log.Printf("uv已安装，跳过安装步骤")
	}

	// 检查是否安装了Python3.11.9
	pythonInstalled, err := isPython3119Installed()
	if err != nil {
		log.Printf("检查Python安装状态失败: %v", err)
		return
	}

	// 如果未安装Python3.11.9，则安装
	if !pythonInstalled {
		log.Printf("正在安装Python 3.11.9...")
		err = installPython(exeDir)
		if err != nil {
			log.Printf("安装Python 3.11.9失败: %v", err)
			return
		}
		log.Printf("Python 3.11.9安装完成")
	} else {
		log.Printf("Python 3.11.9已安装，跳过安装步骤")
	}

	// 运行Python应用
	log.Printf("正在运行Python应用...")
	err = runPythonApp()
	if err != nil {
		log.Printf("运行Python应用失败: %v", err)
		return
	}
}
