package main

import (
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"time"
)

func main() {
	// 设置UTF-8编码
	cmd := exec.Command("chcp", "65001")
	cmd.Stdout = io.Discard
	cmd.Stderr = io.Discard
	cmd.Run()

	// 获取用户名
	username := os.Getenv("USERNAME")
	if username == "" && runtime.GOOS == "windows" {
		username = os.Getenv("USERPROFILE")
		if username != "" {
			username = filepath.Base(username)
		}
	}

	// 检查 uv.exe 是否存在
	uvPath := filepath.Join("C:", "Users", username, ".local", "bin", "uv.exe")
	_, err := os.Stat(uvPath)

	if err == nil {
		fmt.Println("uv 已安装，直接运行程序...")
		err = os.Chdir("python")
		if err != nil {
			fmt.Printf("无法切换到python目录: %v\n", err)
			os.Exit(1)
		}

		// 检查 .venv 目录是否存在
		venvPath := filepath.Join(".venv")
		venvInfo, venvErr := os.Stat(venvPath)
		if venvErr == nil && venvInfo.IsDir() {
			// .venv 存在，直接运行程序
			runStep4()
		} else {
			// .venv 不存在，先同步依赖再运行程序
			fmt.Println("未检测到 .venv 环境，开始使用uv恢复环境...")
			cmd = exec.Command("uv", "sync", "--default-index", "https://pypi.tuna.tsinghua.edu.cn/simple")
			cmd.Stdout = os.Stdout
			cmd.Stderr = os.Stderr
			err = cmd.Run()
			if err != nil {
				fmt.Printf("同步依赖失败: %v\n", err)
				os.Exit(1)
			}
			// 检查.venv\Scripts\pythonw.exe是否存在且可执行，最多等待30秒
			venvPythonw := filepath.Join(".venv", "Scripts", "pythonw.exe")
			timeout := time.After(30 * time.Second)
			ticker := time.NewTicker(100 * time.Millisecond)
			defer ticker.Stop()
			found := false
			for {
				select {
				case <-timeout:
					fmt.Println("等待pythonw.exe超时，环境可能未正确创建。")
					os.Exit(1)
				case <-ticker.C:
					info, err := os.Stat(venvPythonw)
					if err == nil && !info.IsDir() {
						found = true
					}
				}
				if found {
					break
				}
			}
			runStep4()
		}
	} else {
		fmt.Println("uv 未安装，开始安装运行环境")

		// 第一步：从本地安装uv
		fmt.Println("开始安装uv...")
		err = os.Chdir("uv")
		if err != nil {
			fmt.Printf("无法切换到uv目录: %v\n", err)
			os.Exit(1)
		}

		// 获取当前工作目录作为INSTALLER_DOWNLOAD_URL
		currentDir, err := os.Getwd()
		if err != nil {
			fmt.Printf("无法获取当前目录: %v\n", err)
			os.Exit(1)
		}

		// 设置环境变量 INSTALLER_DOWNLOAD_URL
		os.Setenv("INSTALLER_DOWNLOAD_URL", currentDir)

		// 执行 uv-installer.ps1 脚本
		cmd = exec.Command("powershell", "-ExecutionPolicy", "ByPass", "-File", ".\\uv-installer.ps1")
		cmd.Stdout = os.Stdout
		cmd.Stderr = os.Stderr
		err = cmd.Run()
		if err != nil {
			fmt.Printf("执行uv-installer.ps1失败: %v\n", err)
			os.Exit(1)
		}

		// 返回上一级目录，进入 python 目录
		err = os.Chdir("..\\python")
		if err != nil {
			fmt.Printf("无法切换到python目录: %v\n", err)
			os.Exit(1)
		}

		// 第二步：安装Python环境
		fmt.Println("开始安装Python环境...")
		// 获取当前工作目录并转换路径格式
		currentDir, err = os.Getwd()
		if err != nil {
			fmt.Printf("无法获取当前目录: %v\n", err)
			os.Exit(1)
		}
		localMirror := "file:///" + strings.ReplaceAll(currentDir, "\\", "/")

		// 安装Python 3.11.9
		cmd = exec.Command("uv", "python", "install", "3.11.9", "--mirror", localMirror)
		cmd.Stdout = os.Stdout
		cmd.Stderr = os.Stderr
		err = cmd.Run()
		if err != nil {
			fmt.Printf("安装Python失败: %v\n", err)
			os.Exit(1)
		}

		// 第三步：使用uv恢复环境
		fmt.Println("开始使用uv恢复环境...")
		// 同步依赖，指定清华镜像源
		cmd = exec.Command("uv", "sync", "--default-index", "https://pypi.tuna.tsinghua.edu.cn/simple")
		cmd.Stdout = os.Stdout
		cmd.Stderr = os.Stderr
		err = cmd.Run()
		if err != nil {
			fmt.Printf("同步依赖失败: %v\n", err)
			os.Exit(1)
		}

		// 执行第四步
		runStep4()
	}
}

func runStep4() {
	fmt.Println("开始运行程序...")
	cmd := exec.Command(".\\.venv\\Scripts\\pythonw.exe", "app.pyw")
	cmd.Stdout = os.Stdout
	cmd.Stderr = os.Stderr
	err := cmd.Run()
	if err != nil {
		fmt.Printf("运行app.py失败: %v\n", err)
		os.Exit(1)
	}
}
