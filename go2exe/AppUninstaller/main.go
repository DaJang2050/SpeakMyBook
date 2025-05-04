package main

import (
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
)

func main() {
	// 设置控制台编码为 UTF-8 (可选，如果需要正确显示中文)
	// 在 Windows 上，Go 程序默认使用 ANSI 编码，可能导致中文显示乱码。
	// 但直接在 Go 代码中设置 chcp 可能不可靠，建议在运行 Go 程序前设置好控制台编码。
	// 也可以尝试使用第三方库来处理控制台编码问题，例如 "golang.org/x/sys/windows"

	fmt.Println("执行：uv cache clean")
	err := executeCommand("uv", "cache", "clean")
	if err != nil {
		fmt.Printf("执行 'uv cache clean' 失败: %v\n", err)
	}

	uvPythonDir, err := executeCommandAndGetOutput("uv", "python", "dir")
	if err != nil {
		fmt.Printf("获取 uv python dir 失败: %v\n", err)
	} else {
		uvPythonDir = trimNewline(uvPythonDir)
		fmt.Printf("执行：rm -r \"%s\"\n", uvPythonDir)
		err = removeDirectory(uvPythonDir)
		if err != nil {
			fmt.Printf("删除 uv python dir 失败: %v\n", err)
		}
	}

	uvToolDir, err := executeCommandAndGetOutput("uv", "tool", "dir")
	if err != nil {
		fmt.Printf("获取 uv tool dir 失败: %v\n", err)
	} else {
		uvToolDir = trimNewline(uvToolDir)
		fmt.Printf("执行：rm -r \"%s\"\n", uvToolDir)
		err = removeDirectory(uvToolDir)
		if err != nil {
			fmt.Printf("删除 uv tool dir 失败: %v\n", err)
		}
	}

	homeDir, err := os.UserHomeDir()
	if err != nil {
		fmt.Printf("获取用户 Home 目录失败: %v\n", err)
	} else {
		uvExePath := filepath.Join(homeDir, ".local", "bin", "uv.exe")
		uvxExePath := filepath.Join(homeDir, ".local", "bin", "uvx.exe")

		fmt.Printf("执行：rm %s\n", uvExePath)
		err = removeFile(uvExePath)
		if err != nil {
			fmt.Printf("删除 uv.exe 失败: %v\n", err)
		}

		fmt.Printf("执行：rm %s\n", uvxExePath)
		err = removeFile(uvxExePath)
		if err != nil {
			fmt.Printf("删除 uvx.exe 失败: %v\n", err)
		}
	}

	fmt.Println("完成，请按回车键退出！")
	fmt.Scanln() // 等待用户按回车键
}

// executeCommand 执行命令并返回错误
func executeCommand(name string, arg ...string) error {
	cmd := exec.Command(name, arg...)
	cmd.Stdout = os.Stdout // 将命令的标准输出重定向到 Go 程序的标准输出
	cmd.Stderr = os.Stderr // 将命令的标准错误输出重定向到 Go 程序的标准错误输出
	return cmd.Run()
}

// executeCommandAndGetOutput 执行命令并返回输出结果
func executeCommandAndGetOutput(name string, arg ...string) (string, error) {
	cmd := exec.Command(name, arg...)
	output, err := cmd.CombinedOutput() // 获取命令的标准输出和标准错误输出
	return string(output), err
}

// removeDirectory 删除目录及其内容
func removeDirectory(dirPath string) error {
	err := os.RemoveAll(dirPath)
	return err
}

// removeFile 删除文件
func removeFile(filePath string) error {
	err := os.Remove(filePath)
	return err
}

// trimNewline 移除字符串末尾的换行符
func trimNewline(s string) string {
	for len(s) > 0 && (s[len(s)-1] == '\n' || s[len(s)-1] == '\r') {
		s = s[:len(s)-1]
	}
	return s
}
