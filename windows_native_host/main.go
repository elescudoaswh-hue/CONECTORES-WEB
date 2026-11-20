// El navegador inicia este ejecutable y conserva los canales de mensajería al
// abrir el intérprete que viene en el mismo paquete.
package main

import (
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"syscall"
)

const createNoWindow = 0x08000000

func run() error {
	if len(os.Args) < 2 {
		return fmt.Errorf("falta el identificador de la extensión")
	}
	executable, err := os.Executable()
	if err != nil {
		return err
	}
	directory := filepath.Dir(executable)
	python := filepath.Join(directory, "runtime", "python.exe")
	script := filepath.Join(directory, "windows_native.py")
	child := exec.Command(python, script, os.Args[1])
	child.Dir = directory
	child.Stdin, child.Stdout, child.Stderr = os.Stdin, os.Stdout, os.Stderr
	child.SysProcAttr = &syscall.SysProcAttr{HideWindow: true, CreationFlags: createNoWindow}
	return child.Run()
}

func main() {
	if err := run(); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}
