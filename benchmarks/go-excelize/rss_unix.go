//go:build darwin || linux

package main

import (
	"runtime"
	"syscall"
)

func maxRSSBytes() *uint64 {
	var usage syscall.Rusage
	if err := syscall.Getrusage(syscall.RUSAGE_SELF, &usage); err != nil || usage.Maxrss < 0 {
		return nil
	}

	maximum := uint64(usage.Maxrss)
	if runtime.GOOS != "darwin" {
		maximum *= 1_024
	}
	return &maximum
}
