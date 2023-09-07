package main

import (
	"flag"
	"fmt"
	"github.com/devlights/gord"
	"github.com/devlights/gord/constants"
	"log"
	"os"
	"path/filepath"
)

// flag parameters
var (
	src string
	dst string
)

// logs
var (
	appLog = log.New(os.Stdout, ">>> ", 0)
)

func main() {
	var (
		returnCode int
	)

	if err := run(); err != nil {
		_, _ = fmt.Fprint(os.Stderr, err)
		returnCode = -1
	}

	appLog.Println("done")

	os.Exit(returnCode)
}

func run() error {
	abs := func(p string) string {
		v, _ := filepath.Abs(p)
		return v
	}

	flag.StringVar(&src, "src", "", "source file")
	flag.StringVar(&dst, "dst", "result.pdf", "output pdf name")
	flag.Parse()

	if src == "" {
		flag.Usage()
		return nil
	}

	quitFn, _ := gord.InitGord()
	defer quitFn()

	g, r, _ := gord.NewGord()
	defer r()

	_ = g.Silent(false)

	docs, err := g.Documents()
	if err != nil {
		return err
	}

	absPath := abs(src)
	doc, docReleaseFn, err := docs.Open(absPath)
	if err != nil {
		return err
	}
	defer docReleaseFn()
	appLog.Printf("Document Open: %s", absPath)

	absPath = abs(dst)
	err = doc.ExportAsFixedFormat(absPath, constants.WdExportFormatPDF)
	if err != nil {
		return err
	}
	appLog.Printf("Export   PDF : %s", absPath)

	return nil
}
