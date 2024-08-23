package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/devlights/gord/constants"

	"github.com/devlights/gord"
)

type (
	Args struct {
		in  string
		out string
	}
)

var (
	args Args
)

var (
	appLog = log.New(os.Stdout, ">>> ", 0)
)

func init() {
	flag.StringVar(&args.in, "in", "", "input")
	flag.StringVar(&args.out, "out", "", "output")
}

func main() {
	flag.Parse()

	if args.in == "" || args.out == "" {
		flag.PrintDefaults()
		os.Exit(1)
	}

	args.in = abs(args.in)
	args.out = abs(args.out)

	if err := run(); err != nil {
		log.Fatal(err)
	}

	appLog.Println("done")
}

func abs(p string) string {
	v, _ := filepath.Abs(p)
	return v
}

func genErr(procName string, err error) error {
	return fmt.Errorf("%s failed: %w", procName, err)
}

func run() error {
	quitFn, _ := gord.InitGord()
	defer quitFn()

	g, r, _ := gord.NewGord()
	defer r()

	_ = g.Silent(false)

	docs, err := g.Documents()
	if err != nil {
		return genErr("g.Documents()", err)
	}

	doc, docReleaseFn, err := docs.Open(args.in)
	if err != nil {
		return genErr("docs.Open(args.in)", err)
	}
	defer docReleaseFn()

	dr, err := doc.AllRange()
	if err != nil {
		return genErr("doc.AllRange()", err)
	}

	if err := dr.CopyAsPicture(); err != nil {
		return genErr("dr.CopyAsPicture()", err)
	}

	if err := dr.Collapse(constants.WdCollapseDirectionEnd); err != nil {
		return genErr("dr.Collapse(constants.WdCollapseDirectionEnd)", err)
	}

	if err := dr.PasteSpecial(constants.WdPasteDataTypePasteMetafilePicture); err != nil {
		return genErr("dr.PasteSpecial(constants.WdPasteDataTypePasteBitmap)", err)
	}

	if err := doc.Save(); err != nil {
		return genErr("doc.Save()", err)
	}

	return nil
}
