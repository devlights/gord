package main

import (
	"flag"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/devlights/gord"
	"github.com/devlights/gord/constants"
)

type (
	Args struct {
		file       string
		rmOriginal bool
	}
)

var (
	args Args
)

func init() {
	flag.StringVar(&args.file, "file", "", "File to read from")
	flag.BoolVar(&args.rmOriginal, "rm", false, "Remove original file")
}

func main() {
	log.SetFlags(0)
	flag.Parse()

	if args.file == "" {
		flag.Usage()
		os.Exit(1)
	}

	if err := run(); err != nil {
		log.Fatal(err)
	}
}

func run() error {
	quitFn, _ := gord.InitGord()
	defer quitFn()

	g, r, _ := gord.NewGord()
	defer r()

	_ = g.Silent(false)

	docs, err := g.Documents()
	if err != nil {
		return err
	}

	absPath := abs(args.file)
	doc, docReleaseFn, err := docs.Open(absPath)
	if err != nil {
		return err
	}
	defer docReleaseFn()

	err = doc.SaveAsWithFileFormat(abs(strings.ReplaceAll(args.file, "doc", "docx")), constants.WdSaveFormatDocumentDefault)
	if err != nil {
		return err
	}

	if args.rmOriginal {
		if err := os.Remove(args.file); err != nil {
			return err
		}
	}

	return nil
}

func abs(p string) string {
	v, _ := filepath.Abs(p)
	return v
}
