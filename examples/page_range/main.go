package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/devlights/gord"
)

type (
	Args struct {
		file string
		page int
	}
)

var (
	args Args
)

var (
	appLog = log.New(os.Stdout, ">>> ", 0)
)

func init() {
	flag.StringVar(&args.file, "file", "", "file path")
	flag.IntVar(&args.page, "page", 1, "page number")
}

func main() {
	flag.Parse()

	if args.file == "" || args.page <= 0 {
		flag.PrintDefaults()
		os.Exit(1)
	}

	args.file = abs(args.file)

	if err := run(); err != nil {
		log.Fatal(err)
	}

	appLog.Println("done")
}

func abs(p string) string {
	v, err := filepath.Abs(p)
	if err != nil {
		panic(err)
	}

	return v
}

func genErr(procName string, err error) error {
	return fmt.Errorf("%s failed: %w", procName, err)
}

func run() error {
	quit, _ := gord.InitGord()
	defer quit()

	word, wordRelease, _ := gord.NewGord()
	defer wordRelease()

	_ = word.Silent(false)

	docs, err := word.Documents()
	if err != nil {
		return genErr("word.Documents()", err)
	}

	doc, docRelease, err := docs.Open(args.file)
	if err != nil {
		return genErr("docs.Open(args.file)", err)
	}
	defer docRelease()

	pageRange, err := doc.GetPageRange(int32(args.page))
	if err != nil {
		return genErr("doc.GetPageRange(args.page)", err)
	}

	start, _ := pageRange.Start()
	end, _ := pageRange.End()
	text, _ := pageRange.Text()
	page, _ := pageRange.PageNumber()

	appLog.Printf("page=%d, start=%d, end=%d, text=%q)", page, start, end, text)

	return nil
}
