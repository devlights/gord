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
		dir string
	}
)

var (
	args Args
)

var (
	appLog = log.New(os.Stdout, ">>> ", 0)
)

func init() {
	flag.StringVar(&args.dir, "dir", ".", "directory")
}

func main() {
	flag.Parse()

	if args.dir == "" {
		args.dir = "."
	}

	if err := run(); err != nil {
		log.Fatal(err)
	}

	appLog.Println("done")
}

func run() error {
	abs := func(p string) string {
		v, _ := filepath.Abs(p)
		return v
	}

	quit, _ := gord.InitGord()
	defer quit()

	word, wordRelease, _ := gord.NewGord()
	defer wordRelease()

	_ = word.Silent(false)

	docs, err := word.Documents()
	if err != nil {
		return err
	}

	err = filepath.WalkDir(args.dir, func(path string, d os.DirEntry, err error) error {
		if err != nil {
			return err
		}

		if d.IsDir() {
			return nil
		}

		if !strings.HasSuffix(path, ".docx") {
			return nil
		}

		absPath := abs(path)
		doc, docRelease, err := docs.Open(absPath)
		if err != nil {
			return err
		}
		defer docRelease()
		appLog.Printf("Document Open: %s", absPath)

		absPath = abs(strings.ReplaceAll(path, ".docx", ".pdf"))
		err = doc.ExportAsFixedFormat(absPath, constants.WdExportFormatPDF)
		if err != nil {
			return err
		}
		appLog.Printf("Export   PDF : %s", absPath)

		return nil
	})

	if err != nil {
		return err
	}

	return nil
}
