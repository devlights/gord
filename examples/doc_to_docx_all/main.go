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
		dir        string
		rmOriginal bool
	}
)

var (
	args Args
)

func init() {
	flag.StringVar(&args.dir, "dir", ".", "directory")
	flag.BoolVar(&args.rmOriginal, "rm", false, "remove original file")
}

func main() {
	log.SetFlags(0)
	flag.Parse()

	if args.dir == "" {
		args.dir = "."
	}

	if err := run(); err != nil {
		log.Fatal(err)
	}
}

func run() error {
	quit, _ := gord.InitGord()
	defer quit()

	word, wordRelease, _ := gord.NewGord()
	defer wordRelease()

	_ = word.Silent(false)

	docs, err := word.Documents()
	if err != nil {
		return err
	}

	err = filepath.WalkDir(args.dir, func(path string, info os.DirEntry, err error) error {
		if err != nil {
			return err
		}

		if info.IsDir() {
			return nil
		}

		if !strings.HasSuffix(info.Name(), ".doc") {
			return nil
		}

		if err := toDocx(docs, path); err != nil {
			return err
		}

		if args.rmOriginal {
			return os.Remove(path)
		}

		return nil
	})

	return err
}

func toDocx(docs *gord.Documents, p string) error {
	doc, docRelease, err := docs.Open(abs(p))
	if err != nil {
		return err
	}
	defer docRelease()

	err = doc.SaveAsWithFileFormat(abs(strings.ReplaceAll(p, "doc", "docx")), constants.WdSaveFormatDocumentDefault)
	if err != nil {
		return err
	}

	return nil
}

func abs(p string) string {
	v, _ := filepath.Abs(p)
	return v
}
