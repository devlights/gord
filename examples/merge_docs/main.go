package main

import (
	"flag"
	"fmt"
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
	flag.StringVar(&args.dir, "dir", ".", "directory")
	flag.StringVar(&args.out, "out", "result.docx", "output file path")
}

func main() {
	flag.Parse()

	if args.dir == "" {
		args.dir = "."
	}

	if args.out == "" {
		args.out = "result.docx"
	}

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
	quit, _ := gord.InitGord()
	defer quit()

	word, wordRelease, _ := gord.NewGord()
	defer wordRelease()

	_ = word.Silent(false)

	docs, err := word.Documents()
	if err != nil {
		return genErr("word.Documents()", err)
	}

	newDoc, docRelease, err := docs.Add()
	if err != nil {
		return genErr("docs.Add()", err)
	}
	defer docRelease()

	err = filepath.WalkDir(abs(args.dir), func(path string, d os.DirEntry, err error) error {
		if err != nil {
			return err
		}

		if d.IsDir() {
			return nil
		}

		if !strings.HasSuffix(path, ".docx") {
			return nil
		}

		// ドキュメントの最後に移動
		nr, err := newDoc.AllRange()
		if err != nil {
			return genErr("newDoc.AllRange()", err)
		}

		err = nr.Collapse(constants.WdCollapseDirectionEnd)
		if err != nil {
			return genErr("nr.Collapse()", err)
		}

		// ファイルを挿入
		absPath := abs(path)
		err = nr.InsertFile(absPath)
		if err != nil {
			return genErr("nr.InsertFile()", err)
		}

		appLog.Printf("MERGE : %s", absPath)

		return nil
	})

	if err != nil {
		return err
	}

	// 保存
	err = newDoc.SaveAsWithFileFormatDefault(abs(args.out))
	if err != nil {
		return genErr("newDoc.SaveAs2()", err)
	}

	return nil
}
