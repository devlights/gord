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
	quitFn, _ := gord.InitGord()
	defer quitFn()

	g, r, _ := gord.NewGord()
	defer r()

	_ = g.Silent(false)

	docs, err := g.Documents()
	if err != nil {
		return genErr("g.Documents()", err)
	}

	newDoc, docReleaseFn, err := docs.Add()
	if err != nil {
		return genErr("docs.Add()", err)
	}
	defer docReleaseFn()

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
		nr, err := newDoc.Range()
		if err != nil {
			return genErr("newDoc.Range()", err)
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

	c, err := newDoc.Characters()
	if err != nil {
		return genErr("newDoc.Characters()", err)
	}

	cc, err := c.Count()
	if err != nil {
		return genErr("c.Count()", err)
	}

	if cc > 1 {
		// 最後の余分な改行を削除

		r, err := c.Last()
		if err != nil {
			return genErr("c.Last()", err)
		}

		ra, err := r.Previous()
		if err != nil {
			return genErr("r.Previous()", err)
		}

		err = ra.Delete()
		if err != nil {
			return genErr("ra.Delete()", err)
		}
	}

	// 保存
	err = newDoc.SaveAsWithFileFormatDefault(abs(args.out))
	if err != nil {
		return genErr("newDoc.SaveAs2()", err)
	}

	return nil
}
