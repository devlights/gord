package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/devlights/gord/constants"

	"github.com/devlights/gord"
)

type (
	Args struct {
		dir     string
		text    string
		onlyHit bool
		verbose bool
		debug   bool
	}
)

var (
	args Args
)

var (
	appLog = log.New(os.Stdout, "", 0)
)

func init() {
	flag.StringVar(&args.dir, "dir", ".", "directory")
	flag.StringVar(&args.text, "text", "", "search text")
	flag.BoolVar(&args.onlyHit, "only-hit", true, "show ONLY HIT")
	flag.BoolVar(&args.verbose, "verbose", false, "verbose mode")
	flag.BoolVar(&args.debug, "debug", false, "debug mode")
}

func main() {
	flag.Parse()

	if args.text == "" {
		flag.PrintDefaults()
		os.Exit(1)
	}

	if args.dir == "" {
		args.dir = "."
	}

	if err := run(); err != nil {
		log.Fatal(err)
	}
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

	rootDir := abs(args.dir)
	err = filepath.WalkDir(rootDir, func(path string, d os.DirEntry, err error) error {
		if err != nil {
			return err
		}

		if d.IsDir() {
			return nil
		}

		if strings.Contains(filepath.Base(path), "~$") {
			return nil
		}

		if !strings.HasSuffix(path, ".docx") {
			return nil
		}

		absPath := abs(path)
		doc, docReleaseFn, err := docs.Open(absPath)
		if err != nil {
			return err
		}
		defer docReleaseFn()

		if args.debug {
			appLog.Printf("Document Open: %s", absPath)
		}

		allRange, err := doc.AllRange()
		if err != nil {
			return genErr("doc.AllRange()", err)
		}

		find, err := allRange.Find()
		if err != nil {
			return genErr("range.Find()", err)
		}

		if err := find.Forward(true); err != nil {
			return genErr("find.Forward()", err)
		}

		if err := find.Wrap(constants.WdFindWrapFindStop); err != nil {
			return genErr("find.Wrap()", err)
		}

		if err := find.MatchWildcards(true); err != nil {
			return genErr("find.MatchWildcards()", err)
		}

		if err := find.Text(args.text); err != nil {
			return genErr("find.Text()", err)
		}

		found, err := find.Execute()
		if err != nil {
			return genErr("find.Execute() [first time]", err)
		}

		relPath, _ := filepath.Rel(rootDir, absPath)
		if found {
			if args.verbose {
				foundRange := allRange

				for found {
					text, _ := foundRange.Text()
					page, _ := foundRange.PageNumber()
					line, _ := foundRange.LineNo()

					message := fmt.Sprintf("%s (%3d,%3d): %q", relPath, page, line, text)
					appLog.Println(message)

					found, err = find.Execute()
					if err != nil {
						return genErr("find.Execute() [after the second time]", err)
					}

					foundRange = allRange
				}
			} else {
				appLog.Printf("%s: HIT", relPath)
			}
		} else {
			if !args.onlyHit {
				appLog.Printf("%s: NO HIT", relPath)
			}
		}

		return nil
	})

	if err != nil {
		return err
	}

	return nil
}
