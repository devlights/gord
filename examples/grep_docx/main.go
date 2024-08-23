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
	flag.StringVar(&args.text, "text", "", "search text")
	flag.BoolVar(&args.onlyHit, "only-hit", true, "show ONLY HIT")
	flag.BoolVar(&args.verbose, "verbose", false, "verbose mode")
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

	err = filepath.WalkDir(abs(args.dir), func(path string, d os.DirEntry, err error) error {
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

		appLog.Printf("Document Open: %s", absPath)

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

		if found {
			if args.verbose {
				foundRange := allRange

				for found {
					start, _ := foundRange.Start()
					end, _ := foundRange.End()
					text, _ := foundRange.Text()

					message := fmt.Sprintf("\t>>> HIT (start=%d, end=%d, text=%q)", start, end, text)
					appLog.Println(message)

					found, err = find.Execute()
					if err != nil {
						return genErr("find.Execute() [after the second time]", err)
					}

					foundRange = allRange
				}
			} else {
				appLog.Println("\t>>> HIT")
			}
		} else {
			if !args.onlyHit {
				appLog.Println("\t>>> NO HIT")
			}
		}

		return nil
	})

	if err != nil {
		return err
	}

	return nil
}
