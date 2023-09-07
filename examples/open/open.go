package main

import (
	"fmt"
	"github.com/devlights/gord"
	"os"
	"time"
)

type (
	filePath = string
)

func main() {
	if len(os.Args) < 2 {
		_, _ = fmt.Fprintf(os.Stderr, "need file-path")
		os.Exit(1)
	}

	if err := run(os.Args[len(os.Args)-1]); err != nil {
		panic(err)
	}
}

func run(p filePath) error {
	quitFn := gord.MustInitGord()
	defer quitFn()

	g, r := gord.MustNewGord()
	defer r()

	g.MustSetVisible(true)

	docs, err := g.Documents()
	if err != nil {
		return err
	}

	doc, docReleaseFn, err := docs.Open(p)
	if err != nil {
		return err
	}
	defer docReleaseFn()

	time.Sleep(5 * time.Second)
	if err := doc.Close(); err != nil {
		return err
	}

	return nil
}
