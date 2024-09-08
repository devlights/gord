package main

import (
	"fmt"
	"os"
	"time"

	"github.com/devlights/gord"
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
	quit := gord.MustInitGord()
	defer quit()

	word, wordRelease := gord.MustNewGord()
	defer wordRelease()

	word.MustSetVisible(true)

	docs, err := word.Documents()
	if err != nil {
		return err
	}

	doc, docRelease, err := docs.Open(p)
	if err != nil {
		return err
	}
	defer docRelease()

	time.Sleep(5 * time.Second)
	if err := doc.Close(); err != nil {
		return err
	}

	return nil
}
