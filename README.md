# gord
Gord is a library to operate MS Word using go-ole library. Thanks [go-ole](https://github.com/go-ole/go-ole) package!

This library works only on Windows.

## Install

```sh
go get github.com/devlights/gord@latest
```

## Usages

```go
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

// main is entry point of this app.
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
	// 0. Initialize Gord
	quitFn := gord.MustInitGord()
	defer quitFn()

	// 1. Create new Gord instance.
	g, release := gord.MustNewGord()

	// must call gord release function when function exited
	// otherwise WORD process was remained.
	defer release()

	// optional settings
	g.MustSetVisible(true)

	// 2. Get Documents instance.
	docs, err := g.Documents()
	if err != nil {
		return err
	}

	// 3. Open document
	doc, docReleaseFn, err := docs.Open(p)
	if err != nil {
		return err
	}

	// call document's release function
	defer docReleaseFn()

	// simulate something processing..
	time.Sleep(5 * time.Second)

	// 4. Close document
	if err := doc.Close(); err != nil {
		return err
	}

	// Document::SetSaved(true) and Document::Close() is automatically called when `defer docReleaseFn()`.
	// Word::Quit() and Word::Release() is automatically called when `defer release()`.

	return nil
}
```

Also look at the "examples" directory :)

## REFERENCES

- [Word VBA リファレンス](https://learn.microsoft.com/ja-jp/office/vba/api/overview/word)
- [列挙体 (Word)](https://learn.microsoft.com/ja-jp/office/vba/api/word(enumerations))
- [Word オブジェクト モデルの概要](https://learn.microsoft.com/ja-jp/visualstudio/vsto/word-object-model-overview?view=vs-2022&tabs=csharp)

## See also

- [Goxcel](https://github.com/devlights/goxcel)
  - Goxcel is a library to operate Excel using go-ole library.
