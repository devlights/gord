package main

import (
	"log"
	"path/filepath"

	"github.com/devlights/gord"
	"github.com/devlights/gord/constants"
)

func main() {
	log.SetFlags(0)

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

	docs, _ := word.Documents()
	doc, docRelease, _ := docs.Add()
	defer docRelease()

	for _, s := range []string{"hello", "world", "こんにちは", "世界"} {
		paras, err := doc.Paragraphs()
		if err != nil {
			return err
		}

		para, err := paras.Add(nil)
		if err != nil {
			return err
		}

		paraRange, err := para.Range()
		if err != nil {
			return err
		}

		err = paraRange.SetText(s + "\n")
		if err != nil {
			return err
		}
	}

	file := "result.docx"
	abs, _ := filepath.Abs(file)

	if err := doc.SaveAs2(abs, constants.WdSaveFormatDocumentDefault); err != nil {
		return err
	}

	return nil
}
