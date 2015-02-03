package main 

import (
    "archive/zip"
    // "bytes"
    "strings"
    "encoding/json"
    "github.com/beevik/etree"
    "io/ioutil"
    // "io"
    "os"
    "fmt"
    "flag"
)

func main() {

    // Command-line arguments -f, -d, and -o
    fPtr := flag.String("f", "", "the template file")
    dPtr := flag.String("d", "", "the data string")
    oPtr := flag.String("o", "", "the output file")
    flag.Parse()
    
    // Go to func docxmerge()
    docxmerge(*fPtr, *dPtr, *oPtr)
}


func read_docx(fname string) string {
    out := ""
    // This is used to get "word/document.xml" from the docx file
    r, err := zip.OpenReader(fname)
    if err != nil {
        fmt.Println("We have an error opening the zip file!")
    }
    defer r.Close()
    // Iterate through the files in the archive,
    // printing some of their contents.
    for _, f := range r.File {
        if f.Name == "word/document.xml" {
            rc, _ := f.Open()
            out, _ := ioutil.ReadAll(rc)
            rc.Close()
            return string(out)  // Returns XML string of word docx  
        }
    }
    return string(out)
}

func replaceDocx(fname string, newfile string, newfilexml string) {
    out, _ := os.Create(newfile)
    w := zip.NewWriter(out)
    r, err := zip.OpenReader(fname)
    if err != nil {
        fmt.Println("We have an error opening the zip file!")
    }
    defer r.Close()
    for _, f := range r.File {
        curr, _ := w.Create(f.Name)
        if f.Name == "word/document.xml" {
            curr.Write([]byte(newfilexml)) // Here we replace the file
        } else {
            rc, _ := f.Open()
            rf, _ := ioutil.ReadAll(rc)
            curr.Write([]byte(rf))
        }
    }
    w.Close()
}

func checkElementIs(element string, match string) bool {
    if element == match {
        return true        
    }
    return false
}

func replaceHash(kp string, inputString string) string {
    var d map[string]string
    json.Unmarshal([]byte(kp), &d)
    for key, value := range d {
        if strings.Contains(inputString, key) {
            return value
        }
    }
    return inputString
}

func docxmerge(fname string, kp string, outname string) {

    doc := etree.NewDocument()
    str := read_docx(fname)
    doc.ReadFromString(str)
    root := doc.Root()
    for _, r := range root.FindElements("//") {
        if checkElementIs(r.Tag, "fldChar") {
            //we're looking for this attribute: w:fldCharType="separate"
            if r.Attr[0].Value == "separate" {
                // Get the text for the sibling node and replace it with the associated keypair value
                node_value := r.Parent.FindElements("..//")[6].ChildElements()[1].Text()
                r.Parent.FindElements("..//")[6].ChildElements()[1].SetText(replaceHash(kp, node_value))
            }
        } else if checkElementIs(r.Tag, "fldSimple") {
            node_value := r.ChildElements()[0].ChildElements()[1].Text()
            r.ChildElements()[0].ChildElements()[1].SetText(replaceHash(kp, node_value))
        }
    }
    out, _ := doc.WriteToString()
    // fmt.Println(out)
    replaceDocx(fname, outname, out)
}