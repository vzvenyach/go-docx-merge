package main 

import (
    "archive/zip"
    "strings"
    "encoding/json"
    // "encoding/xml"
    "github.com/beevik/etree"
    "io/ioutil"
    // "os"
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
    fmt.Println(*fPtr, *dPtr, *oPtr)
    docxmerge(*fPtr, *dPtr)
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

func replace_docx() {
    //This is used to take the new file (the xml string) and replace the existing "word/document.xml" file with the new file and repackage the docx 

/*
    zin = zipfile.ZipFile(filepath, 'r')
    zout = zipfile.ZipFile(newfilepath, 'w')
    for item in zin.infolist():
        buffer = zin.read(item.filename)
        if (item.filename != 'word/document.xml'):
            zout.writestr(item, buffer)
        else:
            zout.writestr('word/document.xml', newfile)
    zin.close()
    zout.close()
    return True
*/

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

func docxmerge(fname string, kp string) {

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
    fmt.Println(doc.WriteToString())
}