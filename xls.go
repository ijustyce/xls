package xls

import (
	"bytes"
	"errors"
	"io"
	"os"

	"github.com/vstasn/ole2"
)

// ErrWorkbookNotFound is returned when neither "Workbook" nor "Book" stream
// could be found in the OLE2 directory structure.
var ErrWorkbookNotFound = errors.New("xls: no Workbook or Book stream found")

// Open opens an XLS file from the given file path.
// It returns a parsed WorkBook object, or an error if the file could not be opened
// or parsed successfully.
func Open(file string) (*WorkBook, error) {
	fi, err := os.Open(file)
	if err == nil {
		return OpenReader(fi)
	}

	return nil, err
}

// OpenWithCloser is similar to Open, but also returns the file handle (as io.Closer).
// This allows the caller to manually close the file when done.
// Useful when you want to avoid leaking file descriptors.
func OpenWithCloser(file string) (*WorkBook, io.Closer, error) {
	fi, err := os.Open(file)
	if err == nil {
		wb, err := OpenReader(fi)
		return wb, fi, err
	}

	return nil, nil, err
}

// OpenStream loads an XLS workbook from any io.Reader (e.g., network stream, compressed archive).
// Since the XLS format requires seeking, the entire input is buffered into memory.
// Not recommended for very large XLS files due to memory usage.
func OpenStream(r io.Reader) (*WorkBook, error) {
	buf := new(bytes.Buffer)
	if _, err := io.Copy(buf, r); err != nil {
		return nil, err
	}

	return OpenReader(bytes.NewReader(buf.Bytes()))
}

// OpenReader parses an XLS workbook from a seekable input stream (e.g., file, bytes.Reader).
// The reader must implement io.ReadSeeker as the underlying OLE2 format requires random access.
func OpenReader(reader io.ReadSeeker) (*WorkBook, error) {
	// Open the OLE2 compound document structure
	ole, err := ole2.Open(reader)
	if err != nil {
		return nil, err
	}

	// List all files (streams) within the OLE2 document
	dir, err := ole.ListDir()
	if err != nil {
		return nil, err
	}

	var book, root *ole2.File

	// Search for the relevant stream that contains workbook data.
	// The standard name is "Workbook", but some files use "Book" instead.
	for _, file := range dir {
		switch file.Name() {
		case "Workbook":
			if book == nil {
				book = file // Prefer "Workbook" if it's the first one found
			}
		case "Book":
			if book == nil {
				book = file // Fallback to "Book" only if "Workbook" wasn't seen first
			}
		case "Root Entry":
			root = file // Needed as context for resolving internal paths
		}
	}

	if book == nil {
		// Return explicit error if neither stream was found
		return nil, ErrWorkbookNotFound
	}

	// Construct the WorkBook from the selected stream
	return newWorkBookFromOle2(ole.OpenFile(book, root)), nil
}
