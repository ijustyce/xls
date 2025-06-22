# xls

[![GoDoc](https://godoc.org/github.com/MeKo-Christian/xls?status.svg)](https://godoc.org/github.com/MeKo-Christian/xls)

A pure Go library for reading Microsoft Excel `.xls` (BIFF8) files.

This is a maintained and modernized fork of the original library by [Rongshu Tech (Chinese)](http://www.rongshu.tech), based on [libxls](https://github.com/libxls/libxls).  
Special thanks to original contributors including [@tgulacsi](https://github.com/tgulacsi) and [@flyin9](https://github.com/flyin9).

---

## ✨ What's New in This Fork

- ✅ Added `OpenStream` function to read `.xls` files from any `io.Reader`
- ✅ Refactored core internals for clarity and maintainability
- ✅ Added detailed inline documentation and comments
- ✅ Improved test coverage and modernized test suite

---

## 📦 Basic Usage

```go
import "github.com/MeKo-Christian/xls"

// Open from file (auto-closes on error)
wb, err := xls.Open("example.xls")
if err != nil {
	log.Fatal(err)
}

// Open from file with manual control over closing
wb, closer, err := xls.OpenWithCloser("example.xls")
if err != nil {
	log.Fatal(err)
}
defer closer.Close()

// Open from an io.Reader (fully buffered into memory)
f, _ := os.Open("example.xls")
defer f.Close()
wb, err := xls.OpenStream(f)
```

See [GoDoc](https://godoc.org/github.com/MeKo-Christian/xls) for full API documentation and examples.

---

## 📁 Features

- Reads `.xls` (BIFF8) files
- Supports cell values, formats, dates, and SST (shared string table)
- Minimal dependencies
- Zero C bindings – pure Go implementation

---

## 🛠 Limitations

- Write support (`.xls` export) is **not** available
- Formula evaluation is not yet implemented
