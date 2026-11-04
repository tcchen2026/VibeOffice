// Office 2016 test VM HTTP API. Build: GOOS=windows GOARCH=amd64 go build -o win_test_server.exe .
package main

import (
	"flag"
	"log"
	"net/http"
	"time"

	"vibeoffice/win_test_server/api"
)

func main() {
	listen := flag.String("listen", "127.0.0.1:8765", "HTTP address (use an SSH tunnel for remote access)")
	root := flag.String("dir", `C:\vo`, "working directory")
	flag.Parse()
	handler, err := api.New(*root, api.RunOffice)
	if err != nil {
		log.Fatal(err)
	}
	server := &http.Server{Addr: *listen, Handler: handler, ReadHeaderTimeout: 10 * time.Second, ReadTimeout: 2 * time.Minute, IdleTimeout: time.Minute, MaxHeaderBytes: 16384}
	log.Printf("Office test API listening on %s; work directory %s", *listen, *root)
	log.Fatal(server.ListenAndServe())
}
