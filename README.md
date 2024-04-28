# Excel Go binding struct

[![build-img]][build-url]
[![pkg-img]][pkg-url]
[![coverage-img]][coverage-url]

## Installation

Go version 1.21+

```bash
go get github.com/itcomusic/excelstruct
```

## Usage

```go
package main

import (
    "context" 
	"fmt"
	
	"github.com/itcomusic/excelstruct"
)

func main() {
    conn, _ := amqpx.Connect()
    defer conn.Close()

    // simple publisher
    pub := amqpx.NewPublisher[[]byte](conn, amqpx.Direct, amqpx.UseRoutingKey("routing_key"))
    _ = pub.Publish(amqpx.NewPublishing([]byte("hello")).PersistentMode(), amqpx.SetRoutingKey("override_routing_key"))
	
    // simple consumer 
    _ = conn.NewConsumer("foo", amqpx.D(func(ctx context.Context, req *amqpx.Delivery[[]byte]) amqpx.Action {
        fmt.Printf("received message: %s\n", string(*req.Msg))
        return amqpx.Ack
    }))
}
```

## License

[MIT License](LICENSE)

[build-img]: https://github.com/itcomusic/excelstruct/workflows/build/badge.svg

[build-url]: https://github.com/itcomusic/excelstruct/actions

[pkg-img]: https://pkg.go.dev/badge/github.com/itcomusic/excelstruct.svg

[pkg-url]: https://pkg.go.dev/github.com/itcomusic/excelstruct

[coverage-img]: https://codecov.io/gh/itcomusic/excelstruct/branch/main/graph/badge.svg

[coverage-url]: https://codecov.io/gh/itcomusic/excelstruct