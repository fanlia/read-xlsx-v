# read_xlsx_v
read xlsx using vlang

## Usage

```v
import read_xlsx_v

fn main() {
    data := read_xlsx_v.parse('read_xlsx_v/data.xlsx')!
    println(data)
}

// [['hello', 'world'], ['bar', 'foo'], ['yes', 'no']]

```

## Test

```sh
v test .
```

## License

MIT
