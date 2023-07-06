
module read_xlsx_v

fn test_parse_xlsx() {
  actual := parse('data.xlsx')!
  want := [['hello', 'world'], ['bar', 'foo'], ['yes', 'no']]
  assert actual == want
}

fn test_r2ci() {
  actual := r2ci('A1')!
  want := 0
  assert actual == want
}
