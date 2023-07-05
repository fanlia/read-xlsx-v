
module read_xlsx_v

fn test_parse_xlsx() {
  actual := parse('data.xlsx')!
  want := [['hello', 'world'], ['bar', 'foo'], ['yes', 'no']]
  assert actual == want
}
