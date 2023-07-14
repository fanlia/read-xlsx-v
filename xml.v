
module read_xlsx_v

fn is_space(r rune) bool { return u8(r).is_space() }

struct Parser {
  source string
mut:
  index int
}

fn (mut p Parser) skip_whitespace() {
  for p.match_char(is_space, true) {}
}

fn (mut p Parser) parse_until(predict fn(rune) bool) string {
  start := p.index
  for p.match_char(predict, true) {}
  return p.source[start .. p.index]
}

fn (mut p Parser) is_char(ch rune, advance bool) bool {
  return p.match_char(fn [ch] (r rune) bool { return ch == r }, advance)
}

fn (mut p Parser) match_char(predict fn(rune) bool, advance bool) bool {
  if ch, index := p.peek_char() {
    if predict(ch) {
      if advance {
        p.index = index
      }
      return true
    }
    return false
  } else {
    return false
  }
}

fn (mut p Parser) next_char() ?rune {
  if ch, index := p.peek_char() {
    p.index = index
    return ch
  } else {
    return none
  }
}

fn (mut p Parser) peek_char() ?(rune, int) {
  if p.index >= p.source.len {
    return none
  }
  start := p.index
  char_len := utf8_char_len(unsafe { p.source.str[start] })
  end := if p.source.len - 1 >= p.index + char_len { p.index + char_len } else { p.source.len }
  r := unsafe { p.source[start .. end] }
  ch := r.utf32_code()
  return ch, end
}

struct XMLParser {
  Parser
pub mut:
  parent_indexes []int = [-1]
  level int
  tags []&Tag
}

fn XMLParser.new(source string) []&Tag {
  mut parser := XMLParser {
    source: source,
  }

  parser.parse()

  return parser.tags
}

fn (mut p XMLParser) parse() {
  for p.parse_tag() {}
  for tag in p.tags {
    if tag.parent_index > -1 {
      p.tags[tag.parent_index].children << tag
    }
  }
}

fn (mut p XMLParser) parse_tag() bool {
  p.skip_whitespace()
  start := p.index

  if ch := p.next_char() {
    match ch {
      `<` {
        for p.match_char(fn (r rune) bool { return r != `>` }, true) {}
        // consume >
        p.next_char()
        tag_source := p.source[start .. p.index]
        mut tag := Tag.new(tag_source)
        if tag.is_close {
          p.level -= 1
          p.parent_indexes.delete_last()
        } else if tag.is_self_close {
          tag.level = p.level
          tag.parent_index = p.parent_indexes.last()
          p.tags << tag
        } else {
          tag.level = p.level
          tag.parent_index = p.parent_indexes.last()
          p.tags << tag
          p.level += 1
          p.parent_indexes << p.tags.len - 1
        }
      }
      else {
        for p.match_char(fn (r rune) bool { return r != `<` }, true) {}
        text := p.source[start .. p.index].trim_space()
        tag := &Tag {
          name: '#text',
          text: text,
          parent_index: p.parent_indexes.last(),
          level: p.level,
        }
        p.tags << tag
      }
    }
    return true
  } else {
    return false
  }
}

struct Tag {
pub:
  name string
  attributes map[string]string
  is_close bool
  is_self_close bool
pub mut:
  parent_index int = -1
  level int
  text string
  children []&Tag
}

fn (tag &Tag) text() string {
  if tag.name == '#text' {
    return tag.text
  } else {
    return tag.children.map(it.text()).join('')
  }
}

fn Tag.new(source string) &Tag {
  mut parser := TagParser { source: source }
  return parser.parse()
}

struct TagParser {
  Parser
}

fn (mut p TagParser) parse_is_close() bool {
  return p.is_char(`/`, true)
}

fn (mut p TagParser) parse_name() string {
  name := p.parse_until(fn (r rune) bool { return !is_space(r) && r != `/` && r != `>` })
  return name
}

fn (mut p TagParser) parse_attributes() map[string]string {
  mut attributes := map[string]string{}

  for p.match_char(fn (r rune) bool { return r != `>` }, false) {
    p.skip_whitespace()

    if p.match_char(fn (r rune) bool { return r == `>` || r == `/` || r == `?` }, false) {
      break
    }

    key := p.parse_until(fn (r rune) bool { return r != `=` })

    // consume =
    p.next_char()

    p.skip_whitespace()

    // consume left "
    p.next_char()
    
    value := p.parse_until(fn (r rune) bool { return r != `"` })
    
    // consume right "
    p.next_char()

    attributes[key] = value
  }

  return attributes
}

fn (mut p TagParser) parse_is_self_close() bool {
  return p.is_char(`/`, true) || p.is_char(`?`, true)
}

fn (mut p TagParser) parse() &Tag {
  // consume <
  p.next_char()

  is_close := p.parse_is_close()

  name := p.parse_name()

  attributes := p.parse_attributes()

  is_self_close := p.parse_is_self_close()

  // consume >
  p.next_char()

  return &Tag{
    name: name,
    attributes: attributes,
    is_close: is_close,
    is_self_close: is_self_close,
  }
}
