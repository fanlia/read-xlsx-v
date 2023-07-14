
module read_xlsx_v

import time
import szip
import strconv

struct XLSX {
mut:
  zp szip.Zip
}

fn XLSX.new(zipfile string) !XLSX {
  mut zp := szip.open(zipfile, szip.CompressionLevel.no_compression, szip.OpenMode.read_only)!
  return XLSX{ zp: zp }
}

[unsafe]
fn (mut xlsx XLSX) free() {
  xlsx.zp.close()
}

fn (mut xlsx XLSX) open_xml(name string) !string {
  xlsx.zp.open_entry(name)!
  defer { xlsx.zp.close_entry() }
  size := xlsx.zp.size()
  buf := xlsx.zp.read_entry()!
  xml := unsafe { tos(buf, int(size)) }
  return xml
}

fn (mut xlsx XLSX) parse_shared_strings() ![]string {
  xml := xlsx.open_xml('xl/sharedStrings.xml')!
  tags := XMLParser.new(xml)
  return tags.filter(it.name == 'si').map(it.text())
}

fn (mut xlsx XLSX) parse_formats() ![]string {
  xml := xlsx.open_xml('xl/styles.xml')!
  tags := XMLParser.new(xml)
  formats := tags.filter(it.name == 'xf' && it.parent.name == 'cellXfs').map(it.attributes['numFmtId'])
  return formats
}

fn r2ci(r string) !int {
  letters := r.trim('0123456789').to_lower()
  mut sum := 0
  for letter in letters.split('') {
    num := strconv.parse_int(letter, 36, 0)!
    sum = sum * 26 + int(num) - 9
  }
  return sum - 1
}

fn (mut xlsx XLSX) parse_sheet(shared_strings []string, formats []string) ![][]string {
  xml := xlsx.open_xml('xl/worksheets/sheet1.xml')!
  tags := XMLParser.new(xml)
  rows := tags.filter(it.name == 'row')

  mut data := [][]string{cap: rows.len}

  for row in rows {
    cols := row.children
    mut values := []string{len: cols.len}
    for col_i, col in cols {
      ci := if col_r := col.attributes['r'] { r2ci(col_r)! } else { col_i }
      v := col.text() 
      mut value := v
      if t := col.attributes['t'] {
        match t {
          's' {
            value = shared_strings[strconv.parse_int(v, 10, 0)!] or { v }
          }
          'e' {
            value = ''
          }
          else { }
        }
      } else if s := col.attributes['s'] {
        format_id := strconv.parse_int(s, 10, 0)!
        if format := formats[format_id] {
          // println('format=${format} v=${v}')
          match format {
            '14' {
              t := excel_to_date(value)!
              value = t.custom_format('YYYY-MM-DD')
            }
            '15' {
              t := excel_to_date(value)!
              value = t.custom_format('YYYY-MM-DD')
            }
            '16' {
              t := excel_to_date(value)!
              value = t.custom_format('YYYY-MM-DD')
            }
            '17' {
              t := excel_to_date(value)!
              value = t.custom_format('YYYY-MM-DD')
            }
            '22' {
              t := excel_to_date(value)!
              value = t.custom_format('YYYY-MM-DD HH:mm:ss')
            }
            else { }
          }
        }
      }

      if ci >= values.len {
        unsafe {
          values.grow_len(ci - values.len + 1)
        }
      }
      // println('ci=${ci}, v=${v}, value=${value}')
      values[ci] = value
    }
    if values.any(it.trim_space() != '') {
      data << values
    }
  }

  return data 
}

fn excel_to_date(s string) !time.Time {
  v := strconv.atof64(s)!
  t := time.unix(i64((v - 25569) * 24 * 3600))
  return t
}

pub fn parse(xlsxfile string) ![][]string {
  mut xlsx := XLSX.new(xlsxfile)!

  shared_strings := xlsx.parse_shared_strings()!

  formats := xlsx.parse_formats()!

  data := xlsx.parse_sheet(shared_strings, formats)!

  return data
}

