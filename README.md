xlu-to-win
=============

Excel Unicode To Windows

When you save files in excel as 'unicode', it writes out a tab-delimited text file in utf-16le with BOM.

More often than not, I'm parsing such files and sending the resulting data to a sql2008 table with
windows-1252 as the standard table encoding.  

This library takes a filename (expected to be an excel 'unicode' file) and transcodes to windows-1252,
so that I don't keep having to memorize and/or lookup this line of code:

```ruby
file=File.open("name of unicode file.txt", "r:bom|utf-16le:windows-1252")
```

Hurray!

Usage
-----------

```ruby
require 'xlutowin'
```

open an excel unicode file for reading

```ruby
file=XLUTOWIN.open("filename") # file is now open for reading, and will transcode to windows-1252
```

same as last example w/ block

```ruby
XLUTOWIN.open("filename"){|file| ... }
```

get all lines (as hashes) from excel unicode file, transcoded to win1252

```ruby
all_lines=XLUTOWIN.each_row(filename)
puts all_lines[0][:col_header_one]
puts all_lines[0][:col_header_two]
...
```

same as last example, w/ block

```ruby
XLUTOWIN.each_row(filename) do |row_as_hash|
       puts row_as_hash[:col_header_one]
       puts row_as_hash[:col_header_two]
       ...
end
```
