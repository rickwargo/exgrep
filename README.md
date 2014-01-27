exgrep
======

**Grep** for **Excel** files

This command line utility will perform a *regular expression* search on many
Excel objects. It is only meant to be run from *Windows* as it manipulates the
Excel COM object through **Win32ole automation**. 

It decomposes all of the significant objects in Excel into text streams and
applies regex matching to find items of interest. It is also capable of 
find and replace.

```
Usage: exgrep [options] [expression] file ...

Options:
    -l, --files-with-matches         only print file names containing matches
    -L, --files-without-matches      only print file names containing no match
    -i, --ignore-case                ignore case distinctions
    -X, --extended                   use extended regular expressions
    -M, --multi-line                 search across lines
    -v, --invert-match               select non-matching lines
    -n, --line-numbers               print line number with output lines
    -R, --recurse                    recurse into directories
    -e, --regexp PATTERN             use PATTERN as a regular expression (may have multiples)
    -r, --replace STRING             use STRING as a replacement to the regular expression
    -D, --delete-matching-line       delete lines matching the regular expression
    -s, --search WHAT                workbook objects to search
                                       (all, formulas, macros, procedures, data, values, names, addins, references, conditionalformatting, properties, controls, comments, workbookcomments)
    -c, --controls NAME              search only controls matching NAME (should include property)
    -p, --procedure NAME             search only the NAMEd procedure (may have multiples)
    -m, --max-count NUM              stop after NUM matches
        --include PATTERN            files that match PATTERN will be examined
        --exclude PATTERN            files that match PATTERN will be skipped
    -S, --sheets-matching PATTERN    sheets that match PATTERN will be examined

Options that shouldn't be options:
    -C, --recycle-every NUM          recycle excel application every NUM times

Common options:
    -h, --help                       show this message
    -V, --verbose                    show messages indicating progress
        --version                    show version
```
