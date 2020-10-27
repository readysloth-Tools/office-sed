# Office sed

Small utility based on pure-python (sed module) coreutils [library](https://github.com/readysloth/python-coreutils)
i wrote with my friend [@e1four15f](https://github.com/e1four15f).

`ofsed` can handle *only* substitution commands, because it is most frequent usecase and dealing with line breaks in
xml from office files can be kinda tricky.

If you want to `grep` things from office files, check out [office cat](https://github.com/readysloth-Tools/office-cat)

## How to launch

Download repository:

```
git clone --recursive https://github.com/readysloth-Tools/office-sed.git
```

Most likely you want to launch it like this:

```
python3 ofsed.py 's/mY\s[regular] (expression)/\1\1\1/Ig' file.odt document.docx ...
```

We didn't reinvent the wheel, so `s` command is based python regexps. You
can refer to capture groups in the same way, as you can do it in python.

**But I dare you from using `\0` capturing group for full match.** It can corrupt your document.

## Help

If you want better understanding of what is happening, read source code of `ofsed` and
[sed module](https://github.com/readysloth/python-coreutils/tree/master/coreutils/sed).

```
usage: ofsed.py [-h] [-i] [-r] CMD PATH [PATH ...]

Program to "sed" .docx and .odt files

positional arguments:
  CMD            command to execute
  PATH           path to document file

optional arguments:
  -h, --help     show this help message and exit
  -i, --inplace  change files inplace
  -r, --raw      use ofsed on raw xml files
```
