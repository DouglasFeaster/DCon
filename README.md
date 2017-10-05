# Document Converter

Console Application for converting Microsoft Office files to text for Git Diff. Add dcon directory to PATH variable.

Dependencies:

* Microsoft Windows 
* Microsoft Office
* .NET Framework 4.6.2
 

.gitconfig
```git
# .gitconfig file in your home directory
[diff "dcon"]
  textconv=dcon
  prompt = false
[alias]
  docdiff = diff --word-diff=color --unified=1
```

.gitattributes
```git
# .gitattributes file in root folder of your git project
    *.docx diff=dcon
    *.docm diff=dcon
    *.doc diff=dcon
    *.dotx diff=dcon
    *.dotm diff=dcon
    *.dot diff=dcon
    *.rtf diff=dcon
    *.odt diff=dcon

    *.xlsx diff=dcon
    *.xlsm diff=dcon
    *.xlsb diff=dcon
    *.xls diff=dcon
    *.csv diff=dcon
    *.xltx diff=dcon
    *.xltm diff=dcon
    *.xlt diff=dcon
    *.ods diff=dcon 

    *.pptx diff=dcon
    *.pptm diff=dcon
    *.ppt diff=dcon
    *.potx diff=dcon
    *.potm diff=dcon
    *.pot diff=dcon
    *.ppsx diff=dcon
    *.ppsm diff=dcon
    *.pps diff=dcon
    *.odp diff=dcon

    *.pdf diff=dcon
```
