# Office Json Kit

**Overview**

A simple utility used to export json text to office file (*.xlsx, *.docx).

**Usage**

    Usage:
    JsonKit.exe [--help] [--version] [--file=<file>] [--json=<a character sequence>] [-out=<file>] [--xlsx] [--docx] [<a character sequence>]
    Options:
    --help              display this help text and exit.
    --version           display version information and exit.
    --file=<file>       select a json file.
    --json=<a character sequence>
                        select a character sequence that represents a json object.
    --xlsx,--docx       select export type, default: --xlsx.
    --out=<file>        select output file path, default: out.xlsx or out.docx.
    <a character sequence>
                        select a character sequence that represents a json object, only if one positional parameter (not start with '--') provided, Jsonkit will use it as json text and export a *.xlsx.
    Examples:
       JsonKit.exe --file=sample.json --xlsx --out=out.xlsx
       JsonKit.exe --file=sample.json --xlsx
       JsonKit.exe --file=sample.json
       JsonKit.exe --json={"data":[{"symbol":"601288","name":"农业银行","value":"13.05"},{"symbol":"600291","name":"西水股份","value":"12.69"}]} --xlsx --out=out.xlsx
       JsonKit.exe --json={"data":[{"symbol":"601288","name":"农业银行","value":"13.05"},{"symbol":"600291","name":"西水股份","value":"12.69"}]} --xlsx
       JsonKit.exe --json={"data":[{"symbol":"601288","name":"农业银行","value":"13.05"},{"symbol":"600291","name":"西水股份","value":"12.69"}]}
       JsonKit.exe {"data":[{"symbol":"601288","name":"农业银行","value":"13.05"},{"symbol":"600291","name":"西水股份","value":"12.69"}]}

**Using the Windows command interpreter**

    > JsonKit.exe "{^"data^":[{^"symbol^":^"601288^",^"value^":^"13.05^"},{^"symbol^":^"600291^",^"value^":^"12.69^"}]}"
    > JsonKit.exe --json="{^"data^":[{^"symbol^":^"601288^",^"value^":^"13.05^"},{^"symbol^":^"600291^",^"value^":^"12.69^"}]}" --docx

**Quoting and escaping**

You can prevent [the special characters](https://en.wikibooks.org/wiki/Windows_Batch_Scripting#Quoting_and_escaping "Quoting_and_escaping") that control command syntax from having their special meanings as follows, except for the percent sign (%):
    You can surround a string containing a special character by quotation marks.
    You can place caret (^), an escape character, immediately before the special characters. In a command located after a pipe (|), you need to use three carets (^^^) for this to work.

**Notes**
The `sample.json` must be a json object contains a property named `data` with a json array as below:

    {
        "data": [
            {
                "symbol": "600015",
                "full_symbol": "sh600015",
                "name": "华夏银行",
                "value": "12.40"
            },
            {
                "symbol": "600000",
                "full_symbol": "sh600000",
                "name": "浦发银行",
                "value": "12.37"
            },
            {
                "symbol": "600030",
                "full_symbol": "sh600030",
                "name": "中信证券",
                "value": "11.97"
            }
        ]
    }

**Relations**

1. [NuGet](https://www.nuget.org/ "NuGet is the package manager for the Microsoft development platform including .NET.")
1. [JSON (JavaScript Object Notation)](http://json.org/ "JSON (JavaScript Object Notation) is a lightweight data-interchange format.")
1. [Json.NET](http://www.newtonsoft.com/json "Popular high-performance JSON framework for .NET")
1. [NPOI](http://npoi.codeplex.com/ "A .NET version of POI Java project at http://poi.apache.org/.")
1. [POI](http://poi.apache.org/ "Apache POI - the Java API for Microsoft Documents")

**FAQ**

1. [Does NPOI have support to .xlsx format?](http://stackoverflow.com/questions/16079956/does-npoi-have-support-to-xlsx-format)
1. [XWPF examples](http://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/xwpf/)
1. [LINQ to JSON, an API for working with JSON objects.](http://www.newtonsoft.com/json/help/html/LINQtoJSON.htm "LINQ to JSON is an API for working with JSON objects.")
