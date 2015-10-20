#Wordlist Cleaner
<p align="center"><img src="https://raw.githubusercontent.com/wpatoolkit/Wordlist-Cleaner/master/screenshot.jpg" /></p>
This is a simple GUI tool for Windows users to help remove unwanted characters from wordlists.

You can either clean a single wordlist in <b>File Mode</b> or multiple wordlists in <b>Directory Mode</b>.

In <b>File Mode</b> you specify a single wordlist you would like to clean and a single output file you would like to save it to.

In <b>Directory Mode</b> you specify a directory and it will read through every file in that directory cleaning each one and outputting a new cleaned version to your output directory.

<b>Input File</b><br>
This is where you specify your input wordlist (or input directory for Directory Mode).

<b>Output File</b><br>
This is where you specify your output wordlist (or output directory for Directory Mode).

<b>Only allow these characters</b><br>
This option allows you to specify which characters you want to keep from your wordlist. By default only the <a href="https://en.wikipedia.org/wiki/ASCII#ASCII_printable_characters">95 printable ASCII characters</a> are kept.

<b>When a non-allowed character is found</b><br>
You can specify whether to throw away the character or throw away the line when a non-allowed character is found.

<b>Remove numerics</b><br>
This option allows you to remove all numerics (or just numerics of a certain length) from your wordlist.<br>

<b>Replace accented characters</b><br>
If any characters are found with accents (such as á) you have the option to replace them with their non-accented versions (such as a).

<b>Minimum word length</b><br>
This option allows you to specify a minimum word length. Any word smaller than this will be discarded.
If you select the option "Repeat word until it reaches min length" it will keep repeating short words until they reach your min length. For example if you min length was 8 and your word was "abc" it would be changed to "abcabcabc".

<b>Maximum word length</b><br>
This option allows you to specify a maximum word length. Any word larger than this will be discarded.

<b>Convert all characters to lowercase</b><br>
This option allows you to convert all words in your wordlist to lowercase.

<b>Convert all characters to uppercase</b><br>
This option allows you to convert all words in your wordlist to uppercase.

<b>Trim leading and trailing whitespace</b><br>
This option allows you to trim all leading and trailing whitespace from every word in your wordlist. This includes spaces and tabs.

<b>Convert all newlines to Unix format</b><br>
This option allows you to convert all newlines to the Unix (linefeed only) format.

This program was written in Visual Basic 6 which means it should work on any modern version of Windows but just in case you need to download the VB6 runtimes you can do so from here:

<a href="https://www.microsoft.com/en-us/download/details.aspx?id=24417">https://www.microsoft.com/en-us/download/details.aspx?id=24417</a>
