Python's interactive mode is very convenient, because you can try simple stuff
without having to write it to a script first. Since I had to do a lot of
VBScript at the time, I wanted to have something like that for VBScript too.

I found [this blog][1] post that had almost exactly what I wanted, except for
line continuation. Which is what I added (for my own convenience). Plus some
other convenience features, like importing other VBScripts, or looking up help
topics in the VBScript help file.

Note that help lookups will only work with the *english* version of the Windows
Script Technologies help file ([`script56.chm`][2]) present on your system, as
other language versions of that file have different internal paths. The CHM file
must be located in either the Windows help directory, the `%PATH%`, or the
current working directory.

[1]: http://www.kryogenix.org/days/2004/04/01/interactivevbscript
[2]: http://download.microsoft.com/download/winscript56/Install/5.6/W982KMeXP/EN-US/scrdoc56en.exe
