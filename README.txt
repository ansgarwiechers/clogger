What is this?
=============
The Logger class serves as an abstraction layer for logging in VBScripts. That
way a script can define the logging facilities once and then use the same set
of functions for writing log messages to all selected facilities. Supported
facilities are:

- interactive console/desktop
- log file
- eventlog

By default, console logging is enabled when the script is run interactively
(WScript.Interactive = True), otherwise the default is logging to eventlog.
When console logging is enabled and the script is run with the cscript.exe
interpreter, log messages are written to StdOut or StdErr, depending on their
log level. When the script is run with the wscript.exe interpreter, console
log messages are displayed as MsgBox() pop-ups instead.

The class does not do any error handling by itself. All errors MUST be handled
by the parent script.


Copyright
=========
It is distributed according to the terms of the GNU General Public License
Version 2.0 as found at <http://www.gnu.org/licenses/old-licenses/gpl-2.0.html>.

This program is distributed in the hope that it will be useful, but WITHOUT ANY
WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A
PARTICULAR PURPOSE.  See the GNU General Public License for more details.


Including the class
===================
To use the class in your scripts you must either copy/paste it to the script,
or use this neat import procedure:

' <http://gazeek.com/coding/importing-vbs-files-in-your-vbscript-project/>
Sub Import(ByVal filename)
	Dim fso, sh, file, code

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set sh = CreateObject("WScript.Shell")
	filename = sh.ExpandEnvironmentStrings(filename)
	filename = fso.GetAbsolutePathName(filename)
	Set file = fso.OpenTextFile(filename)
	code = file.ReadAll
	file.Close
	ExecuteGlobal(code)
End Sub


Example
=======
' creating a new Logger instance
Set myLogger = New Logger

' configuring the logger
myLogger.Debug = True               ' enable debug logging
myLogger.LogToConsole = False       ' disable logging to console
myLogger.LogToEventlog = True       ' enable logging to eventlog
myLogger.Overwrite = True           ' overwrite log file (default is to append)
                                    ' This property must be set before the log
                                    ' file is opened. It's ignored by all other
                                    ' facilities.
myLogger.LogFile = "C:\my.log"      ' enable logging to file C:\my.log

' logging stuff
myLogger.LogInfo "foo"              ' log information level message
myLogger.Log "foo2"                 ' alias for LogInfo()
myLogger.LogError "bar"             ' log error message
myLogger.LogDebug "baz"             ' log debug message (if Debug is set to
                                    ' True, otherwise these messages are
                                    ' discarded)


Thanks to
=========
Alexander Bernauer <http://blog.copton.net/>
Rico Schiekel <http://downgra.de/>

