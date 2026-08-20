CALL "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\Tools\VsDevCmd.bat"
MSBuild HLTConsole.sln /p:Configuration=Debug /p:Platform=x86 /verbosity:minimal
