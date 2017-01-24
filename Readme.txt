- Right click on My Computer / This PC (depending on OS).
- Properties | Advanced System Settings | Environment Variables.
- Add new System Variable
- PROVIDEX={Path to your ProvideX interpreter}, eg. PROVIDEX=c:\pvxsrc
- Click OK, and OK dismiss all dialogs.
- Run cmd.exe and run >set
- Verify the environment setting is there for PROVIDEX
- Close down VS is running and restart.
- Project should now be pointing to your ProvideX folder for the Output path.
- For debugging, you should use the following for the Sage.Office365.Graph.csproj.user (user based setting file)

<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="14.0" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <PropertyGroup Condition="'$(Configuration)|$(Platform)' == 'Debug|AnyCPU'">
    <StartAction>Program</StartAction>
    <StartProgram>$(PROVIDEX)\pvxwin32.exe</StartProgram>
    <StartWorkingDirectory>$(PROVIDEX)</StartWorkingDirectory>
  </PropertyGroup>
</Project>

- This will run ProvideX when Start is issued in VS. The graph.pvx (text based) program will also be copied to the ProvideX folder
on successful build. This provides example usage for the assembly API.

