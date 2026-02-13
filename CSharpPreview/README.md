# TJ Task Center (C# Preview)

This is a WinForms preview of your dark dashboard in C# with:

- TJ branding (`assets/tj_logo.png`, `assets/tj_icon.ico`)
- Left task cards with `Run Task` buttons
- Right live console output panel
- Running/Ready status and `Stop Current Task` button

## Run

1. Install the .NET 8 SDK for Windows.
2. From the project root, run:

```powershell
dotnet run --project .\CSharpPreview\TJTaskCenterPreview.csproj
```

The app reads tasks from `tasks.json` and also auto-discovers `.py` files.
