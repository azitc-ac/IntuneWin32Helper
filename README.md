# IntuneWin32Helper
Quickly create and deploy Intune Win32 apps using PSADT. Supports WinGet and mutiple tenants.<br><br>
See https://blog.zarenko.net/intune-apps-verteilen-leicht-gemacht/<br><br>
<img width="726" height="443" alt="Screenshot 2025-12-09 12-42-03" src="https://github.com/user-attachments/assets/6537dcc9-3a4a-4c34-a831-f73432481e03" />


## Repository checks

Before committing, run the structural checks:

```powershell
.\Tests\Invoke-RepoChecks.ps1
```

They parse every `.ps1` and fail (exit code 1) on: syntax errors, a missing UTF-8 BOM,
unsuppressed `.Add()` return values (these leak `int` indices into a dialog result),
`Out-GridView`/`ogv` usage, parameters that the called function does not declare,
`break`/`continue` outside a loop of the same function, and a `deploy_template.ps1`
that no longer routes tenant selection through `Initialize-IntuneConnection`.

Each check corresponds to a bug this repository already had, so re-introducing one
turns the check red.
