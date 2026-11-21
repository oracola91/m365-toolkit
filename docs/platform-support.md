# Requirements and platform support

[Back to the README](../README.md)

The tool itself is plain PowerShell and runs on **PowerShell 7+ (Windows, macOS, Linux)**
and **Windows PowerShell 5.1**. It only loads a service's module when you pick that
service, so unsupported services fail with a clear message rather than crashing the tool.

Coverage depends on the vendor module behind each service — and two of Microsoft's
modules are **Windows-only** (built on .NET Framework, which macOS/Linux don't have):

| Service | macOS / Linux (pwsh 7) | Windows (pwsh 7) | Windows PowerShell 5.1 |
|---|:--:|:--:|:--:|
| Exchange Online | ✅ | ✅ | ✅ |
| Security & Compliance (Purview) | ✅ | ✅ | ✅ |
| Entra / Identity (Graph) | ✅ | ✅ | ✅ |
| Intune / Devices (Graph) | ✅ | ✅ | ✅ |
| Microsoft Teams | ✅ | ✅ | ✅ |
| SharePoint Online (PnP) | ✅ | ✅ | ❌ (PnP needs pwsh 7.2+) |
| **SharePoint Admin (SPO)** | ❌ Windows only | ✅ | ✅ |
| **Power Platform (admin)** | ❌ Windows only | ✅ | ✅ |

> The two Windows-only services (**SharePoint Admin**, **Power Platform**) are a Microsoft
> platform limitation, not a bug in this tool — Microsoft doesn't ship those modules for
> macOS/Linux. On a Mac/Linux box the tool detects this and tells you to run them from Windows.
