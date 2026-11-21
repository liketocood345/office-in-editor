# Third-party notices

Office in Editor uses a mixed license; see `LICENSE`. Components from an upstream project keep their upstream license, and everything else is BSD-3-Clause. The project does not contain Microsoft Office or any editor application.

Upstream components in a release build:

| Component | Ships in | License | License text |
| --- | --- | --- | --- |
| .NET runtime and ASP.NET Core runtime | `payload/core` | MIT | `LICENSE` Part 2.1; notices in `payload/core/dotnet-THIRD-PARTY-NOTICES.TXT` and `payload/core/aspnetcore-THIRD-PARTY-NOTICES.TXT` |
| `@modelcontextprotocol/ext-apps` and the packages it bundles | `payload/core/mcp-app.html` | per package | `THIRD_PARTY_LICENSES.md` |

The build regenerates `THIRD_PARTY_LICENSES.md` from the MCP App bundle and copies the runtime notices from the restored runtime packs, so the notices always match what ships. Run `npm ci` from the lockfile; do not substitute unreviewed package versions. Build tools and .NET Framework reference assemblies are used only at build time and are not distributed.
