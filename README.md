# spmt4all-host

Public GitHub Pages host for the **SPMT4all** Excel add-in (compiled Blazor WASM + `manifest.xml`).

Origin: `https://wolfgangmenabruhn.github.io/spmt4all-host/`

This repository is **not** the application source. Source lives in the private `SPMT4all-AppSource` repo. Update this host by running HostPublish from the source solution, then commit and push here.

## Update after an application change

From the source solution root:

```powershell
dotnet run --project SPMT4all.HostPublish/SPMT4all.HostPublish.csproj
```

Then in this repo:

```powershell
git add -A
git commit -m "Publish updated SPMT4all WASM host."
git push origin main
```

Wait for GitHub Pages to rebuild, then reload the Excel task pane (re-sideload `manifest.xml` from this repo if the version changed).

## Preserved files

HostPublish regenerates published assets but keeps `.git`, `.gitignore`, `.gitattributes`, `.nojekyll`, `404.html`, `support.html`, and this `README.md`.
