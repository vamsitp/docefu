## epicdoc
> Generate Word-doc (or) Powerpoint-deck from Azure DevOps (VSTS) Epics / Features / User-Stories (PBIs)

---

### usage
```cmd
# word
epicdoc /w 

# ppt
epicdoc /p

# offline (genrate both word and ppt for previously fetched data)
epicdoc /o

```

---

```cmd
# Publish package to nuget.org
nuget push ./bin/epicdoc.1.0.0.nupkg -ApiKey <key> -Source https://api.nuget.org/v3/index.json

# Install from nuget.org
dotnet tool install -g epicdoc
dotnet tool install -g epicdoc --version 1.0.x

# Install from local project path
dotnet tool install -g --add-source ./bin epicdoc

# Uninstall
dotnet tool uninstall -g epicdoc
```
> **NOTE**: If the Tool is not accessible post installation, add `%USERPROFILE%\.dotnet\tools` to the PATH env-var.
