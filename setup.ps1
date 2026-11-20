# Office Viewer for BB — setup script
#
# Installs the Office Viewer plugin (Word / Excel / PowerPoint / PDF / images /
# video / audio / HTML) into BB from a public Git repository. Run this once on a
# new machine and the office stack is ready to use.
#
# Edit the URL below to point at your fork after you publish it to GitHub
# (keep the `git:` scheme and a `.git` suffix).
$ErrorActionPreference = "Stop"

# The Office Viewer plugin is the whole office stack in one plugin, so it lives
# at the root of its repository. Add more plugin sources here later if you build
# separate plugins.
$PluginSources = @(
  "git:https://github.com/gaen3erich/bb-office-viewer.git"
)

Write-Host "Office Viewer — BB plugin installer" -ForegroundColor Cyan
Write-Host "====================================" -ForegroundColor Cyan

foreach ($source in $PluginSources) {
  Write-Host "`n>>> bb plugin install $source --yes"
  bb plugin install $source --yes
  if ($LASTEXITCODE -ne 0) {
    throw "Install failed for $source (exit $LASTEXITCODE)"
  }
}

Write-Host "`nDone. Confirm with:  bb plugin list" -ForegroundColor Green
Write-Host "The Office Viewer action appears in the file opener for .docx .xlsx .pptx .pdf and more." -ForegroundColor Green
