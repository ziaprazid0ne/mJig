<#
.SYNOPSIS
    Rewrites commits from Cursor Agent, Gmail, or Work identities to your anonymous GitHub identity.
#>

Param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$RepoUrl
)

# --- Configuration & Defaults ---
if ($env:TARGET_NAME) { $TargetName = $env:TARGET_NAME } else { $TargetName = git config --global --get user.name }
if ($env:TARGET_EMAIL) { $TargetEmail = $env:TARGET_EMAIL } else { $TargetEmail = git config --global --get user.email }

# Standard Cursor Matchers
$CursorEmails = if ($env:CURSOR_EMAILS) { $env:CURSOR_EMAILS } else { "cursoragent@users.noreply.github.com,cursoragent@cursor.com" }
$CursorNames = if ($env:CURSOR_NAMES) { $env:CURSOR_NAMES } else { "cursoragent,Cursor Agent,Cursoragent" }
$CursorSubstrings = if ($env:CURSOR_SUBSTRINGS) { $env:CURSOR_SUBSTRINGS } else { "cursoragent,cursor agent,@cursor.com" }

# --- Added Extra Identities to Scrub ---
$ExtraEmails = "grubbsbr22@gmail.com,bgrubbs@tdstickets.com"

$WorkDir = if ($env:WORKDIR) { $env:WORKDIR } else { Join-Path $env:TEMP "repo-clean-$([Guid]::NewGuid().Guid).git" }

# --- Validation ---
if (-not $TargetName -or -not $TargetEmail) {
    Write-Host "ERROR: Missing TARGET_NAME or TARGET_EMAIL." -ForegroundColor Red
    exit 1
}

try { git filter-repo -h | Out-Null } catch {
    Write-Host "ERROR: git-filter-repo is not installed." -ForegroundColor Red
    exit 1
}

# --- Execution ---
if (Test-Path $WorkDir) { Remove-Item -Recurse -Force $WorkDir }

Write-Host "Cloning mirror: $RepoUrl" -ForegroundColor Cyan
git clone --mirror $RepoUrl $WorkDir
Set-Location $WorkDir

# Updated Python Callback with the extra email checks
$PythonCallback = @"
target_name = b"$TargetName"
target_email = b"$TargetEmail"

# Matchers
cursor_emails = {e.strip().lower().encode("utf-8") for e in "$CursorEmails".split(",") if e.strip()}
cursor_names = {n.strip().lower().encode("utf-8") for n in "$CursorNames".split(",") if n.strip()}
cursor_substrings = [s.strip().lower().encode("utf-8") for s in "$CursorSubstrings".split(",") if s.strip()]
extra_emails = {e.strip().lower().encode("utf-8") for e in "$ExtraEmails".split(",") if e.strip()}

def _should_rewrite(name, email):
    nl = name.lower()
    el = email.lower()
    # Check if it's a Cursor identity
    if el in cursor_emails or nl in cursor_names:
        return True
    # Check if it matches your specific Gmail or Work emails
    if el in extra_emails:
        return True
    # Check for substrings (cursor agent)
    return any(sub in nl or sub in el for sub in cursor_substrings)

if _should_rewrite(commit.author_name, commit.author_email):
    commit.author_name = target_name
    commit.author_email = target_email
if _should_rewrite(commit.committer_name, commit.committer_email):
    commit.committer_name = target_name
    commit.committer_email = target_email
"@

Write-Host "Rewriting history with git filter-repo..." -ForegroundColor Yellow
git filter-repo --force --commit-callback $PythonCallback

Write-Host "Restoring remote origin..."
git remote add origin $RepoUrl

Write-Host "Pushing rewritten history to origin..." -ForegroundColor Red
git push --force --mirror origin

Write-Host "Process Complete. All specified identities have been scrubbed." -ForegroundColor Green