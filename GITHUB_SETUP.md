# GitHub Repository Setup Guide

Your local git repository is ready! Now follow these steps to push to GitHub.

## 🚀 Step 1: Create a GitHub Repository

### Option A: Using GitHub Web Interface (Easiest)

1. **Go to GitHub:** https://github.com/new
2. **Fill in Repository Details:**
   - Repository name: `ppt-excel-sync`
   - Description: "Automatically sync PowerPoint one-pagers with Excel data"
   - Visibility: **Public** (recommended for open source) or **Private** (if internal only)
   - Do NOT initialize with README, .gitignore, or LICENSE (we already have these!)

3. **Click "Create repository"**

4. **Copy the repository URL** (you'll need it next)

---

## 🔗 Step 2: Connect Local Repository to GitHub

**In PowerShell, run these commands:**

```powershell
cd "C:\Users\GESZA4\OneDrive - IKEA\Apps\ppt to excel"

# Replace with YOUR GitHub username and repository URL
git remote add origin https://github.com/YOUR-USERNAME/ppt-excel-sync.git

# Verify the connection
git remote -v
```

**Expected output:**
```
origin  https://github.com/YOUR-USERNAME/ppt-excel-sync.git (fetch)
origin  https://github.com/YOUR-USERNAME/ppt-excel-sync.git (push)
```

---

## 📤 Step 3: Push to GitHub

**First push (use --set-upstream):**

```powershell
git branch -M main
git push -u origin main
```

**You'll be prompted to authenticate:**
- Use your GitHub username
- Use a **Personal Access Token** (not your password)
  - Go to: https://github.com/settings/tokens
  - Click "Generate new token"
  - Select scopes: `repo`, `workflow`
  - Copy the token and paste when prompted

**After first push, future pushes are simpler:**
```powershell
git push
```

---

## ✅ Verify on GitHub

1. Go to https://github.com/YOUR-USERNAME/ppt-excel-sync
2. You should see all your files
3. Check that README.md displays correctly

---

## 🎁 Optional Enhancements

### Add GitHub Topics
1. Go to your repository settings
2. Add topics: `python`, `excel`, `powerpoint`, `automation`, `data-sync`

### Create GitHub Pages Documentation
1. Enable Pages in Settings → Pages
2. Set source to `main` branch, `docs` folder
3. Your docs will be at: https://YOUR-USERNAME.github.io/ppt-excel-sync

### Add a Badge to README
```markdown
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python 3.8+](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/)
```

---

## 🔄 Working with the Repository

### After Making Changes

```powershell
# Check what changed
git status

# Add changes
git add .

# Commit with clear message
git commit -m "[FEATURE] Add data validation layer"

# Push to GitHub
git push
```

### Create a Feature Branch

```powershell
# Create and switch to new branch
git checkout -b feature/add-data-validation

# Make changes and commit
git add .
git commit -m "[FEATURE] Add data validation for percentages"

# Push branch to GitHub
git push -u origin feature/add-data-validation

# Then create a Pull Request on GitHub
```

---

## 📋 Recommended Next Steps

### 1. Set Up Branch Protection (Recommended)
1. Go to Settings → Branches
2. Add rule for `main` branch
3. Require pull request reviews before merging

### 2. Add Issue Templates
1. Create `.github/ISSUE_TEMPLATE/bug_report.md`
2. Create `.github/ISSUE_TEMPLATE/feature_request.md`

### 3. Add GitHub Actions (CI/CD)
1. Create `.github/workflows/tests.yml`
2. Auto-run tests on push/PR

### 4. Set Up GitHub Discussions
1. Go to Settings → Features
2. Enable "Discussions"
3. Users can ask questions without creating issues

---

## 🎓 Git Commands Cheatsheet

```powershell
# Basic workflow
git status                           # See what changed
git add .                           # Stage changes
git commit -m "message"             # Commit
git push                            # Push to GitHub
git pull                            # Pull from GitHub

# Branching
git checkout -b feature/name        # Create feature branch
git checkout main                   # Switch to main
git branch                          # List branches
git branch -d feature/name          # Delete branch

# History
git log                             # See commit history
git diff                            # See changes
git show <commit>                   # See specific commit

# Fixing mistakes
git reset --hard HEAD               # Discard changes
git revert <commit>                 # Undo a commit
git reset HEAD~1                    # Undo last commit (keep changes)
```

---

## 🚨 Important Notes

### File Sizes
- GitHub recommends files under 100 MB
- Our Excel files should be fine (< 1 MB)
- Large binary files should use Git LFS (optional)

### .gitignore
We already have a `.gitignore` file that excludes:
- `*.pyc` - Compiled Python files
- `__pycache__/` - Python cache
- `.env` - Environment variables
- `*.xlsx` - Excel files (except the template)
- Large PPT files
- IDE folders (.vscode, .idea)

---

## 🎉 You're Done!

Your repository is now on GitHub and ready for:
- ✅ Version control
- ✅ Collaboration
- ✅ Issue tracking
- ✅ Pull requests
- ✅ Open source contributions

---

## 💡 Tips

1. **Keep commits small** - Each commit should do one thing
2. **Use meaningful messages** - Future you will thank you
3. **Review before pushing** - Use `git diff` to check
4. **Pull before pushing** - Avoid conflicts with `git pull`
5. **Tag releases** - Use `git tag v1.0.0` for releases

---

## 📞 Need Help?

- **GitHub Docs:** https://docs.github.com
- **Git Docs:** https://git-scm.com/doc
- **GitHub CLI:** https://cli.github.com (alternative to web interface)

---

## 🔐 Authentication Issues?

If you get "403 Forbidden" or authentication errors:

```powershell
# Reset credentials
git config --global user.name "Your Name"
git config --global user.email "your.email@example.com"

# Generate new token at: https://github.com/settings/tokens
# Use token as password when prompted
```

---

**You're all set! Start collaborating! 🚀**
