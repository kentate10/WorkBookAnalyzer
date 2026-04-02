# Deployment Guide - GitHub & Streamlit Cloud

This guide will walk you through deploying your WorkBook Analyzer to Streamlit Community Cloud.

## 📋 Prerequisites

- GitHub account (create one at https://github.com if you don't have one)
- Git installed (✅ Already installed - version 2.52.0)

## 🚀 Step-by-Step Deployment

### Step 1: Configure Git (One-time setup)

Open PowerShell or Command Prompt and run these commands with YOUR information:

```bash
# Set your name (this will appear in commits)
git config --global user.name "Your Name"

# Set your email (use the email associated with your GitHub account)
git config --global user.email "your.email@example.com"
```

**Example:**
```bash
git config --global user.name "John Doe"
git config --global user.email "john.doe@example.com"
```

### Step 2: Set Up SSH Key for GitHub (Recommended)

#### 2a. Generate SSH Key

```bash
# Create .ssh directory
mkdir ~/.ssh

# Generate SSH key (press Enter for all prompts to use defaults)
ssh-keygen -t ed25519 -C "your.email@example.com"
```

This will create two files:
- `~/.ssh/id_ed25519` (private key - keep this secret!)
- `~/.ssh/id_ed25519.pub` (public key - this goes to GitHub)

#### 2b. Copy Your Public Key

```bash
# Display your public key
cat ~/.ssh/id_ed25519.pub
```

Copy the entire output (starts with `ssh-ed25519` and ends with your email).

#### 2c. Add SSH Key to GitHub

1. Go to https://github.com/settings/keys
2. Click **"New SSH key"**
3. Give it a title (e.g., "My Laptop")
4. Paste your public key in the "Key" field
5. Click **"Add SSH key"**

#### 2d. Test SSH Connection

```bash
ssh -T git@github.com
```

You should see: `Hi username! You've successfully authenticated...`

### Step 3: Create GitHub Repository

#### Option A: Via GitHub Website (Easier)

1. Go to https://github.com/new
2. Repository name: `WorkBookAnalyzer` (or your preferred name)
3. Description: `Financial dashboard for project management workbooks`
4. Choose **Public** (required for free Streamlit deployment)
5. **DO NOT** initialize with README, .gitignore, or license (we already have these)
6. Click **"Create repository"**

#### Option B: Via GitHub CLI (if installed)

```bash
gh repo create WorkBookAnalyzer --public --source=. --remote=origin
```

### Step 4: Commit and Push Your Code

Now, in your project directory (`c:/Users/KenderTate/Desktop/CodeShi/WorkBookAnalyzer`), run:

```bash
# Add all files to git
git add .

# Create your first commit
git commit -m "Initial commit: WorkBook Analyzer MVP"

# Add your GitHub repository as remote (replace YOUR_USERNAME)
git remote add origin git@github.com:YOUR_USERNAME/WorkBookAnalyzer.git

# Push to GitHub
git branch -M main
git push -u origin main
```

**Replace `YOUR_USERNAME`** with your actual GitHub username!

### Step 5: Deploy to Streamlit Cloud

#### 5a. Sign Up for Streamlit Cloud

1. Go to https://share.streamlit.io/
2. Click **"Sign up"**
3. Sign in with your GitHub account
4. Authorize Streamlit to access your repositories

#### 5b. Deploy Your App

1. Click **"New app"**
2. Select your repository: `YOUR_USERNAME/WorkBookAnalyzer`
3. Branch: `main`
4. Main file path: `app.py`
5. Click **"Deploy!"**

#### 5c. Wait for Deployment

- Streamlit will install dependencies from `requirements.txt`
- This takes 2-5 minutes
- You'll see logs showing the installation progress
- Once complete, your app will be live!

### Step 6: Access Your App

Your app will be available at:
```
https://YOUR_USERNAME-workbookanalyzer-app-RANDOM.streamlit.app
```

You can customize the URL in Streamlit Cloud settings.

## 🔧 Alternative: Using HTTPS Instead of SSH

If you prefer HTTPS over SSH:

### Configure Git Credential Manager

```bash
# This will prompt for username/password on first push
git config --global credential.helper manager
```

### Use HTTPS Remote URL

```bash
# Add remote with HTTPS (replace YOUR_USERNAME)
git remote add origin https://github.com/YOUR_USERNAME/WorkBookAnalyzer.git

# Push (will prompt for credentials)
git push -u origin main
```

**Note:** GitHub now requires Personal Access Tokens instead of passwords for HTTPS.

### Create Personal Access Token

1. Go to https://github.com/settings/tokens
2. Click **"Generate new token (classic)"**
3. Give it a name: "WorkBook Analyzer"
4. Select scopes: `repo` (full control of private repositories)
5. Click **"Generate token"**
6. **Copy the token** (you won't see it again!)
7. Use this token as your password when pushing

## 📝 Quick Reference Commands

### Check Git Status
```bash
git status
```

### Make Changes and Update
```bash
# After making code changes
git add .
git commit -m "Description of changes"
git push
```

### View Commit History
```bash
git log --oneline
```

### Check Remote URL
```bash
git remote -v
```

## 🐛 Troubleshooting

### "Permission denied (publickey)"
- Your SSH key isn't set up correctly
- Make sure you added the public key to GitHub
- Test with: `ssh -T git@github.com`

### "Repository not found"
- Check the remote URL: `git remote -v`
- Make sure you replaced YOUR_USERNAME with your actual username
- Verify the repository exists on GitHub

### "Failed to push some refs"
- Someone else pushed changes (unlikely for new repo)
- Pull first: `git pull origin main --rebase`
- Then push: `git push origin main`

### Streamlit Deployment Fails
- Check `requirements.txt` has all dependencies
- Verify `app.py` is in the root directory
- Check Streamlit Cloud logs for specific errors

### Large File Error
- Excel files are already in `.gitignore`
- If you get errors about large files, they shouldn't be committed
- Remove from git: `git rm --cached filename.xlsx`

## 🔒 Security Notes

### What's Safe to Commit
✅ Python code
✅ Configuration files
✅ Documentation
✅ Requirements.txt

### What NOT to Commit
❌ Excel workbooks with sensitive data
❌ API keys or passwords
❌ Personal information
❌ Large binary files

The `.gitignore` file already excludes:
- `*.xlsx` files
- `*.xlsm` files
- `*.csv` files
- Temporary files
- Python cache files

## 📱 Sharing Your App

Once deployed, you can:
1. Share the Streamlit URL with anyone
2. Users upload their own workbooks
3. All processing happens in their browser session
4. No data is stored on the server

## 🔄 Updating Your App

When you make changes:

```bash
# 1. Make your code changes
# 2. Commit and push
git add .
git commit -m "Description of what you changed"
git push

# 3. Streamlit Cloud will automatically redeploy!
```

## 📊 Monitoring Your App

In Streamlit Cloud dashboard:
- View app logs
- See usage statistics
- Manage settings
- Restart app if needed

## 💡 Tips

1. **Test Locally First**: Always run `streamlit run app.py` locally before pushing
2. **Small Commits**: Make frequent, small commits with clear messages
3. **Branch for Features**: Use branches for major changes
4. **Check Logs**: If deployment fails, check Streamlit Cloud logs
5. **Keep Dependencies Updated**: Regularly update `requirements.txt`

## 🆘 Need Help?

- **Git Issues**: https://docs.github.com/en/get-started
- **Streamlit Cloud**: https://docs.streamlit.io/streamlit-community-cloud
- **SSH Setup**: https://docs.github.com/en/authentication/connecting-to-github-with-ssh

---

## ✅ Quick Checklist

Before deploying, make sure:
- [ ] Git is configured with your name and email
- [ ] SSH key is added to GitHub (or using HTTPS with token)
- [ ] GitHub repository is created
- [ ] Code is committed and pushed
- [ ] Streamlit Cloud account is set up
- [ ] App is deployed and accessible

---

**Ready to deploy?** Follow the steps above and your app will be live in minutes!