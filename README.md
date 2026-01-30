# Reviewer Assignment Automation

Automated weekly assignment of literature review articles to reviewers using GitHub Actions.

## 🔄 How It Works

1. **Every Monday at 10:00 AM IST**, GitHub Actions triggers the workflow
2. Python script connects to MySQL database
3. Runs the reviewer assignment query
4. Generates an Excel report
5. Sends email notifications to all reviewers

## 📁 Project Structure

```
reviewer-assignment-automation/
├── .github/
│   └── workflows/
│       └── weekly_assignment.yml    # GitHub Actions workflow
├── scripts/
│   └── reviewer_assignment.py       # Main Python script
├── requirements.txt                  # Python dependencies
└── README.md
```

## 🚀 Setup Instructions

### Step 1: Create GitHub Repository

1. Create a new **private** repository on GitHub
2. Clone it to your local machine
3. Copy all these files to the repository

### Step 2: Configure GitHub Secrets

Go to your repository → **Settings** → **Secrets and variables** → **Actions** → **New repository secret**

Add the following secrets:

#### Database Secrets
| Secret Name | Description | Example |
|-------------|-------------|---------|
| `DB_HOST` | MySQL server hostname | `your-db-server.com` |
| `DB_PORT` | MySQL port | `3306` |
| `DB_NAME` | Database name | `your_database` |
| `DB_USER` | Database username | `db_user` |
| `DB_PASSWORD` | Database password | `your_password` |

#### Email Secrets (SMTP)
| Secret Name | Description | Example |
|-------------|-------------|---------|
| `SMTP_HOST` | SMTP server | `smtp.gmail.com` |
| `SMTP_PORT` | SMTP port | `587` |
| `SMTP_USER` | SMTP username | `your-email@gmail.com` |
| `SMTP_PASSWORD` | SMTP password/App password | `your_app_password` |
| `EMAIL_FROM` | Sender email address | `noreply@yourcompany.com` |

#### Reviewer Email Secrets
| Secret Name | Description |
|-------------|-------------|
| `REVIEWER_1_EMAIL` | Reviewer 1's email |
| `REVIEWER_2_EMAIL` | Reviewer 2's email |
| `REVIEWER_3_EMAIL` | Reviewer 3's email |
| `REVIEWER_4_EMAIL` | Reviewer 4's email |
| `REVIEWER_5_EMAIL` | Reviewer 5's email |

### Step 3: Push Code to GitHub

```bash
git add .
git commit -m "Initial setup for reviewer assignment automation"
git push origin main
```

### Step 4: Test the Workflow

1. Go to your repository on GitHub
2. Click **Actions** tab
3. Select **Weekly Reviewer Assignment** workflow
4. Click **Run workflow** → **Run workflow** (manual trigger for testing)

## ⏰ Schedule

The workflow runs automatically:
- **Every Monday at 10:00 AM IST** (4:30 AM UTC)

To change the schedule, edit the cron expression in `.github/workflows/weekly_assignment.yml`:

```yaml
schedule:
  - cron: '30 4 * * 1'  # Current: Monday 10:00 AM IST
```

Cron format: `minute hour day-of-month month day-of-week`

| Time (IST) | Cron Expression |
|------------|-----------------|
| Monday 9:58 AM | `28 4 * * 1` |
| Monday 10:00 AM | `30 4 * * 1` |
| Monday 10:30 AM | `0 5 * * 1` |

## 📧 Email Configuration

### Gmail Setup
1. Enable 2-Factor Authentication on your Google account
2. Generate an App Password: Google Account → Security → App passwords
3. Use the app password as `SMTP_PASSWORD`

### Outlook/Office 365 Setup
- SMTP Host: `smtp.office365.com`
- SMTP Port: `587`

## 🔧 Troubleshooting

### Check Workflow Logs
1. Go to **Actions** tab
2. Click on the failed workflow run
3. Expand the failed step to see error details

### Common Issues

| Issue | Solution |
|-------|----------|
| Database connection failed | Check DB_HOST is accessible from GitHub (public IP or use VPN/tunnel) |
| Email not sending | Verify SMTP credentials, check spam folder |
| Query returns no data | Verify date range and data exists |

## ⚠️ Important Notes

1. **Database Access**: Your MySQL database must be accessible from the internet (GitHub's IP ranges). Consider using:
   - Cloud database with public access
   - VPN/SSH tunnel
   - GitHub self-hosted runners inside your network

2. **Security**: All credentials are stored as GitHub Secrets (encrypted)

3. **Timezone**: GitHub Actions uses UTC. The cron is set for 4:30 AM UTC = 10:00 AM IST

## 📝 License

Internal use only.
