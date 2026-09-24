# Handler Monitoring Dashboard (Intest 2)

A real-time monitoring dashboard for **Handler Repair tracking** and **Spare Parts Inventory management** with live Power Automate flow synchronization, SharePoint Excel Online integration, responsive glassmorphic UI, and Light/Dark themes.

---

## 🚀 How to Deploy This to GitHub & GitHub Pages

Follow these simple steps to deploy this dashboard to your GitHub account:

### Step 1: Create a new repository on GitHub
1. Go to [github.com/new](https://github.com/new).
2. Name your repository (e.g. `handler-monitoring-dashboard`).
3. Set visibility to **Public** or **Private**.
4. Leave "Add a README file" **unchecked** (we already have everything).
5. Click **Create repository**.

---

### Step 2: Push this project to GitHub
In your local terminal or command prompt inside this project folder, run:

```bash
# Initialize git repository (if not already done)
git init

# Add all project files
git add .

# Commit files
git commit -m "Deploy Handler Monitoring Dashboard to GitHub"

# Set main branch
git branch -M main

# Link your GitHub repository (replace with your actual repository URL)
git remote add origin https://github.com/<YOUR-GITHUB-USERNAME>/<YOUR-REPOSITORY-NAME>.git

# Push to GitHub
git push -u origin main
```

---

### Step 3: Enable Automatic GitHub Pages Deployment
1. In your GitHub repository, click on **Settings** (top navigation tab).
2. On the left sidebar, click **Pages** (under the "Code and automation" section).
3. Under **Build and deployment > Source**, select **GitHub Actions**.
4. The included workflow (`.github/workflows/deploy.yml`) will automatically build and publish your site!
5. After about 1 minute, your live website URL will appear at the top of the Pages settings page:
   `https://<YOUR-GITHUB-USERNAME>.github.io/<YOUR-REPOSITORY-NAME>/`

---

## ⚡ Alternative Option: Instant 1-Click Single-File Hosting
If you prefer not to use a build system, a complete self-contained HTML version is also available in:
`public/standalone.html`

You can rename `public/standalone.html` to `index.html` on a static branch or upload it directly to any static web host.

---

## 🛠 Features Included
- **Handler Repair Tracking**: Real-time table synced via Power Automate to your SharePoint Excel workbook.
- **Image Evidence Capture**: High-efficiency client-side compression with full-screen lightbox preview.
- **Spare Parts Inventory**: Stock quantities with real-time add/withdraw adjustments and low-stock indicators.
- **Filter & Search**: Quick search across all columns, plus status and category filters.
- **Data Export**: 1-click export to CSV for both repair records and spare parts inventory.
- **Theme & Density**: Switch between Dark Mode / Light Mode and Compact / Normal row density.
- **Admin Access**: Toggle edit privileges with default credentials (`admin` / `admin123`).
