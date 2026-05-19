# TikTok Affiliate Creator Hub — Setup Guide

A private web dashboard for managing TikTok Shop affiliate creators. Track creator performance, analyze top hooks/CTAs/selling points, annotate videos, and share a read-only view with your team — all from a single URL.

---

## What You'll Need

- **GitHub account** (free) — stores your code
- **Vercel account** (free) — hosts the website
- **Supabase account** (free tier works fine) — stores the data
- About 30 minutes

---

## Step 1 — Set Up Supabase (Database)

### 1.1 Create a Project

1. Go to [supabase.com](https://supabase.com) and sign up or log in
2. Click **New Project**
3. Name it (e.g. `tiktok-creator-hub`), set a strong database password, pick a region near you
4. Wait ~2 minutes for the project to finish launching

### 1.2 Create the Database Tables

In your Supabase project, click **SQL Editor** in the left sidebar → **New Query**, then paste and run the following:

```sql
-- Main video data (uploaded from TikTok Shop CSV exports)
create table if not exists tiktok_reports (
  id text primary key,
  source text not null,
  video_id text,
  video_link text,
  creator text,
  revenue numeric default 0,
  items_sold integer default 0,
  views integer default 0,
  likes integer default 0,
  comments integer default 0,
  description text,
  hashtags text,
  product text,
  date_posted text,
  audio_hook text,
  selling_points text,
  key_idea text,
  transcript text,
  rank integer default 0
);

-- Manually edited fields per video (hooks, CTAs, video length, etc.)
create table if not exists tiktok_overrides (
  report_id text primary key,
  audio_hook text,
  visual_hook text,
  text_hook text,
  video_length text,
  cta text,
  selling_points text,
  key_idea text,
  updated_at timestamptz default now()
);

-- App settings and cached aggregation data
create table if not exists tiktok_hub_settings (
  key text primary key,
  value text,
  updated_at timestamptz default now()
);

-- Hidden videos (admin can hide videos from creator view)
create table if not exists tiktok_hidden_videos (
  video_id text primary key
);
```

Click **Run**. You should see "Success. No rows returned" for each statement.

### 1.3 Disable Row Level Security

Run this in the SQL Editor to allow the app to read and write all tables:

```sql
alter table tiktok_reports disable row level security;
alter table tiktok_overrides disable row level security;
alter table tiktok_hub_settings disable row level security;
alter table tiktok_hidden_videos disable row level security;
```

### 1.4 Copy Your API Keys

1. In your Supabase project, go to **Settings → API**
2. Copy your **Project URL** — looks like `https://abcdefghij.supabase.co`
3. Copy the **anon / public** key under "Project API keys"
4. Save both somewhere — you'll need them in Step 3

### 1.5 Enable Email Authentication

1. Go to **Authentication → Providers**
2. Confirm **Email** is enabled (it is by default)
3. Optional: go to **Authentication → Settings** and turn off "Confirm email" so passwords work immediately without needing email verification

---

## Step 2 — Set Up GitHub

### 2.1 Create a New Repository

1. Go to [github.com](https://github.com) and sign in
2. Click the **+** icon (top right) → **New repository**
3. Name it (e.g. `tiktok-creator-hub`), set it to **Private**
4. Leave everything else unchecked — do NOT initialize with a README
5. Click **Create repository**

### 2.2 Edit vercel.json Before Uploading

Open the `vercel.json` file in this template folder and replace `YOUR_SUPABASE_PROJECT_REF` with your actual Supabase project reference ID. This is the part of your Project URL before `.supabase.co`.

For example, if your URL is `https://abcdefghij.supabase.co`, your `vercel.json` should look like:

```json
{
  "rewrites": [
    {
      "source": "/api/sb/:path*",
      "destination": "https://abcdefghij.supabase.co/:path*"
    }
  ]
}
```

### 2.3 Upload the Code

**Option A — GitHub web interface (no coding needed):**
1. On your new empty GitHub repo page, click **uploading an existing file**
2. Drag and drop all files and folders from this `brand-template` folder
3. Make sure the folder structure is preserved (`src/` files stay in `src/`, etc.)
4. Write a commit message like "Initial setup" and click **Commit changes**

**Option B — Terminal:**
```bash
git clone https://github.com/your-username/tiktok-creator-hub.git
cd tiktok-creator-hub
# Copy all files from brand-template/ into this folder
git add .
git commit -m "Initial setup"
git push
```

---

## Step 3 — Deploy on Vercel

### 3.1 Import Your GitHub Repository

1. Go to [vercel.com](https://vercel.com) and sign up / log in with GitHub
2. Click **Add New → Project**
3. Find and click **Import** next to your `tiktok-creator-hub` repository
4. Under **Framework Preset**, select **Create React App**
5. **Do not deploy yet** — add environment variables first (next step)

### 3.2 Add Environment Variables

Still on the "Configure Project" screen, expand **Environment Variables** and add both of these:

| Name | Value |
|------|-------|
| `REACT_APP_SUPABASE_URL` | Your Supabase Project URL (e.g. `https://abcdefghij.supabase.co`) |
| `REACT_APP_SUPABASE_ANON_KEY` | Your Supabase anon/public key |

Click **Deploy**. Vercel builds and deploys in about 1–2 minutes. You'll get a live URL like `https://tiktok-creator-hub.vercel.app`.

---

## Step 4 — Create Your Admin Account

### 4.1 Create a User

1. In your Supabase project, go to **Authentication → Users**
2. Click **Add User → Create new user**
3. Enter your email address and a strong password
4. Click **Create User**

### 4.2 Grant Admin Privileges

The tool uses a metadata flag to distinguish admins. Run this SQL in the Supabase SQL Editor — replace the email with your actual email:

```sql
update auth.users
set raw_user_meta_data = raw_user_meta_data || '{"is_admin": true}'::jsonb
where email = 'you@yourbrand.com';
```

Verify it worked:
```sql
select email, raw_user_meta_data from auth.users;
```

You should see `{"is_admin": true}` in the metadata column for your user.

---

## Step 5 — Customize Your Branding

### 5.1 Update the Brand Name

In two files, replace the placeholder text with your actual brand name:

- `src/TikTokShopReporter.tsx` — find `YOUR BRAND` (in the header)
- `src/Login.tsx` — find `Your Brand` (on the login screen)

### 5.2 Add Your Logo (Optional)

1. Add your logo image (PNG, ideally square, ~200×200px) to the `public/` folder — name it `logo.png`
2. In `src/TikTokShopReporter.tsx`, find the `📦` emoji placeholder in the header section and replace it with:
   ```jsx
   <img src="/logo.png" alt="Your Brand" style={{height:48,width:48,objectFit:"contain",borderRadius:"50%",background:"#fff",flexShrink:0}}/>
   ```

### 5.3 Push Changes

Commit and push your changes to GitHub — Vercel will automatically rebuild and redeploy.

---

## Step 6 — Sign In and Test

1. Open your Vercel URL in a browser
2. You'll see the public "Creator View" by default
3. Click **Enter Admin Mode** (top right)
4. Sign in with the email and password from Step 4
5. The **ADMIN MODE** badge should appear — you now have full access

---

## Using the Tool

### Uploading Reports

1. Click **Update Reports** (top right, admin mode only)
2. Select the report type: **All-Time**, **Last Month**, or **In-House**
3. Drag and drop your TikTok Shop CSV export
4. A progress overlay shows each stage — wait for it to complete
5. Repeat for each report type

**Where to export CSVs from TikTok:**
Log in to TikTok Shop Seller Center → **Creator Marketplace → Affiliate → Performance** → Export CSV for your date range.

**Required CSV columns:** `URL`, `Creator`, `Video Revenue`, `Items Sold`, `Views Count`. Other columns (Date, Description, Product, Hooks, Selling Points, etc.) are used if present. If you upload a wrong file, the tool will detect missing columns and reject it without changing any data.

---

### Dashboard Tabs

**All-Time High GMV** — All uploaded all-time videos sorted by revenue. Set a GMV filter to focus on top performers. Click any video card to expand and edit fields (visual hook, text hook, CTA, etc.).

**Last Month High GMV** — Same layout, scoped to last month's data.

**In-House High GMV** — For your own brand's produced videos.

**Top Creators** — One card per creator showing total videos, units sold, total GMV, and their top 3 videos embedded. Set a minimum GMV threshold to filter to your top tier.

**Hooks, CTA & Selling Points** — Aggregated performance by hook text, CTA phrase, and selling point. Each entry shows total videos using that hook, total views, units sold, GMV, and videos with any GMV. Expand any entry to see the top videos using it. Stats are computed from your full dataset — not just the displayed rows.

---

### Editing Video Fields

In Admin Mode, click the edit button on any video card to manually fill in:
- **Visual Hook** — what's shown on screen in the opening seconds
- **Text Hook** — on-screen text overlay
- **Audio Hook** — spoken hook (pre-filled from TikTok export if available)
- **CTA** — call to action phrase used
- **Video Length** — duration
- **Selling Points** — key product benefits mentioned
- **Key Idea** — one-line summary

Edits are saved to the database instantly and survive CSV re-uploads. They feed directly into the Hooks, CTA & Selling Points tab analysis.

---

### Publishing to Creator View

1. Use the **Visible / Hidden** buttons on video cards to control what appears publicly
2. Set GMV threshold filters if you want to limit what's shown
3. Click **Update Dashboard** to publish everything to the creator-facing URL

The public URL (your Vercel URL with no login) shows the creator view. Creators can browse top videos and creator stats but cannot edit anything.

---

### Exporting to Excel

Click **Export XLSX** to download all current data as a spreadsheet.

---

## Troubleshooting

**"Could not connect" on load**
Check that `REACT_APP_SUPABASE_URL` and `REACT_APP_SUPABASE_ANON_KEY` are set in Vercel → Project Settings → Environment Variables. Also confirm `vercel.json` has your correct Supabase project ref.

**Login fails with correct password**
Make sure the user was created in Supabase **Authentication → Users**, not just in the database. If email confirmation is enabled, check your inbox and confirm the email first — or disable email confirmation in Supabase Auth Settings.

**Logged in but "Enter Admin Mode" button doesn't work / no admin access**
Run the SQL in Step 4.2 to set `is_admin: true`. Then sign out completely and sign back in — the metadata is read at login time.

**Upload rejected with "missing columns" error**
You're uploading the wrong CSV type. The tool expects a TikTok Shop **affiliate performance** export. Other TikTok reports (product analytics, order exports, etc.) have different column names.

**Creator counts seem wrong after refresh**
After uploading a new CSV, always click **Update Dashboard** before sharing the link with creators. The admin view updates immediately; the public creator view only updates when you explicitly publish.

**Data looks correct in-session but resets after refresh**
This means the database write didn't complete during upload. The upload progress overlay should indicate success. If it's happening repeatedly, check your Supabase project's database usage — the free tier has a 500MB limit.

---

## Adding More Admin Users

To give another person admin access:

1. Create their user in Supabase **Authentication → Users**
2. Run:
   ```sql
   update auth.users
   set raw_user_meta_data = raw_user_meta_data || '{"is_admin": true}'::jsonb
   where email = 'their@email.com';
   ```

To create a read-only user (can log in but not edit), create the user but don't add the `is_admin` flag. They'll be able to see the tool but all edit controls will be hidden.

---

*Built with React, TypeScript, Supabase, and Vercel.*
