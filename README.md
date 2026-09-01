# Alloro Vineyards Winery Tracker — standalone version

This is the same app you've been using inside Claude, restructured to run as its
own independent website with its own database. Nothing about how the app
works has changed — the only thing swapped out is *where the data is stored*.

Follow these steps in order. Steps 1–2 need to happen before the app will work
at all. Steps 3–5 get it onto the internet.

---

## 1. Create a free Supabase project (your database)

1. Go to **https://supabase.com** and sign up (free tier is plenty for this).
2. Click **"New project"**. Pick any name (e.g. `alloro-winery`), set a database
   password (save it somewhere), and choose a region close to you.
3. Wait ~2 minutes for the project to finish setting up.
4. In the left sidebar, click **SQL Editor** → **New query**.
5. Open the `supabase-setup.sql` file included in this project, paste its
   entire contents into the editor, and click **Run**. This creates the one
   table the app needs.
6. In the left sidebar, click **Project Settings** → **API**. You'll need two
   values from this page in the next step:
   - **Project URL**
   - **anon / public** key (NOT the `service_role` key — that one must stay secret)

## 2. Configure the app with your Supabase details

1. In this project folder, copy `.env.example` to a new file named `.env`.
2. Paste in your Project URL and anon key from step 1.6:
   ```
   VITE_SUPABASE_URL=https://your-project-ref.supabase.co
   VITE_SUPABASE_ANON_KEY=your-anon-public-key
   ```
3. Save the file. (`.env` is already in `.gitignore` — it will never be
   uploaded to GitHub, which is correct: keep it private.)

## 3. Run it locally to test

You'll need [Node.js](https://nodejs.org) installed (any recent version).

```bash
npm install
npm run dev
```

Open the URL it prints (usually `http://localhost:5173`). The app should load
and behave exactly like it did in Claude — because it's the same code. Try
adding a work order to confirm it saves (check the Supabase dashboard →
**Table Editor** → `app_storage` to see the row appear).

## 4. Put the code on GitHub

1. Create a free account at **https://github.com** if you don't have one.
2. Create a new repository (e.g. `winery-tracker`), and follow GitHub's
   instructions to push this folder to it. Typically:
   ```bash
   git init
   git add .
   git commit -m "Initial version"
   git branch -M main
   git remote add origin https://github.com/YOUR-USERNAME/winery-tracker.git
   git push -u origin main
   ```

## 5. Deploy it live with Vercel

1. Go to **https://vercel.com** and sign up (free tier is fine), using
   "Continue with GitHub" so it can see your repositories.
2. Click **"Add New" → "Project"**, and import the repository you just pushed.
3. Vercel will auto-detect it's a Vite project — leave the build settings as
   default.
4. Before deploying, add your environment variables: in the project setup
   screen, find **Environment Variables** and add the same two values from
   your `.env` file:
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_ANON_KEY`
5. Click **Deploy**. In about a minute you'll get a live URL
   (like `winery-tracker.vercel.app`) that works for anyone, on any device,
   with no Claude account involved at all.
6. Optional: in Vercel's project settings under **Domains**, you can attach a
   custom domain (e.g. `tracker.allorovineyard.com`) if you own one.

From here on, whenever you want changes to the app, they get pushed to
GitHub and Vercel automatically redeploys the new version within a minute or
two.

---

## Important things to know

- **Data lives in Supabase now, not in Claude.** The data in your current
  Claude artifact does *not* automatically transfer here — this starts as a
  fresh, empty database. If you want to bring existing data over, the
  simplest path is exporting it from the Claude version (Export Data → Excel)
  and re-entering it here, since there isn't an automatic one-click migration.
- **Anyone with the link can view and edit data** — there's no login system
  built in (same as the Claude version). That's fine for an internal tool
  used by a small trusted crew, but if you ever want to restrict who can
  edit, that requires adding real user accounts (Supabase supports this, but
  it's a separate follow-up project).
- **Costs**: both Supabase and Vercel have generous free tiers that should
  comfortably cover a small operational tool like this. You'd only need to
  pay if usage grew far beyond what one winery needs.

## Getting an app icon on your phone (no App Store needed)

This project is already set up as a installable web app (a "PWA"), complete
with a home-screen icon, so your crew can get something that looks and feels
like a real app without going anywhere near the App Store.

**On iPhone (Safari):**
1. Open your live Vercel URL in Safari (must be Safari, not Chrome — Add to
   Home Screen only works from Safari on iOS).
2. Tap the **Share** icon (square with an arrow pointing up).
3. Scroll down and tap **"Add to Home Screen."**
4. Tap **Add** in the top-right.

A green icon with a gold "A" now appears on the home screen. Opening it
launches full-screen, with no browser address bar — it behaves like a real
installed app.

**On Android (Chrome):** Chrome will usually prompt "Add to Home Screen"
automatically, or you can trigger it from the ⋮ menu → "Add to Home Screen."

This is genuinely the easiest path for internal, team-only use — no Apple
Developer account, no app review, no cost. If down the road you want this in
the actual App Store (e.g. to distribute to a much larger team through
Apple Business Manager, or to sell as a public product), that's a
substantially bigger project — wrapping the app natively (e.g. with
Capacitor), enrolling in Apple's Developer Program, and going through App
Store review. Worth a separate conversation if that need ever comes up.

## Testing this in Xcode on your own iPhone (free, just for trying it out)

This project already has Capacitor set up, which wraps the web app into a
real native iOS project that Xcode can open. This lets you run it on your
own iPhone, plugged into your Mac, using nothing but your personal Apple ID
— no paid Developer account needed for this. (You'd only need the paid
account later if you want to install it on teammates' phones without a
cable, or publish it anywhere.)

**Requirements:** a Mac, with Xcode installed (free, from the Mac App
Store), and this project already running locally (you should already have
done `npm install` and have a working `.env` file from the earlier setup).

**Step 1 — Install Capacitor's pieces**

In your terminal, inside the project folder:
```
npm install
```
(This picks up the Capacitor packages already added to package.json.)

**Step 2 — Build the web app**

Capacitor wraps the *built* app, not the dev server:
```
npm run build
```
This creates a `dist` folder — that's what gets loaded inside the iOS app.

**Step 3 — Add the iOS project**

```
npx cap add ios
```
This creates a new `ios/` folder containing a real Xcode project. You only
need to run this command once, ever — after this, the `ios/` folder is a
permanent part of your project going forward.

**Step 4 — Copy the build into the iOS project**

```
npx cap sync ios
```
Run this command every time you rebuild the web app (`npm run build`) and
want that update reflected in Xcode.

**Step 5 — Open it in Xcode**

```
npx cap open ios
```
This launches Xcode with the project already open.

**Step 6 — Run it on your iPhone**

1. Plug your iPhone into your Mac with a cable.
2. In Xcode's toolbar (top-left area, next to the Play/Stop buttons), select
   your iPhone as the target device instead of a simulator.
3. Click the **Play** (▶) button.
4. The first time, your iPhone will show an "Untrusted Developer" warning —
   go to **Settings → General → VPN & Device Management** on the phone, tap
   your Apple ID under "Developer App," and tap **Trust**.
5. Run it again from Xcode, and the app installs and opens on your phone —
   a real app icon, no browser, no home-screen shortcut trick.

**One real limitation to know about:** apps installed this way (via a free
personal Apple ID, straight from Xcode) expire after about **7 days** and
need to be reinstalled from Xcode again. That's an Apple restriction on free
accounts, not something specific to this project — it's meant for testing,
not daily long-term use. For that, the paid Developer Program (or, more
simply, the PWA "Add to Home Screen" approach above) is the right tool.

## Getting this on TestFlight (so you can download and use it, no cable needed)

This is the real, permanent way to get the app on your phone through Apple's
own system — no 7-day expiration, no Xcode required after the first setup.
It does involve a real cost and a few genuine steps with Apple, so here's
the complete, honest path.

**What this costs:** $99/year for Apple's Developer Program. There's no way
around this — TestFlight is Apple's own distribution system and requires a
paid account, full stop.

**Decision to make first: Individual or Organization account?**
- **Individual** (recommended to start) — enrolls under your own name,
  usually approved within a day, no business paperwork. Good enough for
  TestFlight regardless of whether the app is for personal or business use.
- **Organization** (e.g. "Alloro Vineyard, Inc.") — makes the account
  belong to the business rather than you personally, but requires a
  D-U-N-S number (a business ID number — free to get if you don't have one,
  but can take 1–2 weeks) and proof you're authorized to bind the company.
  Worth doing eventually if you want the business's name on the account,
  but not necessary just to get this on your phone.

### Step 1 — Enroll in the Apple Developer Program
1. Go to **developer.apple.com/programs/enroll**.
2. Sign in with your Apple ID (the same one you use for the App Store).
3. Choose **Individual** (or **Organization**, per above).
4. Pay the $99 annual fee.
5. Wait for approval — usually same-day to 48 hours for Individual.

### Step 2 — Point Xcode at your new account
1. Open Xcode → **Settings** (or **Preferences**) → **Accounts**.
2. Click **+** and sign in with the same Apple ID you just enrolled with.
3. Open this project in Xcode (`npx cap open ios`, same as before).
4. Click the blue project file at the top of the left sidebar → select the
   **App** target → **Signing & Capabilities** tab.
5. Under **Team**, choose your name/organization (it'll now show your real
   paid account instead of just "Personal Team").
6. Leave **"Automatically manage signing"** checked — Xcode handles the
   certificates and provisioning for you.

### Step 3 — Create the app's record in App Store Connect
1. Go to **appstoreconnect.apple.com** and sign in.
2. **My Apps → the "+" button → New App.**
3. Fill in:
   - **Platform:** iOS
   - **Name:** "Alloro Winery Tracker" (must be unique across the whole App
     Store — if it's taken, try adding a distinguishing word)
   - **Primary language:** English
   - **Bundle ID:** select `com.allorovineyard.winerytracker` from the
     dropdown (this project is already configured with that ID)
   - **SKU:** any unique text for your own records, e.g. `alloro-winery-01`
4. Click **Create.**

### Step 4 — Build and upload from Xcode
1. Back in Xcode, at the top toolbar where it shows the device/simulator
   selector, choose **"Any iOS Device (arm64)"** — not a simulator, not your
   specific phone, since you're building for upload, not for a cable
   connection this time.
2. Menu bar: **Product → Archive.** This takes a minute or two.
3. When it finishes, the **Organizer** window opens automatically showing
   your archive. Click **Distribute App.**
4. Choose **App Store Connect → Upload**, then click through the prompts
   (defaults are fine — keep automatic signing).
5. You may be asked an **"Export Compliance"** question about encryption —
   for an app like this that only uses standard HTTPS, the answer is
   "No" / "None of the algorithms mentioned above."
6. Click **Upload.** Apple will process the build — this can take anywhere
   from a few minutes to a couple of hours. You'll get an email when it's
   ready.

### Step 5 — Set up TestFlight and add yourself as a tester
1. In App Store Connect, open your app → **TestFlight** tab.
2. Once your build appears (after processing finishes), click on it —
   you may need to answer a couple of quick compliance questions.
3. Scroll to **Internal Testing → click "+"** to create a group (e.g.
   "Alloro Crew"), and add yourself — as the account holder, you're already
   available to add immediately, no waiting on Apple review for internal
   testers.
4. To add teammates later, you can either add them here as internal testers
   (they need to be invited as Users in App Store Connect under **Users and
   Access**, using their Apple ID email), or set up **External Testing**
   instead, which supports up to 10,000 testers via just an email or public
   link — but external testers do require a quick Apple review first
   (usually 24–48 hours), unlike internal testers.

### Step 6 — Install it on your phone
1. On your iPhone, download the **TestFlight** app from the App Store
   (it's free, made by Apple).
2. Open TestFlight — if you're an internal tester, "Alloro Winery Tracker"
   should already be listed. If not, check your email for an invite link.
3. Tap **Install.**

From here on, whenever you want to push an update: rebuild
(`npm run build && npx cap sync ios`), repeat Step 4 (Archive → Upload), and
everyone with TestFlight installed gets a notification that a new version
is ready.

