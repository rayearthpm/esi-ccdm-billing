# ESI CCDM Billing server

Your billing app with a login and shared results. The publisher uploads the monthly database, clicks
**Publish to viewers**, and everyone else signs in to see the Dashboard and download the bills.
No npm packages to install. It needs only **Node.js 22.5 or newer** (https://nodejs.org).

## Run it (any Windows, Mac or Linux PC or server)
    cd esi-billing-server
    node server.js
Open http://localhost:3000. On the first run it prints the publisher username and password. Write it down.
To choose your own: `ADMIN_USER=admin ADMIN_PASSWORD=YourPassword node server.js` (Windows PowerShell:
`$env:ADMIN_PASSWORD="YourPassword"; node server.js`) on the very first start.

## Everyday use
1. Sign in as the publisher, upload the monthly database, pick the billing month.
2. Click **Publish to viewers** (about a minute). Done.
3. **Users** (top right) adds viewers. Give each bank or colleague their own login. They only see results.
   Type bank codes in the **Banks** box (for example `PNB` or `BDO,SBC`) to limit a viewer to those banks; blank means all banks.
   Use the **Banks** button on an existing person to change it. The server removes other banks' numbers, bills and downloads before sending anything.

Other computers on your network open `http://<this-computer's-IP>:3000`.

## Easiest: read the folder from this PC (OneDrive desktop app)
If the OneDrive app syncs your monthly folder to this computer, no link or sign-in is needed. Put the folder's path in `data/config.json`
(use forward slashes):

    {"LOCAL_FOLDER":"C:/Users/YourName/OneDrive/CCDM/Billing_Database"}
The month list and auto-load work exactly as for the online folder. Files that are only in the cloud (the cloud icon in File Explorer)
are downloaded by OneDrive when the server opens them. To avoid the wait, right-click the folder > **Always keep on this device**.
The PC must be on, and signed in to OneDrive, whenever the server runs.

## Personal OneDrive without a public link (recommended for personal accounts)
Some share links cannot be read by a server. Instead, sign in to your Microsoft account once and the folder can stay private.

1. Go to https://portal.azure.com and sign in with your Microsoft account. Search for **App registrations** and click **New registration**.
2. Name: `ESI Billing`. Supported account types: **Personal Microsoft accounts only** (or "...any organizational directory and personal").
3. Redirect URI: choose the platform **Public client/native (mobile & desktop)** and enter `http://localhost:3000/auth/callback`. Click Register.
4. Copy the **Application (client) ID** from the overview page.
5. Under **API permissions**, add *Microsoft Graph > Delegated > Files.Read* and *offline_access* (if not already listed).
6. Put this in `data/config.json` (use your folder's path inside OneDrive, without the drive name):

        {"MS_CLIENT_ID":"paste-the-client-id","ONEDRIVE_PATH":"Documents/CCDM/Billing_Database"}
7. Start the server, sign in to the billing page and click **Connect OneDrive**. Approve the Microsoft screen. Back in the billing page, click **Refresh list**.
It connects once; the server keeps the sign-in in `data/ms_token.json` (keep that file private). Only read access to your files is requested.
Open the billing page at `http://localhost:3000` if you can; the Microsoft redirect always returns to localhost:3000 (change with `MS_REDIRECT` if you use another address or port).

## One workbook per month (OneDrive folder)
Keep every month's workbook in one OneDrive folder, named with the month, for example `CCDM_DB_OCT_2026.xlsm`
(`SEP_2026`, `September 2026` or `2026-09` also work). Point the server at the folder instead of a single file:

    ONEDRIVE_FOLDER="https://1drv.ms/f/s!..." node server.js      (or "ONEDRIVE_FOLDER" in data/config.json)

The publisher then gets a **month list** in the upload bar. The newest month loads automatically when the page opens.
Pick another month, or change the billing month in the generator, and the app switches to that month's file.
**Refresh list** picks up files you have just added.
- Personal OneDrive: share the folder as "Anyone with the link can view".
- Work/SharePoint folders cannot be listed through a public link. They need the private setup below
  (`MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET`), and the app needs read access to that folder.

## Loading a single workbook from OneDrive (no upload)
The publisher gets a **Refresh from OneDrive** button, and the workbook loads by itself each time the page opens.

**Easiest (any OneDrive):** in OneDrive, open the workbook, Share > "Anyone with the link" > **Can view**, copy the link, then start the server with it:

    ONEDRIVE_LINK="https://1drv.ms/x/s!..." node server.js
(Windows PowerShell: `$env:ONEDRIVE_LINK="https://..."; node server.js`.)
Or put it in `data/config.json`: `{"ONEDRIVE_LINK":"https://..."}`. Anyone holding that link can read the file, so only use it for a file you are happy to share that way.

**Private (Microsoft 365 / SharePoint):** your IT admin registers an app in Microsoft Entra ID with the *application* permission
`Files.Read.All` (or `Sites.Read.Selected` for one site) and gives you the tenant ID, client ID and a client secret. Add them to `data/config.json` along with the file's normal sharing link:

    {"ONEDRIVE_LINK":"https://yourco.sharepoint.com/:x:/s/...","MS_TENANT_ID":"...","MS_CLIENT_ID":"...","MS_CLIENT_SECRET":"..."}
The link then does not need to be public. If the button reports an error, the message says why (for example, link not set to "anyone can view").
Publishing to viewers is still one click after the refresh.

## Put it online with Render (https, always on)
1. Create a free GitHub account, then a **new private repository** (for example `esi-ccdm-billing`).
2. In the repository, click **Add file > Upload files** and drag in everything *inside* the `esi-billing-server` folder
   (`server.js`, `public`, `Dockerfile`, `render.yaml`, and so on). Do not upload `node.exe` or a `data` folder. Click **Commit**.
3. Create an account at render.com, then **New > Blueprint**, connect GitHub and pick the repository.
4. Render asks for **ADMIN_PASSWORD**. Type a strong password; the publisher username is `admin`.
5. Click **Apply**. The first build takes a few minutes. You get an address like `https://esi-ccdm-billing.onrender.com`.
6. Open it, sign in as `admin`, add your viewers on the **Users** page.
The paid "Starter" plan (about $7 per month plus about $0.25 for the 1 GB disk) is needed so the server stays on and keeps its data.
On a cloud host the server cannot read a folder on your PC, so the publisher uploads the month's workbook in the page, then clicks **Publish to viewers**.
To update the app later, upload the new files to the same repository; Render redeploys by itself.

## Putting it on the internet
1. Rent a small Linux server (about $5/month) or use a company server, install Node 22.
2. Copy this folder there and run it (or `docker compose up -d`).
3. Point a domain at the server and use HTTPS. The easiest is Caddy: see `Caddyfile.example`.
   Do not expose it over plain http on the internet. Passwords would travel unprotected.
4. Keep it running after a reboot, on Linux with systemd:
        [Unit]
        Description=ESI Billing
        After=network.target
        [Service]
        WorkingDirectory=/opt/esi-billing-server
        ExecStart=/usr/bin/node server.js
        Restart=always
        User=www-data
        [Install]
        WantedBy=multi-user.target

## Backups
Everything is in `data/billing.db` (users and published results). Copy that file while the server is stopped,
or daily with `sqlite3 data/billing.db ".backup backup.db"`.

## Forgot the publisher password
    node server.js passwd admin NewPassword
Add a user from the command line: `node server.js adduser jane Password123 viewer`

## Notes
- Passwords are stored hashed (scrypt). Sessions last 12 hours. 5 wrong logins locks that user for 5 minutes.
- Only the publisher can change results; viewers can only read and download.
- To update the app itself, replace `public/app.html` with the new index.html. Restart not needed.

## Live window (one running OneDrive file)

Publishers see a **Live | Billing** toggle at the top of the page and land on Live. Live shows the dashboard for one file that you keep editing in OneDrive, with a "Connected to …" line instead of an upload bar. The server checks the file every 5 minutes and reloads the numbers only when the file has changed.

* The file defaults to `CCDM/CCDM_DB_Live.xlsm` in your OneDrive (the same Microsoft sign-in you already use: `MS_CLIENT_ID`, then **Connect OneDrive** once on the main page).
* To use a different file set `ONEDRIVE_LIVE_FILE` (path inside your OneDrive, for example `CCDM/Other.xlsm`).
* Other ways to point at the file: `LIVE_LOCAL_FILE` (a full path on the computer running the server, such as a synced OneDrive folder) or `ONEDRIVE_LIVE_LINK` (a share link).
* The file should have the same sheets as the monthly database (Deployed, Pullout, Services).
* The Live is open to every signed-in user and shows all banks. Bank limits apply to Billing only.
