# Security guide – Metal printing cost estimator

No website is "unhackable". The goal is to make attacks hard, limit the damage if something goes wrong, and notice quickly.
This site has two parts, and each is protected differently.

| Part | What it is | Main risk | Protection built in |
|---|---|---|---|
| Website (HTML/CSS/JS) | Static files, visible to anyone who opens them | Code injection (XSS), fake copies, clickjacking, tampered libraries | Escaping of all data, strict Content-Security-Policy, self-hosted libraries, anti-framing, security headers |
| Google Sheet backend (`Code.gs`) | Stores quotes and price lists | Anyone with the URL reading, changing or deleting data | Per-user name + password login (salted hashes, session token, 30-min idle timeout), lockout after wrong tries, input validation, formula-injection guard, audit log |

## A. One-time setup (10 minutes)

1. **Redeploy the new `Code.gs`:** Deploy > Manage deployments > pencil > Version: **New version** > Deploy. Run `setup` once to authorise (it creates the **Users** tab).
2. **Create the first user (admin):** in `Code.gs` find `addUser()`, type a user name, a strong password (12+ characters) and `'admin'` between the quotes, select `addUser` in the toolbar and click **Run**. Then **delete the password from the code again** and save. Repeat for each staff member (role `'user'`). Running `addUser` for an existing user name resets that person's password.
3. **Set the admin e-mail:** in `Code.gs` type the admin's address into `ADMIN_EMAIL` near the top, redeploy, then run `setup` again once and accept the new permission (sending e-mail). Only this address ever receives a password-reset code, and only for admin accounts.
4. **Paste the `/exec` URL** into `API_URL` at the top of `script.js` (it stays the same when you redeploy a new version of the same deployment).
5. **Keep the Google Sheet private.** Share it only with the few people who manage it. The Users tab holds password hashes (never the passwords), but still limit who can open the sheet.
6. **Turn on 2-Step Verification** on the Google account that owns the sheet and script. This is the single most important protection for the backend.

## B. Hosting (choose one)

* **Cloudflare Pages / Netlify** (recommended): upload the folder. The `_headers` file adds the security headers automatically. HTTPS is automatic.
* **Any other host:** make sure HTTPS is forced, and add the headers from `_headers` in the host / Cloudflare settings.
* **Hide the whole site from the public (strongest):** put it behind **Cloudflare Access** (free for up to 50 users) or Netlify password protection, so only logged-in staff can even load the page. This also hides your pricing logic, which is otherwise readable in `script.js` by anyone who can open the page.

## C. Day-to-day rules

* Give every person their own user name and password. Never share one login. Send the first password privately and ask them to keep it secret.
* **Users tab (admins only, in the website):** add users, set a new password for a user, disable / enable. Users cannot change their own password; ask the admin.
* **Forgot admin password:** on the sign-in screen type the admin user name and click *Forgot admin password?*. A 6-digit code (valid 10 minutes, 5 tries, 3 requests per 10 minutes) is e-mailed to `ADMIN_EMAIL`; enter it with a new password.
* The editor functions still work too. To remove someone: run `disableUser()` (type their user name first). Their open session stops working on its next request. To reset a forgotten password run `addUser()` again with the same user name.
* The page signs out automatically after 30 minutes of inactivity (the server also ends idle sessions), and has a **Sign out** button.
* Check the **Audit** tab in the sheet now and then: `LOGIN_FAILED` rows mean someone is guessing passwords (the User column shows the name tried). If you see many for a real user, reset that password.
* After 5 wrong passwords for one user name, that user name is locked for 10 minutes (and 40 wrong tries overall lock everyone for 10 minutes). An attacker can therefore annoy a user by locking them, but cannot get in.
* Use Google Sheets *File > Version history* as your backup. Download a copy of the sheet monthly.
* Never paste passwords, or your Google password, into chat tools, emails or screenshots.
* Keep `libs/` files as they are. They are copies of ExcelJS 4.4.0, jsPDF 2.5.1 and jsPDF-AutoTable 3.5.31, hosted by you, so no outside server can change the code your page runs. Update them yourself from time to time (download new versions from npm).

## D. What this does NOT protect against (be aware)

* Someone who steals a staff member's password or Google account. Use strong passwords, 2-Step Verification and sign out on shared PCs.
* Anyone who can open the page can read the calculation *formulas* in `script.js` (unless you use Cloudflare Access, section B). Machine prices, powder prices, interest rates and all default rates are **not** in the website files: they are kept in the Google Sheet (Powders, Machines, **Settings** tabs) and are sent only after a successful sign-in.
* The login check runs in Google Apps Script, which cannot see the visitor's IP address, so it cannot block individual attackers, only lock a user name (or everyone) for 10 minutes.
* Passwords are hashed with salted, iterated SHA-256 (the best Apps Script offers). Long passwords matter; there is no password-change page, so the admin resets passwords from the editor.

## E. Backend protections (API)

* Every request except sign-in and the admin forgot-password steps needs a valid session token. Wrong, expired or idle tokens are refused.
* **Roles are enforced on the server:** only admins can see or change users and change the powder / machine lists. A normal user cannot do it even by calling the API directly.
* Changing or resetting a password immediately ends that user's open sessions.
* Passwords: at least 10 characters with letters and numbers. Only salted hashes are stored (Users tab). No password, key or user detail is in any website file.
* Rate limit: 120 requests per minute per signed-in user. 5 wrong passwords lock a user name for 10 minutes.
* Unknown or malformed requests are refused. Errors never show internal details. The API URL (`GET`) shows no data and no version.
* **Default rates:** the first time the Powders, Machines or **Settings** tab is empty, the script fills it from the defaults at the top of `Code.gs`. After that, edit the values in the Sheet (the Settings tab holds shop rates such as preparation hours, margin, advance, furnace and wire-cut data). Keep the Sheet and Apps Script private.
* **Saved quotes are private per user:** the server returns, opens, overwrites and deletes only quotes saved by the signed-in user. (Old quotes saved before this rule have no owner and belong to the admin.)
* **Model pictures:** a quote's model picture is stored in the **Images** tab (small JPEG, in pieces) and is returned only to the user who saved that quote. The server accepts JPEG data only, up to a size limit, and deletes the picture when the quote is deleted.

## F. Approval workflow (3 levels, by e-mail)

* **Setup:** redeploy `Code.gs` (New version), type the website address into `SITE_URL`, run `setup` once and accept the new permission (sending e-mail). The **Approvals** tab is created automatically.
* **Admin:** Users tab > pick a user > fill *User e-mail* and the **Level 1 / 2 / 3** approver names and e-mails > *Save e-mail & approvers only* (or *Add user / set password*, which saves them together).
* **User:** *Send approval* saves the quote and e-mails Level 1. Each approver opens the link in the e-mail and chooses **Approve** or **Reject** (remarks are compulsory for reject). Approve at Level 1 or 2 e-mails the next level; Approve at Level 3 e-mails everyone. A reject e-mails the remarks to the user who prepared the quote, who changes the quote and clicks *Send approval* again (it starts again at Level 1).
* Each approver link holds a long random one-way token for that level only, so Level 1 cannot open the Level 2 link. A link stops working once the quote is decided or has moved to another level. Anyone who forwards the e-mail forwards the power to decide, so tell approvers not to forward it.
* Gmail accounts can send about 100 approval e-mails a day (Google Workspace about 1,500).
