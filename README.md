# Wilmington Brewing Runner Portal

Free Google Apps Script + Google Sheets app for Wilmington Brewing Co. Run Club.

## What v1 includes
- Member registration page
- Required waiver agreement
- Automatic WBRC Runner IDs
- Fist Bumps rewards currency
- QR code generation
- Welcome email
- Member dashboard lookup
- Volunteer check-in page
- Admin event creation and member search
- Google Sheets backend

## Install
1. Open script.google.com and create a new Apps Script project.
2. Create files matching this repo:
   - Code.gs
   - Index.html
   - Styles.html
   - JavaScript.html
3. Paste each file's contents.
4. Run `setupPortal` once.
5. Deploy as Web App:
   - Execute as: Me
   - Who has access: Anyone with the link
6. Copy the Web App URL.
7. Run `setWebAppUrlFromPrompt` and paste the Web App URL when prompted.

## First admin actions
1. Open the web app URL.
2. Go to Admin.
3. Create an active event.
4. Register a test member.
5. Use Volunteer Check-In to enter/scan the Runner ID.
