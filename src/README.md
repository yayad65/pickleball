# Source for the Pickleball round robin app

- `pickleball.html`: the one source file for the app. Edit this, never the built pages.
- `build_site.py`: run `python3 src/build_site.py` from the repo root. It writes the four pages in the repo root:
  `index.html` (Paper Studio look, the default), `paper.html` (same), `classic.html`, `dink.html`.
- `apps-script/Code.gs`: the Google Apps Script behind the email invites. It lives in the "Get Pickled RSVP" Google Sheet.
  Changing this file does nothing until the script is pasted into the Apps Script editor, saved, and deployed as a new version
  of the existing web app deployment (Deploy > Manage deployments > edit > New version), so the web app URL stays the same.
- `rsvp.html` (repo root): the players' one-tap reply page linked from the emails.

After building, bump `APP_VERSION` in `pickleball.html` so phones show "New version available".
