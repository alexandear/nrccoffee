# coffee

Apps Script for the NRC. Coffee spreadsheet.
The sources live in `src/`, and [clasp](https://github.com/google/clasp) syncs them with the script bound to that
spreadsheet.

## Setup

1. Turn on the Apps Script API at <https://script.google.com/home/usersettings>.
   Without it every `push` fails with a 403.
2. Install and log in:

   ```sh
   npm install
   npx clasp login
   ```

   Tick two scopes on the consent screen.
   *See, edit, create, and delete only the specific Google Drive files you use with this app* (`drive.file`) lets clasp
   reach the spreadsheet and its script.
   *Create and update Google Apps Script projects* (`script.projects`) covers clone, create, push and pull.
   Watch out for its neighbour: another entry starts the same way and ends in *deployments*, and picking that one makes
   `create-script` fail with "Request had insufficient authentication scopes".

   Leave the rest unticked.
   They cover commands this repo does not use: deployments, web apps, log tailing, API management, and `clasp
   list-scripts`, which wants *See information about your Google Drive files*.
   Running `clasp login` again adds a scope later.

3. Link this repo to the script.

   If the spreadsheet already has one, copy its ID from Extensions -> Apps Script -> Project Settings:

   ```sh
   npx clasp clone-script <SCRIPT_ID> --rootDir src
   ```

   Otherwise create it:

   ```sh
   npx clasp create-script --type sheets \
     --parentId <SPREADSHEET_ID> \
     --rootDir src
   ```

   Either command writes `.clasp.json`, which holds the script ID and is safe to commit.
   The OAuth token stays outside the repo, in `~/.clasprc.json`.

## What the script does

`Code.js` adds a **Кава** menu with one item, *Заповнити з посилань*.
It works on the active sheet: it finds the header row, then for every row holding a MadHeads cart link
(`https://madheadscoffee.com/p/<slug>`) it writes Назва, Обсмажка, Помел and Ціна, грн.
A row pointing at another shop, 3champsroastery among them, is left alone, and so is every column the script does not
own, because cells are written one at a time.
One run fetches at most 60 products.

Columns are found by their header text, so the current layout and the older one with Обсмаження both work.

`MadHeads.js` reads the shop's own data endpoint, `/p/<slug>/__data.json`, and picks the variant whose slug matches the
link.
The price is always the wholesale one.
Drips carry no grind option, so Помел becomes `-`.

## Usage

```sh
npm run push   # local src/ -> spreadsheet
npm run pull   # spreadsheet -> local src/
npm run open   # open the script in the Apps Script IDE
```

`push` replaces the remote files wholesale, so pull first if anyone edited the script in the browser.
