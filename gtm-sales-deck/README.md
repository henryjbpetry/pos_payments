# GTM investment deck (static HTML)

Single-page “slide” deck: full-height sections, anchor navigation, dark theme. Edit `index.html` and push to GitHub.

## Publish on GitHub Pages

1. Create a new repository on GitHub (any name, e.g. `gtm-sales-deck`).
2. Upload this folder’s contents to the repo root (`index.html`, `.nojekyll`, this README), or push from your machine / Cursor.
3. In the repo: **Settings → Pages → Build and deployment → Source: Deploy from a branch**.
4. Choose branch **main** (or **master**) and folder **/ (root)**, then save.
5. After the workflow runs, the site will be at:

   `https://<your-username>.github.io/<repo-name>/`

If the repository is named `<username>.github.io`, the site root is `https://<username>.github.io/` with files in that repo’s default branch.

## Edit content

- Change the `<title>`, nav labels, and each `<section class="slide" id="slide-N">` block in `index.html`.
- Replace placeholder numbers and tables with your model.
- Optional: add exported PNG/SVG charts under an `assets/` folder and use `<img>` tags on slide 4 or elsewhere.

## Cursor + GitHub

Connecting GitHub in Cursor lets you use the Source Control view to commit and push. This template does not run any install step; it is static HTML/CSS only.
