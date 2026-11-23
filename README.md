# Survey Finder

A searchable catalogue of questionnaires and codebooks from our research projects. People enter an independent variable (IV) and a dependent variable (DV) and see which surveys measure both. Only questionnaires are listed here; the data always stays with its owners.

## What's in this folder

| File | What it is |
|---|---|
| `index.html` | The website (search page). |
| `data/surveys.csv` | One row per survey: title, countries, fieldwork, sample, access status, data link, citation, contact. |
| `data/variables.csv` | One row per question (or question battery): variable name, English summary, original wording, scale, concepts. |
| `data/concepts.csv` | The concepts questions are tagged with (e.g. "Gender attitudes & sexism"), the themes they belong to, and search synonyms in English/Spanish. |
| `data/themes.csv` | The broad themes used for browsing (e.g. "Socio-economic status"). |
| `data/preview-data.js` | A copy of the CSVs so `index.html` also opens directly from your disk. Not needed online. |
| `tools/` | Scripts used to build the first version of the tables. |

## Adding a survey (manual workflow)

1. Someone sends you the questionnaire and the completed contribution form.
2. Add one row to `data/surveys.csv` with the form answers. Use a short, lowercase `id` (e.g. `ess-spain-2025`).
3. Add one row per question or battery to `data/variables.csv`, using that `id` in `survey_id`. In `concepts`, list one or more concept ids from `concepts.csv`, separated by `;`.
4. If a needed concept doesn't exist, add it to `concepts.csv` (id, label, theme ids, synonyms).
5. Upload the changed files to the GitHub repository. The site updates in a minute or two.

Steps 2–4 can be done by Claude from the questionnaire file: send it and ask for the new rows.

**Editing in Excel:** open the CSV, edit, then *Save As → CSV UTF-8 (comma delimited)*. Plain "CSV" breaks accents.

## Field conventions

- `status`: `open`, `request` or `embargoed`.
- `open_to_collaboration`: `yes`, `no` or leave empty.
- `countries`: separated by `;`, in English (e.g. `Spain;Portugal`).
- `questionnaire_file`: leave empty unless the owner agrees to the questionnaire being downloadable from the site. Then put the file in `questionnaires/` and write e.g. `questionnaires/mysurvey.pdf`.
- Anything unknown: write `To confirm`. It shows in grey italics.

## Putting it online (GitHub Pages)

1. Create a GitHub account and a new **public** repository (e.g. `survey-finder`).
2. Upload `index.html` and the `data` folder (*Add file → Upload files*).
3. *Settings → Pages → Branch: main / root → Save*. The site will be at `https://<username>.github.io/survey-finder/`.

Before going public, check that every owner is happy for their survey's questions to be listed publicly.
