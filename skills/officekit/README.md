# OfficeKit Skill

OfficeKit is the coordination entry point for broad or cross-format Office
work. It turns a request into an explicit artifact route, loads the required
Documents, Spreadsheets, Presentations, or PDF Skill, and decides whether zero
or one available template helps.

Install the coordinated core Skills together:

```sh
npx skills add w31r4/open-office-artifact-tool \
  --skill officekit documents spreadsheets excel-live-control presentations pdf template-creator \
  --yes
```

The installer only deploys Skill instructions and resources. OfficeKit does
not replace the format-specific workflows or provide a second file codec.

Repository templates and locally created `artifact-template-*` Skills are
queried through compact metadata. Individual template Skills remain directly
installable for people who want explicit template invocation. OfficeKit loads a
template's full instructions only after that template has been selected.
