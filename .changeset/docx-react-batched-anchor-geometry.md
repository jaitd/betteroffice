---
"@betteroffice/docx-react": minor
---

Plugin geometry adds `readAnchorGeometry(targets)`, which resolves once batched worker reads land, and `getPageRect(pageIndex)` with a `built` flag. Answers reaching unbuilt pages list `unbuiltPages`. Worker-only viewers report in-flight targets as `pending` and keep answers across frames of one version.
