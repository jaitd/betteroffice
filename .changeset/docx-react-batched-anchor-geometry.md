---
"@betteroffice/docx-react": minor
---

Plugin geometry adds `readAnchorGeometries(targets)`, which resolves once batched worker reads land, `getPageRect(pageIndex)` and `isPageBuilt(pageIndex)`. Answers reaching unbuilt pages list `unbuiltPages`. Worker-only viewers report in-flight targets as `pending` and failed reads as `read-failed`.
