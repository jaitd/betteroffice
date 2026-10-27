---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/python-docx": patch
"@betteroffice/rust-crates": patch
---

Image atoms in `readParagraphs` and `readStories`, and `image` and drawing `unsupported` inlines in the structured export, now carry the picture's `name`, `title` and `description` from `wp:docPr`. `altText` is unchanged.
