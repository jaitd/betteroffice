use docx_edit::structured::{ExportOptions, RevisionView};
use docx_edit::*;
use serde_json::{Value, json};

const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape""#;
const PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78, 0xDA, 0x63, 0x64, 0xF8, 0xCF, 0x50,
    0x0F, 0x00, 0x03, 0x86, 0x01, 0x80, 0x5A, 0x34, 0x7D, 0x6B, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45,
    0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
];

fn picture(container: &str, id: usize, properties: &str) -> String {
    let position = if container == "anchor" {
        r#"<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>"#
    } else {
        ""
    };
    let attributes = if container == "anchor" {
        r#" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" relativeHeight="1" simplePos="0""#
    } else {
        ""
    };
    let wrap = if container == "anchor" {
        "<wp:wrapSquare wrapText=\"bothSides\"/>"
    } else {
        ""
    };
    format!(
        r#"<w:r><w:drawing><wp:{container}{attributes}>{position}<wp:extent cx="914400" cy="914400"/>{wrap}<wp:docPr id="{id}" {properties}/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="image.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:{container}></w:drawing></w:r>"#
    )
}

const SHAPE: &str = r#"<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="9" name="Callout" title="Shape title" descr="Shape alt"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#;

fn package() -> Vec<u8> {
    let body = format!(
        r#"<w:p w14:paraId="00000001"><w:r><w:t>a</w:t></w:r>{}<w:r><w:t>b</w:t></w:r>{}</w:p><w:p w14:paraId="00000002">{}{}</w:p><w:p w14:paraId="00000003">{SHAPE}</w:p><w:sectPr/>"#,
        picture(
            "inline",
            1,
            r#"name="{{slot:cover}}" title="Hero title" descr="A harbour at dawn""#
        ),
        picture("inline", 2, ""),
        picture("anchor", 3, r#"name="Picture 3""#),
        picture("inline", 4, r#"title="Only title""#),
    );
    let parts = vec![
        ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec()),
        ("_rels/.rels".to_owned(), format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>"#).into_bytes()),
        ("word/_rels/document.xml.rels".to_owned(), format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="{R}/image" Target="media/image1.png"/></Relationships>"#).into_bytes()),
        ("word/document.xml".to_owned(), format!(r#"<w:document {NS}><w:body>{body}</w:body></w:document>"#).into_bytes()),
        ("word/media/image1.png".to_owned(), PNG.to_vec()),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn open() -> EditingDoc {
    let doc = EditingDoc::new(1837);
    seed_from_docx(&doc, &package()).unwrap();
    doc
}

fn drawing_inlines(value: &Value, found: &mut Vec<Value>) {
    match value {
        Value::Array(values) => values
            .iter()
            .for_each(|value| drawing_inlines(value, found)),
        Value::Object(fields) => {
            if matches!(value["kind"].as_str(), Some("image" | "unsupported"))
                && fields.contains_key("altText")
            {
                found.push(value.clone());
            }
            fields
                .values()
                .for_each(|value| drawing_inlines(value, found));
        }
        _ => {}
    }
}

fn drawing(name: Option<&str>, title: Option<&str>, description: Option<&str>) -> Value {
    json!({ "name": name, "title": title, "description": description })
}

#[test]
fn read_image_atoms_carry_doc_properties() {
    let doc = open();
    let paragraphs = doc
        .read_paragraphs(&ReadParagraphsRequest {
            story: Some("body".into()),
            para_ids: None,
            view: EditTextView::Accepted,
        })
        .unwrap()
        .paragraphs;
    let atoms: Vec<Value> = paragraphs
        .iter()
        .flat_map(|paragraph| &paragraph.atoms)
        .map(|atom| serde_json::to_value(atom).unwrap())
        .collect();
    assert_eq!(
        atoms,
        vec![
            json!({ "offset": 1, "kind": "image", "name": "{{slot:cover}}", "title": "Hero title", "description": "A harbour at dawn" }),
            json!({ "offset": 3, "kind": "image" }),
            json!({ "offset": 0, "kind": "image", "name": "Picture 3" }),
            json!({ "offset": 1, "kind": "image", "title": "Only title" }),
            json!({ "offset": 0, "kind": "other" }),
        ]
    );
}

#[test]
fn export_image_inlines_carry_doc_properties() {
    let doc = open();
    let content = serde_json::to_value(
        doc.export_structured(&ExportOptions::new(RevisionView::Accepted))
            .unwrap()
            .content,
    )
    .unwrap();
    let mut found = Vec::new();
    drawing_inlines(&content, &mut found);
    let inlines: Vec<Value> = found
        .iter()
        .map(|inline| {
            json!({
                "kind": inline["kind"],
                "altText": inline["altText"],
                "drawing": drawing(
                    inline["name"].as_str(),
                    inline["title"].as_str(),
                    inline["description"].as_str(),
                ),
            })
        })
        .collect();
    assert_eq!(
        inlines,
        vec![
            json!({ "kind": "image", "altText": "A harbour at dawn", "drawing": drawing(Some("{{slot:cover}}"), Some("Hero title"), Some("A harbour at dawn")) }),
            json!({ "kind": "image", "altText": null, "drawing": drawing(None, None, None) }),
            json!({ "kind": "image", "altText": null, "drawing": drawing(Some("Picture 3"), None, None) }),
            json!({ "kind": "image", "altText": "Only title", "drawing": drawing(None, Some("Only title"), None) }),
            json!({ "kind": "unsupported", "altText": "Shape alt", "drawing": drawing(Some("Callout"), Some("Shape title"), Some("Shape alt")) }),
        ]
    );
    for inline in &found {
        for key in ["name", "title", "description"] {
            assert!(inline.get(key).is_some(), "{key} missing on {inline}");
        }
    }
}
