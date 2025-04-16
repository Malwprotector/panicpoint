use std::fs::{self, File};
use std::io::{self, Write, Read};
use std::path::{Path, PathBuf};
use chrono::Local;
use zip::write::FileOptions;
use zip::ZipWriter;
use thiserror::Error;
use walkdir::WalkDir;
use std::process;

#[derive(Error, Debug)]
enum PanicPointError {
    #[error("File IO error")]
    IoError(#[from] std::io::Error),
    #[error("Zip creation error")]
    ZipError(#[from] zip::result::ZipError),
    #[error("Directory walk error")]
    WalkDirError(#[from] walkdir::Error),
    #[error("Path prefix error")]
    StripPrefixError(#[from] std::path::StripPrefixError),
}

fn show_welcome() {
    println!(
        r"
     ____   __   __ _  _  _  ____  ____  ____  _____  ____ 
    (  _ \ / _\ (  ( \( \/ )(  _ \(  __)(  _ \/ ___(  __)
     ) __//    \/    / )  /  ) _ ( ) _)  )   /\___ \ ) _) 
    (__)  \_/\_/\_)__)(__/  (____/(____)(__\_)(____/(____)
    "
    );
    println!("\n🚨 Your emergency PowerPoint generator for last-minute panics! 🚨");
    println!("\nDon't worry! I'll help you create a professional presentation in seconds.");
    println!("Pro Tip: Just press Enter with no text when you're done adding slides.");
}

fn get_presentation_title() -> String {
    println!("\n{}\n", "=".repeat(50));
    println!("📝 Let's start with the basics:");

    loop {
        print!("\nWhat's the title of your presentation? ");
        io::stdout().flush().unwrap();

        let mut title = String::new();
        io::stdin().read_line(&mut title).unwrap();
        let title = title.trim().to_string();

        if !title.is_empty() {
            return title;
        }
        println!("⚠️  The presentation needs a title! Try again.");
    }
}

fn get_slide_data() -> Vec<(String, Vec<String>)> {
    println!("\n{}\n", "=".repeat(50));
    println!("🖼️  Now let's add your slides (press Enter with no text when done)");

    let mut slides = Vec::new();
    let mut slide_number = 1;

    loop {
        println!("\n➕ Slide #{}", slide_number);
        print!("Slide title (or Enter to finish): ");
        io::stdout().flush().unwrap();

        let mut title = String::new();
        io::stdin().read_line(&mut title).unwrap();
        let title = title.trim().to_string();

        if title.is_empty() {
            if slide_number == 1 {
                println!("⚠️  You haven't added any slides yet! Add at least one.");
                continue;
            }
            break;
        }

        println!("\nWhat content should this slide have?");
        println!("1. Paragraph text");
        println!("2. Bullet points");
        print!("Enter choice (1 or 2): ");
        io::stdout().flush().unwrap();

        let mut choice = String::new();
        io::stdin().read_line(&mut choice).unwrap();
        let choice = choice.trim();

        let content = if choice == "1" {
            println!("\nEnter your paragraph text (press Enter when done):");
            let mut paragraph = String::new();
            loop {
                let mut line = String::new();
                io::stdin().read_line(&mut line).unwrap();
                let line = line.trim();

                if line.is_empty() {
                    if paragraph.is_empty() {
                        println!("⚠️  Please add some content!");
                        continue;
                    }
                    break;
                }
                paragraph.push_str(line);
                paragraph.push('\n');
            }
            vec![paragraph.trim().to_string()]
        } else if choice == "2" {
            println!("\nEnter your bullet points (one per line, blank line to finish):");
            let mut bullets = Vec::new();
            loop {
                print!("• ");
                io::stdout().flush().unwrap();
                let mut line = String::new();
                io::stdin().read_line(&mut line).unwrap();
                let line = line.trim().to_string();

                if line.is_empty() {
                    if bullets.is_empty() {
                        println!("⚠️  Please add at least one bullet point!");
                        continue;
                    }
                    break;
                }
                bullets.push(line);
            }
            bullets
        } else {
            println!("⚠️  Invalid choice! Please try again.");
            continue;
        };

        slides.push((title, content));
        slide_number += 1;
        println!("✅ Slide added! ({} slides total)", slides.len());
    }

    slides
}

fn create_presentation(title: &str, slides: &[(String, Vec<String>)]) -> Result<(), PanicPointError> {
    println!("\n{}\n", "=".repeat(50));
    println!("⚡ Creating your presentation... Almost there!");

    // Create a temp directory
    let dir = "temp_pptx";
    fs::create_dir_all(dir)?;

    // Create necessary directory structure
    fs::create_dir_all(format!("{}/_rels", dir))?;
    fs::create_dir_all(format!("{}/docProps", dir))?;
    fs::create_dir_all(format!("{}/ppt/_rels", dir))?;
    fs::create_dir_all(format!("{}/ppt/slides/_rels", dir))?;
    fs::create_dir_all(format!("{}/ppt/slides", dir))?;
    fs::create_dir_all(format!("{}/ppt/theme", dir))?;
    fs::create_dir_all(format!("{}/ppt/slideLayouts", dir))?;
    fs::create_dir_all(format!("{}/ppt/slideMasters", dir))?;
    fs::create_dir_all(format!("{}/ppt/slideMasters/_rels", dir))?;

    // 1. Create [Content_Types].xml
    let mut content_types = File::create(format!("{}/[Content_Types].xml", dir))?;
    writeln!(content_types, r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#)?;
    writeln!(content_types, r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#)?;
    writeln!(content_types, r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#)?;
    writeln!(content_types, r#"<Default Extension="xml" ContentType="application/xml"/>"#)?;
    writeln!(content_types, r#"<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>"#)?;
    writeln!(content_types, r#"<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>"#)?;
    for i in 0..slides.len() {
        writeln!(content_types, r#"<Override PartName="/ppt/slides/slide{}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#, i + 1)?;
    }
    writeln!(content_types, r#"</Types>"#)?;

    // 2. Create _rels/.rels
    let mut rels = File::create(format!("{}/_rels/.rels", dir))?;
    writeln!(rels, r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#)?;
    writeln!(rels, r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#)?;
    writeln!(rels, r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>"#)?;
    writeln!(rels, r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>"#)?;
    writeln!(rels, r#"</Relationships>"#)?;

    // 3. Create ppt/presentation.xml
    let mut presentation = File::create(format!("{}/ppt/presentation.xml", dir))?;
    writeln!(presentation, r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#)?;
    writeln!(presentation, r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#)?;
    writeln!(presentation, r#"<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>"#)?;
    writeln!(presentation, r#"<p:sldIdLst>"#)?;
    for i in 0..slides.len() {
        writeln!(presentation, r#"<p:sldId id="{}" r:id="rId{}"/>"#, 256 + i, i + 2)?;
    }
    writeln!(presentation, r#"</p:sldIdLst>"#)?;
    writeln!(presentation, r#"<p:sldSz cx="9144000" cy="6858000"/>"#)?; // Slide size (16:9 aspect ratio)
    writeln!(presentation, r#"</p:presentation>"#)?;

    // 4. Create ppt/_rels/presentation.xml.rels
    let mut pres_rels = File::create(format!("{}/ppt/_rels/presentation.xml.rels", dir))?;
    writeln!(pres_rels, r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#)?;
    writeln!(pres_rels, r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#)?;
    writeln!(pres_rels, r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>"#)?;
    for i in 0..slides.len() {
        writeln!(pres_rels, r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{}.xml"/>"#, i + 2, i + 1)?;
    }
    writeln!(pres_rels, r#"</Relationships>"#)?;

    // 5. Create slide master (simplified)
    let mut slide_master = File::create(format!("{}/ppt/slideMasters/slideMaster1.xml", dir))?;
    writeln!(slide_master, r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#)?;
    writeln!(slide_master, r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#)?;
    writeln!(slide_master, r#"<p:cSld>"#)?;
    writeln!(slide_master, r#"<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:bgPr></p:bg>"#)?;
    writeln!(slide_master, r#"<p:spTree>"#)?;
    writeln!(slide_master, r#"<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"#)?;
    writeln!(slide_master, r#"<p:grpSpPr><a:xfrm/></p:grpSpPr>"#)?;
    writeln!(slide_master, r#"</p:spTree>"#)?;
    writeln!(slide_master, r#"</p:cSld>"#)?;
    writeln!(slide_master, r#"</p:sldMaster>"#)?;

    // 6. Create slide master relationships
    let mut slide_master_rels = File::create(format!("{}/ppt/slideMasters/_rels/slideMaster1.xml.rels", dir))?;
    writeln!(slide_master_rels, r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#)?;
    writeln!(slide_master_rels, r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#)?;
    writeln!(slide_master_rels, r#"</Relationships>"#)?;

    // 7. Create slides with content
    for (i, (slide_title, content)) in slides.iter().enumerate() {
        let mut slide = File::create(format!("{}/ppt/slides/slide{}.xml", dir, i + 1))?;
        writeln!(slide, r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#)?;
        writeln!(slide, r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#)?;
        writeln!(slide, r#"<p:cSld>"#)?;
        writeln!(slide, r#"<p:spTree>"#)?;
        
        // Title text box
        writeln!(slide, r#"<p:sp>"#)?;
        writeln!(slide, r#"<p:nvSpPr><p:cNvPr id="1" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>"#)?;
        writeln!(slide, r#"<p:spPr><a:xfrm><a:off x="914400" y="457200"/><a:ext cx="7315200" cy="457200"/></a:xfrm></p:spPr>"#)?;
        writeln!(slide, r#"<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>{}</a:t></a:r></a:p></p:txBody>"#, slide_title)?;
        writeln!(slide, r#"</p:sp>"#)?;
        
        // Content text box
        writeln!(slide, r#"<p:sp>"#)?;
        writeln!(slide, r#"<p:nvSpPr><p:cNvPr id="2" name="Content"/><p:cNvSpPr/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>"#)?;
        writeln!(slide, r#"<p:spPr><a:xfrm><a:off x="914400" y="1143000"/><a:ext cx="7315200" cy="3657600"/></a:xfrm></p:spPr>"#)?;
        writeln!(slide, r#"<p:txBody><a:bodyPr/><a:lstStyle/>"#)?;
        
        if content.len() == 1 {
            // Single paragraph
            writeln!(slide, r#"<a:p><a:r><a:rPr lang="en-US"/><a:t>{}</a:t></a:r></a:p>"#, content[0])?;
        } else {
            // Bullet points
            for item in content {
                writeln!(slide, r#"<a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US"/><a:t>{}</a:t></a:r></a:p>"#, item)?;
            }
        }
        
        writeln!(slide, r#"</p:txBody>"#)?;
        writeln!(slide, r#"</p:sp>"#)?;
        
        writeln!(slide, r#"</p:spTree>"#)?;
        writeln!(slide, r#"</p:cSld>"#)?;
        writeln!(slide, r#"</p:sld>"#)?;

        // Slide relationships
        let mut slide_rels = File::create(format!("{}/ppt/slides/_rels/slide{}.xml.rels", dir, i + 1))?;
        writeln!(slide_rels, r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#)?;
        writeln!(slide_rels, r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#)?;
        writeln!(slide_rels, r#"</Relationships>"#)?;
    }

    // 8. Create docProps/core.xml
    let mut core = File::create(format!("{}/docProps/core.xml", dir))?;
    writeln!(core, r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#)?;
    writeln!(core, r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">"#)?;
    writeln!(core, r#"<dc:title>{}</dc:title>"#, title)?;
    writeln!(core, r#"<dc:creator>PanicPoint</dc:creator>"#)?;
    writeln!(core, r#"<cp:lastModifiedBy>PanicPoint</cp:lastModifiedBy>"#)?;
    writeln!(core, r#"<dcterms:created xsi:type="dcterms:W3CDTF">{}</dcterms:created>"#, Local::now().to_rfc3339())?;
    writeln!(core, r#"<dcterms:modified xsi:type="dcterms:W3CDTF">{}</dcterms:modified>"#, Local::now().to_rfc3339())?;
    writeln!(core, r#"</cp:coreProperties>"#)?;

    // Package everything into a PPTX file
    let filename = format!("PanicPoint_{}.pptx", title.replace(' ', "_"));
    package_ppt_directory(dir, &filename)?;

    // Clean up temp directory
    fs::remove_dir_all(dir)?;

    println!("\n🎉 Success! Presentation saved as: {}", filename);
    Ok(())
}

fn package_ppt_directory(dir: &str, output_file: &str) -> Result<(), PanicPointError> {
    let file = File::create(output_file)?;
    let mut zip = ZipWriter::new(file);
    let options = FileOptions::default()
        .compression_method(zip::CompressionMethod::Stored);

    for entry in WalkDir::new(dir) {
        let entry = entry?;
        let path = entry.path();
        if path.is_file() {
            let relative_path = path.strip_prefix(dir)?;
            let relative_path = relative_path.to_str().ok_or_else(|| {
                std::io::Error::new(std::io::ErrorKind::InvalidData, "Invalid path encoding")
            })?;
            
            zip.start_file(relative_path, options)?;
            let mut file = File::open(path)?;
            io::copy(&mut file, &mut zip)?;
        }
    }

    zip.finish()?;
    Ok(())
}

fn main() {
    show_welcome();
    let title = get_presentation_title();
    let slides = get_slide_data();
    
    if let Err(e) = create_presentation(&title, &slides) {
        eprintln!("\n😱 Error: {}", e);
        process::exit(1);
    }
}
