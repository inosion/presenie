use std::fs::File;
use std::io::BufReader;
use std::path::Path;

use ooxmlsdk::parts::presentation_document::PresentationDocument;

pub fn list_slide_details(template: &Path, verbose: bool) {
    println!(":: Slide Layouts for {}", template.display());

    let file = File::open(template).expect("Failed to open template file");
    let reader = BufReader::new(file);
    let pptx = PresentationDocument::new_from_reader(reader).expect("Failed to parse template file");

    if (verbose) {
        println!("{}", pptx.presentation_part.root_element.to_string().unwrap());
    }

    for (i, slide) in pptx.presentation_part.slide_parts.iter().enumerate() {
        println!("    Name: {} - Type: {:?}", slide.r_id, slide.rels_path);
        number_of_objects(slide);
    }

    // print number of objects on the slide

    // for (i, master) in ppt.slide_masters().iter().enumerate() {
    //     println!("  :: Master [{}]", i);
    //     for layout in master.slide_layouts() {
    //         println!("    Name: {} - Type: {:?}", layout.name(), layout.layout_type());
    //     }
    // }
}
