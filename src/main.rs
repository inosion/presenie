use clap::{Parser, Subcommand};
use std::path::Path;

mod tools;

const VERSION: &str = "1.0.0";

#[derive(Parser)]
#[clap(name = "Presenie", version = VERSION, author = "Inosion (c) 2024", about = "Presenie CLI tool")]
struct Cli {
    #[clap(short, long, help = "Sets the level of verbosity")]
    verbose: bool,

    #[clap(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
    #[clap(about = "List slide layouts")]
    List {
        #[clap(help = "The template to merge data with")]
        template: String,
    },
    #[clap(about = "Merge data with template")]
    Merge {
        #[clap(help = "The template to merge data with")]
        template: String,
        #[clap(help = "The data file to merge with the template")]
        data: String,
        #[clap(help = "Config file to drive merging")]
        config: Option<String>,
        #[clap(help = "Output filename")]
        out_file: String,
    },
    #[clap(about = "Clone slides")]
    Clone {
        #[clap(help = "The template to merge data with")]
        template: String,
        #[clap(help = "Output filename")]
        out_file: String,
    },
}

fn main() {
    let cli = Cli::parse();

    match &cli.command {
        Commands::List { template } => {
            crate::tools::list_slide_details(Path::new(template), cli.verbose);
        }
        Commands::Merge {
            template,
            data,
            config: _,
            out_file,
        } => {
            pptx_merger::render(Path::new(data), Path::new(template), Path::new(out_file));
            println!("Wrote {}", Path::new(out_file).display());
        }
        Commands::Clone { template, out_file } => {
            println!(
                "Cloned {} --> {}",
                Path::new(template).display(),
                Path::new(out_file).display()
            );
            pptx_tools::clone_ppt_slides(Path::new(template), Path::new(out_file));
        }
    }
}

mod pptx_tools {
    use std::path::Path;

    pub fn clone_ppt_slides(template: &Path, out_file: &Path) {
        // Implement the function to clone PPT slides
    }
}

mod pptx_merger {
    use std::path::Path;

    pub fn render(data: &Path, template: &Path, out_file: &Path) {
        // Implement the function to merge data with template
    }
}