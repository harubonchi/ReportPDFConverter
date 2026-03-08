use std::error::Error;
use std::fs::File;
use std::path::{Path, PathBuf};

fn main() -> Result<(), Box<dyn Error>> {
    println!("cargo:rerun-if-changed=static/favicon.svg");

    let manifest_dir = PathBuf::from(std::env::var("CARGO_MANIFEST_DIR")?);
    let svg_path = manifest_dir.join("static").join("favicon.svg");
    let static_ico_path = manifest_dir.join("static").join("favicon.ico");
    let out_dir = PathBuf::from(std::env::var("OUT_DIR")?);
    let ico_path = out_dir.join("app.ico");

    generate_ico_from_svg(&svg_path, &ico_path)?;
    generate_ico_from_svg(&svg_path, &static_ico_path)?;

    if std::env::var("CARGO_CFG_WINDOWS").is_ok() {
        let mut res = winres::WindowsResource::new();
        res.set_icon(ico_path.to_string_lossy().as_ref());
        res.compile()?;
    }

    Ok(())
}

fn generate_ico_from_svg(svg_path: &Path, ico_path: &Path) -> Result<(), Box<dyn Error>> {
    let svg_data = std::fs::read(svg_path)?;
    let options = usvg::Options::default();
    let tree = usvg::Tree::from_data(&svg_data, &options)?;

    let width: u32 = 256;
    let height: u32 = 256;

    let mut pixmap =
        resvg::tiny_skia::Pixmap::new(width, height).ok_or("failed to create pixmap")?;

    let svg_size = tree.size();
    let scale_x = width as f32 / svg_size.width();
    let scale_y = height as f32 / svg_size.height();
    let transform = resvg::tiny_skia::Transform::from_scale(scale_x, scale_y);

    let mut pixmap_mut = pixmap.as_mut();
    resvg::render(&tree, transform, &mut pixmap_mut);

    let image = ico::IconImage::from_rgba_data(width, height, pixmap.take());
    let mut icon_dir = ico::IconDir::new(ico::ResourceType::Icon);
    icon_dir.add_entry(ico::IconDirEntry::encode(&image)?);

    let mut file = File::create(ico_path)?;
    icon_dir.write(&mut file)?;

    Ok(())
}
