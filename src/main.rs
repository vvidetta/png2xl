// png2xl
//
// Program to convert png file into excel spreadsheet with formatting
// Inspired by https://youtu.be/UBX2QQHlQ_I
//

use std::{env, path::{PathBuf}, fs::File};
use xlsxwriter::{self, FormatColor};

struct PngRasterData {
    width: u16,
    height: u32,
    buffer: Vec<u8>
}

struct Position {
    x : u16,
    y : u32
}

struct Color {
    r: u8,
    g: u8,
    b: u8
}

struct Pixel {
    pos : Position,
    color : Color
}

struct ExcelRasterData {
    buffer: Vec<Vec<Pixel>>
}

fn read_png(png_file : &str) -> PngRasterData {
    println!("READING {} ...", png_file);
    let file = File::open(png_file).unwrap();
    let decoder = png::Decoder::new(file);
    let mut reader = decoder.read_info().unwrap();
    println!("Bytes per pixel: {}", reader.info().bytes_per_pixel() as i32);

    let width = reader.info().width as u16;
    let height = reader.info().height as u32;
    let frame_size = (width * height) as u32;
    let buffer_size = frame_size * 3;
    let mut png_raster_data = PngRasterData{
        width : width,
        height : height,
        buffer : vec![0u8; (width as u32 * height * 3) as usize]
    };

    println!("Interlaced: {}", reader.info().interlaced);

    if !reader.info().interlaced {
        for i in 0..4 {
            if i == 0 {
                continue;
            } else if i == 1 {
                reader.next_frame(&mut png_raster_data.buffer[0..frame_size]);
            } else if i == 2 {
                reader.next_frame(&mut png_raster_data.buffer[(framesize)..(2*frame_size)]);
            } else if i == 3 {
                reader.next_frame(&mut png_raster_data.buffer[(2*frame_size)..(3*frame_size)]);
            }
           if i == 1 {

           }
        }
    }

    png_raster_data
    // PngRasterData{
    //     width : 2,
    //     height : 2,
    //     buffer : vec![
    //         0u8, 0u8, 0u8,
    //         0u8, 85u8, 85u8,
    //         0u8, 170u8, 170u8,
    //         0u8, 255u8, 255u8
    //     ]
    // }
}

fn transform_png_to_excel(png_raster_data: PngRasterData) -> ExcelRasterData {
    println!("TRANSFORMING png to excel ...");
    let mut i = 0;
    let mut pixel_buffer : Vec<Vec<Pixel>> = vec![];
    for y in 0..png_raster_data.height {
        let mut row : Vec<Pixel> = vec![];
        for x in 0..png_raster_data.width {
            row.push(Pixel{
                pos : Position{ x : x, y : y },
                color : Color{
                    r : png_raster_data.buffer[i],
                    g : png_raster_data.buffer[i + 1],
                    b : png_raster_data.buffer[i + 2]
                }
            });
            i += 3;
        }
        pixel_buffer.push(row);
    }
    ExcelRasterData{
        buffer : pixel_buffer
    }
}

fn write_pixel(workbook : &mut xlsxwriter::Workbook, sheet_name : &str, pixel: &Pixel) {
    let row = pixel.pos.y * 2;
    let column = pixel.pos.x * 2 ;
    let mut worksheet = workbook.get_worksheet(sheet_name).unwrap();

    worksheet.set_column(column, column + 1, 1.0, None);
    worksheet.set_row(row, 10.0, None);
    worksheet.set_row(row + 1, 10.0, None);

    let r_format = workbook.add_format().set_bg_color(FormatColor::Custom((pixel.color.r as u32) << 16));

    worksheet.write_number(row, column, pixel.color.r.into(), Some(&r_format));

    let g_format = workbook.add_format().set_bg_color(FormatColor::Custom((pixel.color.g as u32) << 8));
    worksheet.write_number(row, column + 1, pixel.color.g.into(), Some(&g_format));
    worksheet.write_number(row + 1, column, pixel.color.g.into(), Some(&g_format));

    let b_format = workbook.add_format().set_bg_color(FormatColor::Custom(pixel.color.b as u32));
    worksheet.write_number(row + 1, column + 1, pixel.color.b.into(), Some(&b_format));
}

fn write_excel(excel_file : PathBuf, excel_raster_data : ExcelRasterData) -> () {
    println!("WRITING {} ...", excel_file.display());
    let mut workbook = xlsxwriter::Workbook::new(excel_file.to_str().unwrap());
    let sheet_name = "Image";
    let _ = workbook.add_worksheet(Some(sheet_name));

    for row in excel_raster_data.buffer {
        for pixel in row {
            write_pixel(&mut workbook, sheet_name, &pixel)
        }
    }

    let _ = workbook.close();
}

fn output_filename(args: &Vec<String>) -> PathBuf {
    if args.len() >= 3 { 
        PathBuf::from(args[2].clone())
    } else {
        let mut out_file = PathBuf::from(args[1].clone());
        out_file.set_extension("xlsx");
        out_file
    }
}

fn main() -> std::process::ExitCode {
    let args : Vec<String> = env::args().collect();
    if args.len() < 2 {
        println!("Usage: {} <path to png> [<path to output xlsx>]", args[0]);
        return std::process::ExitCode::FAILURE;
    }
    let png_file = &args[1];
    let png_raster_data = read_png(png_file);
    let excel_raster_data = transform_png_to_excel(png_raster_data);
    let excel_file = output_filename(&args);
    write_excel(excel_file, excel_raster_data);
    std::process::ExitCode::SUCCESS
}
