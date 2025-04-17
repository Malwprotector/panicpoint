# PanicPoint

**PanicPoint** is a presentation generator written in Rust, designed to help you quickly create presentations (PPTX files) while you're in a rush. Whether you're preparing for a sudden meeting, last-minute presentation, or any other situation where time is of the essence, PanicPoint is here !

## Table of Contents

1. [Features](#features)
2. [Download](#download)
3. [Building](#building)
    - [Prerequisites](#prerequisites)
    - [Clone the Repository](#clone-the-repository)
    - [Build the Project](#build-the-project)
    - [Run the Application](#run-the-application)
    - [Dependencies](#dependencies)
4. [How It Works](#how-tt-works)
5. [Contribution](#contribution)
6. [Reporting Issues](#reporting-issues)

## Features

- **Quick and Easy**: Generate PowerPoint presentations in a few seconds.
- **Rust-Based**: Built in Rust, ensuring speed.
- **Cross-Platform**: Works on Windows, macOS, and Linux.

## Download

_I'm gonna add download link option here soon ! If you want to use PanicPoint you must build it at this time._

## Building

### Prerequisites

Before you install PanicPoint, ensure you have the following installed:

- **Rust** (version 1.60.0 or later): You can install Rust from [https://www.rust-lang.org/](https://www.rust-lang.org/).

### Clone the Repository

Start by cloning the PanicPoint repository:

```bash
git clone https://github.com/Malwprotector/panicpoint.git
cd panicpoint
```

### Build the Project

To build the PanicPoint application, use the following command:

```bash
cargo build --release
```

This will compile the code and generate the executable in the `target/release/` directory.

### Run the Application

Once the build is complete, you can run PanicPoint:

```bash
./target/release/panicpoint
```

## ⚙️ How It Works

**PanicPoint** is written in Rust and focuses on quickly generating `.pptx` (PowerPoint) files from user input directly in the terminal. Here's a breakdown of how the tool works and how Rust features are leveraged:

### 1. 🧠 Input Handling

The application starts by displaying a welcome message and then prompts the user for:

- A **presentation title**
- A series of **slides**, each containing:
  - A **title**
  - Either **paragraph text** or **bullet points**

User input is handled via standard input/output using Rust’s `std::io` library with robust flushing and trimming to ensure clean and intuitive terminal interactions.

### 2. 🧱 Directory & Structure Creation

Once the slides are collected, the program builds a valid PowerPoint file structure in a temporary directory (`temp_pptx`). This involves:

- Creating the required Office Open XML directory structure and files:
  - `_rels/.rels`, `ppt/presentation.xml`, `ppt/_rels/presentation.xml.rels`, `docProps/core.xml`, and others.
  - Separate XML files for each slide.
- Writing well-formed XML files conforming to the `.pptx` specification using `std::fs::File` and `std::io::Write`.

### 3. 🧾 Slide Rendering

For each slide, the program creates:

- A specific `slideN.xml` file, where `N` is the slide number.
- Appropriate content inside the slide:
  - Paragraphs are wrapped in `a:t` tags inside `a:p`.
  - Bullet points use XML tags to represent list items with consistent formatting.

### 4. 🧵 Relationships & References

PowerPoint requires internal file linking via `.rels` files, so:

- The main `presentation.xml` references all slides.
- Each slide has its own `.rels` file to link to shared components like slide layouts or styles.
- A minimal slide master is created to meet the `.pptx` format requirements.

### 5. 📦 Compression into .pptx

The entire folder structure is zipped using the [`zip`](https://docs.rs/zip) crate:

- A `ZipWriter` collects all files in order, preserving their relative paths.
- The `.pptx` file is named using the current timestamp for uniqueness.

```rust
let zip_path = format!("{}.pptx", Local::now().format("%Y-%m-%d_%H-%M-%S"));
```

### 6. 🧹 Cleanup and Feedback

After generating the file:

- The temporary `temp_pptx` directory is deleted.
- The user is notified of the successful creation and file location.

### 🛡️ Error Handling

Rust’s `Result` and the `thiserror` crate are used to handle errors gracefully, providing:

- Custom error messages for:
  - IO errors
  - Zip compression issues
  - Path resolution errors
- Immediate exit with feedback if a critical error occurs.

### Dependencies

PanicPoint relies on the following Rust crates:

- `pptx`: A crate to handle the creation of PowerPoint (.pptx) files.
- `serde`: For serializing and deserializing data.
- `regex`: For pattern matching and string manipulations.

These are automatically installed when you build the project with `cargo build`.

## Contribution

Contributions are welcome! If you have any ideas for improvements or encounter bugs, please feel free to open an issue or submit a pull request. You can contribute to the project by following these steps:

1. Fork the repository.
2. Create a new branch for your feature or fix.
3. Make your changes.
4. Write tests if applicable.
5. Submit a pull request with a description of your changes.

## Reporting Issues

If you encounter any issues or bugs while using PanicPoint, please report them on the [GitHub Issues page](https://github.com/Malwprotector/panicpoint/issues). Provide as much detail as possible, including the version of PanicPoint, operating system, and any error messages you received.