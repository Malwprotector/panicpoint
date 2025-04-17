# PanicPoint

![https://raw.githubusercontent.com/Malwprotector/panicpoint/refs/heads/main/screenshot.png](https://raw.githubusercontent.com/Malwprotector/panicpoint/refs/heads/main/screenshot.png)

**PanicPoint** is designed to help you quickly create presentations (PPTX files) while you're in a rush. Whether you're preparing for a sudden meeting, last-minute presentation, or any other situation where time is of the essence, PanicPoint is here !

This project enabled me to understand how to manage XML in Rust.

## Table of Contents

1. [Features](#features)
2. [Download](#download)
3. [Building](#building)
    - [Prerequisites](#prerequisites)
    - [Clone the Repository](#clone-the-repository)
    - [Build the Project](#build-the-project)
    - [Run the Application](#run-the-application)
    - [Dependencies](#dependencies)
4. [How It Works](#how-it-works)
5. [Contribution](#contribution)
6. [Reporting Issues](#reporting-issues)

## Features

- **Quick and Easy**: Generate PowerPoint presentations in a few seconds.
- **Rust-Based**: Built in Rust, ensuring speed.
- **Cross-Platform**: Works on Windows, macOS, and Linux.

## Download

Choose your platform to download the correct build:

- 🪟 **[Download for Windows (EXE)](https://github.com/Malwprotector/panicpoint/raw/refs/heads/main/target/x86_64-pc-windows-gnu/release/panicpoint.exe)**
- 🐧 **[Download for Linux](https://github.com/Malwprotector/panicpoint/raw/refs/heads/main/target/release/panicpoint)**

Make sure to mark the file as executable on Linux:

```bash
chmod +x panicpoint
```

## Building

### Prerequisites

- **This is the steps to build PanicPoint on Linux.** It's pretty similar on windows, but some things may be different : please refer to [Rust docs.](https://doc.rust-lang.org/stable/)

Before you install PanicPoint, ensure you have the following installed:

- **Rust** (version 1.80.0 or later): You can install Rust from [https://www.rust-lang.org/](https://www.rust-lang.org/).

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

## How It Works

Here's a breakdown of how PanicPoint works and how Rust features are leveraged:

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
- The `.pptx` file is created.

### 6. 🧹 Cleanup and Feedback

After generating the file:

- The temporary `temp_pptx` directory is deleted.
- The user is notified of the successful creation and file location.

### 🛡️ Error Handling

Rust’s `Result` and the `thiserror` crate are used to handle errors gracefully.

### Dependencies

PanicPoint relies on the following Rust crates:
```
chrono = "0.4"
zip = "0.6"
thiserror = "1.0"
walkdir = "2.3"
```
These are automatically installed when you build the project with `cargo build`.

## Contribution

Contributions are welcome! If you have any ideas for improvements or encounter bugs, please feel free to open an issue or submit a pull request.

## Reporting Issues

If you encounter any issues or bugs while using PanicPoint, please report them on the [GitHub Issues page](https://github.com/Malwprotector/panicpoint/issues).