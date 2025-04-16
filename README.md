# PanicPoint

**PanicPoint** is a presentation generator written in Rust, designed to help you quickly create PowerPoint presentations (PPTX files) while you're in a rush. Whether you're preparing for a sudden meeting, last-minute presentation, or any other situation where time is of the essence, PanicPoint is here !

## Table of Contents

1. [Features](#features)
2. [Download](#download)
3. [Building](#building)
    - [Prerequisites](#prerequisites)
    - [Clone the Repository](#clone-the-repository)
    - [Build the Project](#build-the-project)
    - [Run the Application](#run-the-application)
    - [Dependencies](#dependencies)
4. [Contribution](#contribution)
5. [Reporting Issues](#reporting-issues)

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
git clone https://github.com/your-username/panicpoint.git
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