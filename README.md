# Excel Art

Convert a picture to a piece of Excel art.

## But... why ?

I saw [this tweet](https://twitter.com/labnol/status/966637211939635200) earlier and figured it would be easy and fun to replicate.

## Installation

First, start by creating a virtual environment to prevent installing the dependencies globally on your system.

```bash
$ python3 -m venv .venv
```

Activate that virtual environment

```bash
$ source .venv/bin/activate
```

Install the project dependencies

```bash
$ make install
```

## Usage

```bash
$ ./excel-art.py --path=your-image.jpg
````
