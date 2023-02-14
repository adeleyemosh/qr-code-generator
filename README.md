# Generate QR Codes with Labels

This function generates QR codes with corresponding text labels for a given range of tags.

## Usage

```python
from generate_qr import generate_qr_with_label

generate_qr_with_label(prefix, start, end, step=1)

```
## Arguments

The `generate_qr_with_label` function takes in the following parameters:

* `prefix` (string): The prefix for the tags. This is usually a string that identifies the type of asset the tag belongs to.
* `start` (int): The starting number of the tag range.
* `end` (int): The ending number of the tag range.
* `step` (int, optional): The step size between the tags in the range. Defaults to 1.

## Return Value

The `generate_qr_with_label` function does not return a value. It generates and saves PNG files for each tag in the range.

### Example

Here's an example of how to generate QR codes and labels for tags AEDCBD0010001 through AEDCBD0010010 with a step size of 2:

```python
generate_qr_with_label('AEDCBD00', 10001, 10010, 2)
```

This will generate QR codes and labels for tags AEDCBD0010001, AEDCBD0010003, AEDCBD0010005, AEDCBD0010007, and AEDCBD0010009 and save them as PNG files in the `QR Codes` directory.

## Dependencies

The generate_qr_with_label function requires the following Python packages to be installed:

* `Pillow`
* `qrcode`

These packages can be installed using pip.
