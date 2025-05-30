# comintext2csv
A utility to convert COMINTEXT formatted files into CSV format using a reference Excel file.

---

## Usage

```bash
./comintext2csv data [-r reference.xls] [-o output.csv]
```

- `data`: A `.txt` file or a folder with `.txt` files
- `-r`: (Optional) Excel file that defines layout.
If omitted, the program will try to find an `.xls` or `.xlsx` file with a similar name in the same directory.
 - `-o`: (Optional) Output CSV file path (only for single file mode).
 If omitted, the program will use the base name of the data file.

---

## Examples

Convert one file with a reference layout
```bash
./comintext2csv ./legdat.txt -r LEGDAT\ Extract\ Data\ Format.xls -o result.csv
```

Convert one file and let the script find the layout automatically
```bash
./comintext2csv ./legdat.txt
```

Convert all `.txt` files in a folder. The resulting CSV files will be placed in your working directory.
```bash
./comintext2csv folder/
```

---

## macOS Binary

A prebuilt macOS binary is provided. It must be kept in the same directory as the \_internal directory. To mark it as executable, run the following command:
```bash
chmod +x dist/comintext2csv
```

If you want to make this command available in your terminal, run the following command:
```bash
echo 'export PATH="$PATH:'"$(realpath dist)"'"' >> ~/.profile
source ~/.profile
```
If you move your directory, this will break.

---

## Windows

You will have to run the `.py` script if you are using Windows.

1. Ensure you have python3 installed
2. Install the dependencies with `pip install pandas openpyxl`
3. Run the script with `python comintext2csv.py ...`
