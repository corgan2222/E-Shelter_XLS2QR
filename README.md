# E-Shelter Excel2QR

 Reads Data from given Excel File and exports Single PDFs and a complete PDF grouped by Gateway
 
 ![grafik](https://user-images.githubusercontent.com/12233951/141602707-1fcbabad-8489-4a6a-b5d7-d8b7c5be71d1.png)
 
<br>

![grafik](https://user-images.githubusercontent.com/12233951/141602720-b96a8bfd-4e37-4f38-a37a-75974541edd8.png)

<br>

## Features

- Reads Excel 2021
- Export Single Labels as PDF and SVG
- Export PDF with all rows grouped by Gateway

<br>

## Folder Structure

- Data<br>
   \  xlsx           -   Insert Excel Files here <br>
    \ output         -   Storage Folder for the combined PDFs<br>
    \ output_pdf     -   Storage Folder for the PDFs grouped by Gateway<br>
    \ output_svg     -   Storage Folder for the SVGs grouped by Gateway<br>

<br>

## Informations

### Workflow
- put your Excel Assetlist into data/xlsx Folder
- start

# Windows executable

- Download latest Release Version from https://github.com/corgan2222/E-Shelter_XLS2QR/releases/ 
- Extract Zip File
- Start exe

# Source Installation Linux

```shell
git clone https://github.com/corgan2222/E-Shelter_XLS2QR.git
cd E-Shelter_XLS2QR
./install_linux.sh
python3 -m pip install -r requirements.txt 
```
<br>

There are two running modes.

## 1. Folder Based

```shell
python3 xls2qr.py
```

- just run the script with no parameter.
- The script will look in the data/xlsx Folder for an Excel File

<br>

## 2. File based

<br>You can define all files and Folder via command line parameter

```shell
python3 xls2qr.py -x [excelfile] 
```

example Windows

```powershell
python3 xls2qr.py -x "c:\tmp\file.xlsx" 
```

example Linux

```shell
python3 xls2qr.py -i "/home/user/files/file.xlsx" 
```
<br>
<br>

# Command line Settings

| Option | Description | Notes |
| --- | ----------- | ---------- |
| -x | Excelfile | "/home/user/files/file.xlsx"
| -o | Output Folder | "/home/user/files/data/output"
| -l | Logfile | "/car/log/debug.log"
| -v | verbose | show debug logs if set
| -q | quit | no output


<br>
<br>
