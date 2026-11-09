
import pandas as pd
import requests
from pathlib import Path
from io import BytesIO
from PIL import Image
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import time

# ==========================================
# 1. SETTINGS
# ==========================================

CSV_FILE = "galaxies.csv"
OUTPUT_FOLDER = Path("galaxy_images")
NUMBER_OF_GALAXIES = 10

IMAGE_WIDTH = 256
IMAGE_HEIGHT = 256
IMAGE_SCALE = 0.4

IMAGE_API = (
    "https://skyserver.sdss.org/dr20/"
    "SkyServerWS/ImgCutout/getjpeg"
)

OUTPUT_FOLDER.mkdir(exist_ok=True)

# ==========================================
# 2. LOAD ORIGINAL SDSS DATA
# ==========================================

print("Loading SDSS galaxy data...")

df = pd.read_csv(
    CSV_FILE,
    dtype={"objid": str}
)

df.columns = df.columns.str.strip()

required = ["objid", "ra", "dec"]

for col in required:
    if col not in df.columns:
        raise ValueError(
            f"Missing column: {col}"
        )

df = df.head(NUMBER_OF_GALAXIES).copy()

print(f"Loaded {len(df)} galaxies.")

# ==========================================
# 3. DOWNLOAD GALAXY IMAGES
# ==========================================

session = requests.Session()
image_paths = []

for i, galaxy in df.iterrows():

    objid = str(galaxy["objid"])
    ra = float(galaxy["ra"])
    dec = float(galaxy["dec"])

    print(f"\nGalaxy {len(image_paths) + 1}/{len(df)}")
    print(f"Object ID: {objid}")

    params = {
        "ra": ra,
        "dec": dec,
        "scale": IMAGE_SCALE,
        "width": IMAGE_WIDTH,
        "height": IMAGE_HEIGHT
    }

    filepath = OUTPUT_FOLDER / f"{objid}.jpg"

    try:
        response = session.get(
            IMAGE_API,
            params=params,
            timeout=30
        )

        response.raise_for_status()

        image = Image.open(
            BytesIO(response.content)
        ).convert("RGB")

        image.save(filepath, quality=95)

        image_paths.append(str(filepath))

        print("Image downloaded successfully!")

    except Exception as e:
        print(f"Download failed: {e}")
        image_paths.append("")

    time.sleep(0.5)

# ==========================================
# 4. SAVE COMPLETE CSV
# ==========================================

df["image_file"] = image_paths

df.to_csv(
    "galaxies_with_images.csv",
    index=False
)

print("\nCSV file created.")

# ==========================================
# 5. CREATE EXCEL WORKBOOK
# ==========================================

print("Creating Excel workbook...")

wb = Workbook()
ws = wb.active
ws.title = "Galaxy Data"

columns = list(df.columns)

# Preserve large Object IDs as text
ws.append(columns + ["Galaxy Image"])

for col_num, name in enumerate(
    columns + ["Galaxy Image"], start=1
):
    cell = ws.cell(row=1, column=col_num)
    cell.font = Font(
        bold=True,
        color="FFFFFF"
    )
    cell.fill = PatternFill(
        "solid",
        fgColor="17365D"
    )
    cell.alignment = Alignment(
        horizontal="center"
    )

image_column = len(columns) + 1

for row_number, (_, galaxy) in enumerate(
    df.iterrows(), start=2
):

    for col_number, column in enumerate(
        columns, start=1
    ):
        value = galaxy[column]

        if pd.isna(value):
            value = None

        if column == "objid" and value is not None:
            value = str(value)

        ws.cell(
            row=row_number,
            column=col_number,
            value=value
        )

    ws.row_dimensions[row_number].height = 125

    image_path = galaxy["image_file"]

    if image_path and Path(image_path).exists():

        excel_image = ExcelImage(image_path)

        excel_image.width = 160
        excel_image.height = 160

        cell_location = (
            f"{get_column_letter(image_column)}"
            f"{row_number}"
        )

        ws.add_image(
            excel_image,
            cell_location
        )

# ==========================================
# 6. FORMAT EXCEL
# ==========================================

for col_number, column in enumerate(
    columns, start=1
):

    letter = get_column_letter(col_number)

    if column == "objid":
        ws.column_dimensions[letter].width = 23

    elif column == "image_file":
        ws.column_dimensions[letter].width = 35

    else:
        ws.column_dimensions[letter].width = 23

ws.column_dimensions[
    get_column_letter(image_column)
].width = 25

ws.freeze_panes = "A2"
ws.auto_filter.ref = ws.dimensions

# ==========================================
# 7. SAVE EXCEL FILE
# ==========================================

excel_filename = "galaxies_with_images.xlsx"

wb.save(excel_filename)

print("\n================================")
print("PROCESS COMPLETED")
print("================================")
print(f"Galaxies processed: {len(df)}")
print(
    "Images downloaded:",
    sum(bool(path) for path in image_paths)
)
print("CSV: galaxies_with_images.csv")
print("Excel: galaxies_with_images.xlsx")
print("Images folder: galaxy_images")
print("================================")
