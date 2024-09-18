import logging
import os
import sys
import pandas as pd
import argparse
from tkinter import Label, Tk, filedialog
from PIL import Image, ImageTk
import random


logging.basicConfig(level=logging.INFO)


def get_file_path_from_gui() -> str:
    """
    Opens a file explorer to select the flies Excel file
    Returns:
        str: The file path selected by the user
    """
    root = Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(
        title="Select the flies Excel file",
        filetypes=[("Excel files", "*.xlsb *.xlsx *.xls")]
    )
    return file_path


def display_image(image_path: str) -> None:
    """
    Displays an image in a GUI window
    Args:
        image_path (str): The location of the image to display
    """
    root = Tk()
    root.title("Scott cares about you...")

    img = Image.open(image_path)
    img = ImageTk.PhotoImage(img)

    label = Label(root, image=img)
    label.pack()
    logging.info("Scott is here to help you through this tough time")
    root.mainloop()
    logging.info("Scott leaves you to your own devices")


# Set up argument parsing
parser = argparse.ArgumentParser(description="Fly Swapper Script")
parser.add_argument('--gui', action='store_true', help="Open file explorer to select the file")
args = parser.parse_args()

# Get the file path
if args.gui:
    file_name = get_file_path_from_gui()
    if not file_name:
        logging.error("No file selected. Exiting.")
        sys.exit(-1)
else:
    file_name = r".\example\trajectories (version 1).xlsb.xlsx"  # default path

# Check if the file exists in the system, crash if not
if os.path.exists(file_name):
    logging.info(f"Good job, the path {file_name} exists!!")
else:
    logging.error(
        "It looks like the path you gave me doesn't exist.... "
        "You sure you have supplied the right path?"
    )
    logging.error("Closing Script, try again")
    sys.exit(-1)

# Sheet Names
core_sheet_name = "trajectories"
swap_sheet_name = "swap"

# Load the sheets in
core_sheet = pd.read_excel(file_name, sheet_name=core_sheet_name)
try:
    swap_sheet = pd.read_excel(file_name, sheet_name=swap_sheet_name)
except ValueError:
    logging.error(
        f"Could not find the {swap_sheet_name} sheet in the file. Did someone not follow my instructions?? hmmmm."
    )
    logging.error("Closing Script, try again")
    logging.error("Or am i just a bad programmer.....  possibly")
    sys.exit(-3)

# Check how many flys there are
num_columns = core_sheet.columns.size
num_rows = len(core_sheet)
num_flys = (num_columns - 1) / 2

logging.info(
    f"There are {num_columns} in this spreadsheet, this means there are {num_flys} flys in this sheet"
)

# Check the formatting of the swap sheet
CONST_NUM_SWAP_COLUMNS = 4
# Check how many flys there are
num_swap_columns = swap_sheet.columns.size
if num_swap_columns is not CONST_NUM_SWAP_COLUMNS:
    logging.error(
        f"Your {swap_sheet_name} doesn't have the right amount of columns: {num_swap_columns}. It should have {CONST_NUM_SWAP_COLUMNS} columns"
    )
    logging.error("Closing Script, try again")
    sys.exit(-2)
else:
    logging.info(
        f"It looks like the {swap_sheet_name} part of {file_name} is correctly formatted!"
    )
# TODO: Check the format of the cells (do they have the right headers)

logging.info(
    "Brace yourself—it's about to get bumpier than a night out with too much tequila"
)
logging.info(f"I Detected {len(swap_sheet)} swaps! Time to get swapping!!")

# Make an list of tuples
for index, row in swap_sheet.iterrows():
    fly_A = row["flyA"]
    fly_B = row["flyB"]
    frame_start = row["FrameStart"]
    frame_end = row["FrameEnd"]

    logging.info(
        f"Fly {fly_A} will be swapped with fly {fly_B} at frame {frame_start} to frame {frame_end}"
    )
    for frame in range(frame_start, frame_end + 1):
        # Temp Store the fly A
        temp_fly_pos_x = core_sheet.at[frame, f"x{fly_A}"]
        temp_fly_pos_y = core_sheet.at[frame, f"y{fly_A}"]
        logging.debug(
            f"Temporarily Storing Fly {fly_A} which was at [{temp_fly_pos_x},{temp_fly_pos_y}] at Frame {frame}"
        )
        # Swap fly A with B
        core_sheet.at[frame, f"x{fly_A}"] = core_sheet.at[frame, f"x{fly_B}"]
        core_sheet.at[frame, f"y{fly_A}"] = core_sheet.at[frame, f"y{fly_B}"]
        # Swap Fly B with the Temp
        core_sheet.at[frame, f"x{fly_B}"] = temp_fly_pos_x
        core_sheet.at[frame, f"y{fly_B}"] = temp_fly_pos_y
        logging.debug(
            f"Fly {fly_A} has now been swapped with fly {fly_B} at Frame {frame}"
        )


logging.info("Right about now I would start thinking about thanking scott....")

directory, filename = os.path.split(file_name)
basename, ext = os.path.splitext(filename)
swapped_file_name = f"{basename}_scott_processed{ext}"
swapped_file_name = os.path.join(directory, swapped_file_name)

core_sheet.to_excel(swapped_file_name,
                    sheet_name="Scott Processed Flies",
                    index=False)

# 10% chance to display an image
sadness_factor = 0.1  # Increase if more sad
if random.random() < sadness_factor:
    display_image(".\\docs\\Hang in there.jpg")
