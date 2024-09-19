# Description: This script is used to swap flies in an Excel file based on a swap sheet and a threshold value
import logging
import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import Button, Entry, Label, StringVar, filedialog, Tk
from PIL import Image, ImageTk
import random


logging.basicConfig(level=logging.INFO)


def display_image(title: str, image_path: str) -> None:
    """
    Displays an image in a GUI window
    Args:
        image_path (str): The location of the image to display
    """
    window = tk.Toplevel()
    window.title(title)

    img = Image.open(image_path)
    img = img.resize((500, 500))
    img = ImageTk.PhotoImage(img)
    label = Label(window, image=img)
    label.image = img  # Keep a reference to the image to prevent garbage collection
    label.pack()
    logging.info("Scott is here to help you through this tough time")
    logging.info("Scott leaves you to your own devices")


def get_random_picture_file_path() -> str:
    """
    Gets a random picture file path from the docs directory
    Returns:
        str: The file path of the random picture
    """
    docs_dir = os.path.join(os.path.dirname(__file__), "docs/happy photos")
    files = os.listdir(docs_dir)
    files = [f for f in files if f.endswith((".jpg", ".png"))]
    file_path = os.path.join(docs_dir, random.choice(files))
    return file_path


def get_file_path_from_gui():
    """
    Opens a file explorer to select the flies Excel file
    Returns:
        str: The file path selected by the user
    """
    global file_name
    file_name = filedialog.askopenfilename(
        title="Select the flies Excel file",
        filetypes=[("Excel files", "*.xlsb *.xlsx *.xls")],
    )
    logging.info(f"Selected file: {file_name}")


def main_gui() -> None:
    """
    Main GUI for the Fly Swapper
    """
    global threshold_input, threshold
    root = Tk()
    root.title("Fly Swapper")
    text = Label(root, text="Fly Swapper")
    text.pack()

    # File Selection
    text = Label(root, text="Select the flies Excel file")
    button = Button(root, text="Select File", command=get_file_path_from_gui)
    text.pack()
    button.pack()

    # Threshold Entry
    text = Label(root, text="Enter the threshold value")
    text.pack()
    threshold_input = StringVar(value=threshold)
    entry = Entry(root, textvariable=threshold_input, justify="center")
    entry.pack()

    # Submit Button
    button = Button(root, text="Submit", command=process_sheet)
    button.pack()

    # Happy Button
    button = Button(
        root,
        text="Press For Happiness",
        command=lambda: display_image(
            "You should probably stop pressing this button....",
            get_random_picture_file_path(),
        ),
    )
    button.pack()
    root.mainloop()


def validate_path() -> None:
    """
    Check if the file path exists
    """
    global file_name
    if os.path.exists(file_name):
        logging.info(f"Good job, the path {file_name} exists!!")
    else:
        logging.error(
            "It looks like the path you gave me doesn't exist.... "
            "You sure you have supplied the right path?"
            "The path you gave me was: " + file_name
        )
        logging.error("Closing Script, try again")
        sys.exit(-1)


def load_sheet() -> tuple:
    """
    Load the core and swap sheets from the Excel file
    Returns:
        : The core and swap sheets
    """
    global file_name, core_sheet_name, swap_sheet_name
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
    return core_sheet, swap_sheet


def validate_sheet(core_sheet: pd.DataFrame, swap_sheet: pd.DataFrame) -> None:
    global file_name, core_sheet_name, swap_sheet_name
    # Check how many flys there are
    num_columns = core_sheet.columns.size
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
            f"Your {core_sheet_name} doesn't have the right amount of columns: {num_swap_columns}. It should have {CONST_NUM_SWAP_COLUMNS} columns"
        )
        logging.error("Closing Script, try again")
        sys.exit(-2)
    else:
        logging.info(
            f"It looks like the {swap_sheet_name} part of {file_name} is correctly formatted!"
        )


def find_invalid_flies(core_sheet: pd.DataFrame) -> list:
    """
    Goes through the core sheet and finds the flies that are invalid based on the threshold
    Args:
        core_sheet (pd.DataFrame): The core fly sheet
    Returns:
        list: The index number of the invalid flies
    """
    global threshold
    invalid_flies = []
    for column in core_sheet.columns:
        if column.startswith("y"):
            first_coord = core_sheet[column].iloc[0]
            logging.debug(f"First coordinate is {first_coord} of fly {column}")
            if first_coord > threshold:
                logging.debug(f"Fly {column} is invalid")
                invalid_flies.append(
                    column.strip("y")
                )  # Remove the y from the column name
    logging.info(
        f"Invalid flies are: {invalid_flies} since they are above the threshold of {threshold}"
    )
    return invalid_flies


def set_invalid_flies_to_nan(
    core_sheet: pd.DataFrame, invalid_files: list
) -> pd.DataFrame:
    """
    Sets the invalid flies to NaN in the core sheet
    Args:
        core_sheet (pd.DataFrame): The core sheet
        invalid_files (list): The list of invalid flies

    Returns:
        pd.DataFrame: The core sheet with the invalid flies set to NaN
    """
    for fly in invalid_files:
        core_sheet[f"x{fly}"] = core_sheet[f"x{fly}"].apply(lambda x: pd.NA)
        core_sheet[f"y{fly}"] = core_sheet[f"y{fly}"].apply(lambda x: pd.NA)
    return core_sheet


def process_sheet():
    """
    Process the flies based on the swap sheet provided and the threshold and saves the new file
    """
    global file_name, threshold, threshold_input
    # Check if the file exists in the system, crash if not
    validate_path()
    [core_sheet, swap_sheet] = load_sheet()
    validate_sheet(core_sheet, swap_sheet)
    logging.info(
        "Brace yourself—it's about to get bumpier than a night out with too much tequila"
    )
    logging.info("I'm going to start processing the flies now...")

    try:
        logging.info("Retrieving the threshold value")
        threshold = float(threshold_input.get())
    except ValueError:
        logging.error(f"Threshold value is not a number: {threshold}. Exiting.")
        sys.exit(-2)
    logging.info(f"Threshold has been set to {threshold}")
    # Find the invalid flies
    invalid_files = find_invalid_flies(core_sheet)
    core_sheet = set_invalid_flies_to_nan(core_sheet, invalid_files)

    logging.info(f"I Detected {len(swap_sheet)} swaps! Time to get swapping!!")

    # Make an list of tuples
    for _, row in swap_sheet.iterrows():
        fly_A = row["flyA"]
        fly_B = row["flyB"]
        frame_start = row["FrameStart"]
        frame_end = row["FrameEnd"]

        logging.info(
            f"Fly {fly_A} will be swapped with fly {fly_B} at frame {frame_start} to frame {frame_end}"
        )

        # Check if the flies are invalid
        if str(fly_A) in invalid_files or str(fly_B) in invalid_files:
            logging.error(
                "Either Fly A or Fly B is invalid, so should not be swapped. Exiting."
            )
            sys.exit(-2)

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

    core_sheet.to_excel(
        swapped_file_name, sheet_name="Scott Processed Flies", index=False
    )

    # 30% chance to display an image
    sadness_factor = 0.3  # Increase if more sad
    if random.random() < sadness_factor:
        display_image("Scott cares about you...", os.path.join(os.path.dirname(__file__),"docs\\Hang in there.jpg"))

    logging.info(f"Someone saved the file as {swapped_file_name}")
    logging.info("Scott has left the building.... why was he even here?")

    # Display Success Popup
    display_image("Much Wow, you've swapped some flies", os.path.join(os.path.dirname(__file__),"docs\\Thumbs Up.jpg"))


# Global Variables
threshold_input = None
threshold = 900.0
file_name = "Remember to select the file button"
swap_sheet_name = "swap"
core_sheet_name = "trajectories"

main_gui()
