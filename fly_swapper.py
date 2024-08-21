import logging
import os
import sys
import pandas as pd
logging.basicConfig(level=logging.INFO)

file_name = r".\example\trajectories (version 1).xlsb.xlsx"  # path to file + file name

if(os.path.exists(file_name)):
    logging.info(f"Good job, the path {file_name} exists!!")
else:
    logging.error("It looks like the path you gave me doesnt exist.... You sure you have supplied the right path?")
    logging.error("Closing Script, try again")
    sys.exit(-1)

## Sheet Names
core_sheet_name = "trajectories"
swap_sheet_name = "swap"

## Load the sheets in
core_sheet = pd.read_excel(file_name,sheet_name=core_sheet_name)
swap_sheet = pd.read_excel(file_name, sheet_name=swap_sheet_name)

# Check how many flys there are
num_columns = core_sheet.columns.size
num_rows = len(core_sheet)
num_flys = (num_columns-1)/2

logging.info(f"There are {num_columns} in this spreadsheet, this means there are {num_flys} flys in this sheet")

# Check the formatting of the swap sheet
CONST_NUM_SWAP_COLUMNS = 4
# Check how many flys there are
num_swap_columns = swap_sheet.columns.size
if num_swap_columns is not CONST_NUM_SWAP_COLUMNS:
    logging.error(f"Your {swap_sheet_name} sheet has too many columns ({num_swap_columns}) it should have {CONST_NUM_SWAP_COLUMNS} ")
    logging.error("Closing Script, try again")
    sys.exit(-2)
else:
    logging.info(f"It looks like the {swap_sheet_name} part of {file_name} is correctly formatted!")

logging.info("Brace yourself—it's about to get bumpier than a night out with too much tequila")
logging.info(f"I Detected {len(swap_sheet)} swaps! Time to get swapping!!")

# Make an list of tuples
for index, row in swap_sheet.iterrows():
    fly_A = row['flyA']
    fly_B = row['flyB']
    frame_start = row['FrameStart']
    frame_end = row['FrameEnd']

    logging.info(f"Fly {fly_A} will be swapped with fly {fly_B} at frame {frame_start} to frame {frame_end}")
    for frame in range(frame_start, frame_end+1):
        # Temp Store the fly A
        temp_fly_pos_x =core_sheet.at[frame,f"x{fly_A}"]
        temp_fly_pos_y =core_sheet.at[frame,f"y{fly_A}"]
        logging.debug(f"Temporarily Storing Fly {fly_A} which was at [{temp_fly_pos_x},{temp_fly_pos_y}] at Frame {frame}")
        # Swap fly A with B
        core_sheet.at[frame,f"x{fly_A}"]=core_sheet.at[frame,f"x{fly_B}"]
        core_sheet.at[frame,f"y{fly_A}"]=core_sheet.at[frame,f"y{fly_B}"]
        # Swap Fly B with the Temp
        core_sheet.at[frame,f"x{fly_B}"]=temp_fly_pos_x
        core_sheet.at[frame,f"y{fly_B}"]=temp_fly_pos_y
        logging.debug(f"Fly {fly_A} has now been swapped with fly {fly_B} at Frame {frame}")


logging.info("Right about now I would start thinking about thanking scott....")

directory, filename = os.path.split(file_name)
basename, ext = os.path.splitext(filename)
swapped_file_name = f"{basename}{"_scott_processed"}{ext}"
swapped_file_name= os.path.join(directory, swapped_file_name)

core_sheet.to_excel(swapped_file_name, sheet_name='Scott Proccesed Flies', index=False)
