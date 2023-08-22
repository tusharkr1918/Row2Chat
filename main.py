from sys import exit
from pandas import read_excel
from subprocess import Popen, PIPE
from time import sleep
from cv2 import imread, matchTemplate, minMaxLoc, rectangle, imwrite, TM_CCOEFF_NORMED
from re import findall
from argparse import ArgumentParser
from os import path
from PIL import Image
import shutil

def error_counter():
    num = 1
    while True:
        yield num
        num += 1
error_gen = error_counter()

def find_template_and_get_coordinate(image_path, template_path, method=TM_CCOEFF_NORMED, threshold=0.8):
    # Load the original image
    original_image = imread(image_path)
    template = imread(template_path)

    # Define the coordinates of the ROI
    rheight = 260
    image_height, image_width = original_image.shape[:2]
    roi_x = 0
    roi_y = image_height - rheight  # Starting from 160 pixels up from the bottom
    roi_width = image_width
    roi_height = rheight  # A height of 160 pixels

    # Extract the ROI from the original image
    roi = original_image[roi_y:roi_y + roi_height, roi_x:roi_x + roi_width]

    ## FOR DEBUGGING ##
    # rectangle(roi, (0, 0), (roi_width, roi_height), (0, 255, 0), 2)
    # imwrite("Images/roired.png", original_image)

    # Process the ROI as needed
    result = matchTemplate(roi, template, method)
    _, max_val, _, max_loc = minMaxLoc(result)

    if max_val >= threshold:
        # Get the top-left corner of the detected area in the ROI
        top_left_roi = max_loc

        # Calculate the corresponding top-left corner in the original image's global coordinates
        top_left_original = (top_left_roi[0] + roi_x, top_left_roi[1] + roi_y)

        # Calculate the bottom-right corner of the detected area in the original image's global coordinates
        bottom_right_original = (top_left_original[0] + template.shape[1], top_left_original[1] + template.shape[0])

        # Calculate the center pixel coordinates
        center_x = (top_left_original[0] + bottom_right_original[0]) // 2
        center_y = (top_left_original[1] + bottom_right_original[1]) // 2

        # FOR DEBUGGING ##
        # Save the modified image with the bounding box
        # rectangle(original_image, top_left_original, bottom_right_original, (0, 255, 0), 2)
        # imwrite("Images/red.png", original_image)
        return (center_x, center_y)
    return False

def adb(command):
    # Use the Popen class to run the specified ADB command
    proc = Popen(command.split(' '), stdout=PIPE, shell=True)
    
    # Communicate with the process and capture its standard output and error streams
    (out, _) = proc.communicate()
    
    # Return the captured standard output
    return out

def unlock_phone(pin):
    # Press the power button to wake up the device
    adb('adb shell input keyevent 26')
    
    # Simulate a swipe gesture to unlock the device (from coordinates (100, 100) to (500, 500))
    adb('adb shell input swipe 100 100 500 500')
    
    # Enter the provided PIN using the input text command
    adb(f'adb shell input text {pin}')
    
    # Press the enter key to confirm the PIN
    adb('adb shell input keyevent 66')


def send(tap_x, tap_y):
    # Construct the ADB shell command to simulate a touch event at the specified coordinates
    adb(f"adb shell input tap {tap_x} {tap_y}")


# Modify as you want..
def log_status(mobile, message, file_base_name, filterby, mark):

    sub_name = filterby[0]+'_'+ '_'.join(filterby[1]) if filterby != None else '_'

    # Open the CSV file named 'dispatch.csv' in append mode for writing
    with open(rf'logs/{file_base_name}_{sub_name}_info.csv', mode='a', encoding='utf-8') as log:
    #     # Extract policy number from the message using regular expression
    #     policy_num = findall(r'Policy\s*No\s*([0-9]*)', message)[0]
        
    #     # Extract agent code from the message using regular expression
    #     agent_code = findall(r'संख्या\s*([0-9A-Z]*)', message)[0]
        
    #     # Extract agent name from the message using regular expression
    #     agent_name = findall(r'प्रिय\s(.+?),', message)[0]
        
    #     # Write the extracted information along with status (mark) to the CSV file
        log.write(f"{mobile}, {message}, {mark}\n")

        if mark == 'Failed' and True:
            source_file = r'Images\sendScreenCompressed.jpeg'
            source_destination_folder = rf'error_snap\{mobile}_{next(error_gen)}.jpeg'
            shutil.move(source_file, source_destination_folder)

    pass

def compress_png(input_path, output_path, factor=80):
    try:
        image = Image.open(input_path)
        image = image.convert("RGB")
        image.save(output_path, format="JPEG", quality=factor)
    except Exception as e:
        print("An error occurred:", e)

def sendMessages(agt_mobile, message, file_base_name, filterby, release):
    # Launch WhatsApp with the specified phone number and message using ADB command
    adb(f'adb shell am start -a android.intent.action.VIEW -d "https://api.whatsapp.com/send?phone={agt_mobile}&text={message}"')

    # Pause for a few seconds
    adb('ping 127.0.0.1 -n 4 > nul')

    sleep(0.3)
    # Capture a screenshot of the screen
    adb("adb exec-out screencap -p > Images/sendScreen.png")
    sleep(0.8)

    # Paths for the main and template images for image recognition
    main_image_path = 'Images/sendScreen.png'
    compressed_main = 'Images/sendScreenCompressed.jpeg'
    template_image_path = 'Images/sendLogo.png'

    compress_png(main_image_path, compressed_main, factor=10)

    # Find the coordinates of a template image within the main image
    status = find_template_and_get_coordinate(compressed_main, template_image_path)
    
    if status:
        x_axis, y_axis = status

        # Perform an action (e.g., clicking) at the identified coordinates
        send(x_axis, y_axis) if release != False else None
        # Pause for a second
        sleep(0.2)
        adb(f'adb shell am force-stop com.whatsapp')

        # Log the status as Success
        log_status(agt_mobile, message, file_base_name, filterby, 'Success')
    else:
        # Log the status as Failed
        log_status(agt_mobile, message, file_base_name, filterby, 'Failed')
        adb(f'adb shell am force-stop com.whatsapp')

    # Force-stop the WhatsApp application

    # Return True if the status is not None, otherwise return False
    return True if status else False


def main(filename, filterby=None, dropna_number=True, sortby=None, skip_from=None, password=None, release=False):
    try:
        # Attempt to read the Excel file
        df = read_excel(filename)
    except FileNotFoundError as e:
        # If the file is not found, print an error message and exit
        exit(f"\n{e.__str__()[10:]} or filename is incorrect.\nPlease consider rechecking it.")

    # Extracting the file base name for logging
    filename_with_extension = path.basename(filename)
    file_base_name = path.splitext(filename_with_extension)[0]

    start = 1
    _end = df.shape[0]
    apply_skip = True if skip_from != None else False
    if apply_skip:
        start, end = skip_from
        df = df[start-1:end]
        df.reset_index(drop=True, inplace=True)

    try:
        # Check if a filter is specified
        apply_filter = True if filterby != None else False
        if apply_filter:
            # Unpack the filter criteria
            column, filter_value = filterby
            
            # Apply the filter to the DataFrame
            df = df[df[column].isin(filter_value)]

        # Check if a sorting criterion is specified
        apply_sort = None if sortby == None else True
        if apply_sort:
            # Get the column to sort by
            sortby_column = sortby[0]
            # Sort the DataFrame by the specified column in ascending order

            df.sort_values(by=sortby_column, ascending=sortby[1], inplace=True)
        
        if dropna_number == True:
            columns_to_check = ['Mobile']
            df.dropna(subset=columns_to_check, inplace=True)

    except KeyError as e:
        # If a specified column does not exist, print an error message and exit
        exit(f"\nThe {e} Column doesn't exist in the sheet.\nPlease consider rechecking it.")

    # If a password is provided, unlock the phone using the password
    if password is not None:
        unlock_phone(pin=password)

    # Close the WhatsApp if it is open before we start sending message
    adb(f'adb shell am force-stop com.whatsapp')

    # Initialize counters for successful and failed message sends
    success, failed = 0, 0
    # Iterate through the rows of the DataFrame
    for index, (_, row) in enumerate(df.iterrows(), 1 if start == 1 else start):
        # Extract agent's mobile number and message from the row
        mobile = row['Mobile']
        message = ' '.join(row['Message'].split())
        # Display progress and agent's phone number
        print(f'\u001b[0m{index}/{_end} - Mobile: {mobile}', end='')

        # Call sendMessages function to send WhatsApp message and get status
        status = sendMessages(mobile, message, file_base_name, filterby, release)
        if status:
            # If message was sent successfully, update the success counter
            success += 1
        else:
            # If message sending failed, update the failed counter
            failed += 1
        # Color-coded status message for success or failure
        status_color = '\u001b[1;32m Success \u001b[0m' if status else '\u001b[1;31m Failed  \u001b[0m'
        # Print status and cumulative counts
        print(f"{status_color} | Success: {success}, Failed: {failed}")

    # Simulate pressing the power button to lock the phone using ADB
    adb('adb shell input keyevent 26')

if __name__ == "__main__":
    # main(
    #     # File name: Specify the path to the Excel file.
    #     filename = r'Data/BGL-MTY-AGT-1208-FIN.xlsx',

    #     # Filtering by 'Branch' column. Provide a list of values to filter. (Optional)
    #     filterby = ['Branch', ['Katihar', 'Forbesganj']],

    #     # Sorting by column name. Use True for ascending order, and False for descending order. (Optional)
    #     sortby = ('S R', False),

    #     # Set to True by default. Set to False to escalate the message.
    #     release = False

    #     # Phone password: Provide the device's PIN|PASSWORD for unlocking. (Optional)
    #     # password = '1026'
    # )



    def parse_skipfrom(value):
        values = value.split(',')
        if len(values) == 1:
            return int(values[0]), None
        elif len(values) == 2:
            return tuple(map(int, values))
        else:
            raise ValueError("Invalid format for --skipfrom argument")
    
    parser = ArgumentParser(description="Automated WhatsApp Messaging from Excel Rows.")
    
    parser.add_argument("filename", help="path to the Excel file")
    parser.add_argument("--filterby", nargs="+", metavar=("column_name", "values"), help="filter data by column values")
    parser.add_argument("--sortby", nargs="+", metavar=("column_name", "ascending"), help="sort data by column in ascending order")
    parser.add_argument("--release", action="store_true", help="use the release flag to escalate the message.")
    parser.add_argument("--password", help="device's PIN|PASSWORD for unlocking")
    parser.add_argument("--skipfrom", type=parse_skipfrom, help="skip data starting from row and column")
    parser.epilog = (
        "Usage:\n"
        "  python main.py \"Data/BGL-MTY-AGT-1208-FIN.xlsx\" --filterby Branch \"Katihar,Forbesganj\" --sortby \"Agt Mobile\" True --password \"1026\" --skipfrom 1001,0"
    )
    args = parser.parse_args()
    
    filterby = None
    if args.filterby:
        filterby = [args.filterby[0], args.filterby[1].split(',')]
    
    sortby = None
    if args.sortby:
        sortby_column = args.sortby[0]
        sortby_ascending = args.sortby[1] == "True" if len(args.sortby) == 2 else True
        sortby = (sortby_column, sortby_ascending)
    
    print(args.filename, filterby, sortby, args.release, args.password, args.skipfrom)
    main(
        filename=args.filename,
        filterby=filterby,
        sortby=sortby,
        skip_from=args.skipfrom,
        release=args.release,
        password=args.password
    )