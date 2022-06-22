# Import required libraries
from util import excel, image
import sys
import os
import time
import threading

if __name__ != "__main__":
    exit()

# Fetch the arguments
invalid_args = lambda start : print(f"""{start}

Correct usage:
{os.path.split(__file__)[-1]} <input file> <output file> [image dimension division]""")
sys.argv.pop(0) # Remove file arg
if len(sys.argv) < 2 or len(sys.argv) > 3:
    invalid_args("Too little / much arguments passed!")
    exit()
elif not os.path.isfile(sys.argv[0]):
    if os.path.isdir(sys.argv[0]):
        invalid_args("The input file must be a file, not directory!")
        exit()
    else:
        invalid_args("The input file provided does not exist!")
        exit()
elif not os.path.isdir("\\".join(os.path.split(sys.argv[1])[:-1])):
    invalid_args("The directory of the output file does not exist!")
    exit()
elif os.path.isfile(sys.argv[1]):
    inp = input("The output file provided already exists and will be overwritten, continue? (y/N): ")
    if inp.lower() == "y":
        pass
    else:
        print("Okay, cancelling...")
        time.sleep(1)
        exit()
elif len(sys.argv) == 3:
    try:
        div = float(sys.argv[2])
    except:
        invalid_args("The image dimension division amount must be an integer or decimal!")
        exit()
input_file, output_file = sys.argv[0], sys.argv[1]
if len(sys.argv) == 3: image_size_division_amount = float(sys.argv[2])
else: image_size_division_amount = 3
cell_size_division_amount = 3 # Will change at some point

# Loading displays
def load(message: str, func, *args):
    cont = True
    def loop():
        n = 1
        while cont:
            sys.stdout.write("\r"+f"{message}{'.'*n}")
            time.sleep(0.75)
            sys.stdout.flush()
            n += 1
        sys.stdout.flush()
    threading.Thread(target=loop).start()
    try:
        func(*args)
    except Exception as e:
        sys.stdout.flush()
        raise e
    finally:
        cont = False

# Create the Excel document and load the image
try:
    exdat = excel.ExcelDoc()
    imdat = image.ImageData(input_file, image_size_division_amount)
    print("+ - - - - - - - - - - - - - - - - - +")
    load("Drawing image to Excel workbook", exdat.draw, imdat.rgba_map(), cell_size_division_amount) # Will add changable cell div amounts at some point
    print("\n")
    load("Writing to file", exdat.write_to, output_file)
    print(f"\n\nFinished, output: {output_file}")
    print("+ - - - - - - - - - - - - - - - - - +")
except KeyError or image.ImageTooLargeError:
    print("Image file too large for Excel cell limit, shutting down!")
    quit()
except PermissionError:
    print("\nPermission to the file denied! Make sure that Excel is closed before running the program.")
except Exception as e:
    print(e.with_traceback())
    print("!^^^ PLEASE REPORT AT https://github.com/jibbystar/image-to-excel/issues ^^^!")

