# Meseleler
# Değişken adına yasak karakter basılmamalı +
# Reserved değişlenlerine numara eklenmeli +
# Max_row'u olduğundan büyük verince hata veriyor -> bunu düzeltmiyorum, manuel girilmeli
# Byte sayımında döküman hatalı ise fark etmiyor +
# Türkçe karakterleri ingilizceleştir +
# Program bir hatadan dolayı kullanıcıya prompt verip kesinitiye uğradığında 
# byte sayısı sıfırlanıyor. Bunun önüne geçmek lazım +
# Program tekrar çalıştırıldığında reserved sayısını da unutuyor +
# Byte uyumsuzluğunda hata ver +
# Açıklamaları comment olarak yaz + 
# Hatalı satırı bastır +
# Byte uyumsuzluğu için byte farkını alan bir flag yarat - gerek yok
# Byte içindeki bit order'ını değiştir 7-0 -> 0-7

import openpyxl
import sys

excel_path = "message.xlsx"

out_file_path = "output.txt"

# Returns True if type of cell is "Cell". Otherwise returns False. 
# Used to bypass empty rows
def check_type(cell):
    type = cell.__str__().split(" ", 1)[0].split("<",1)[1]
    
    if type == "Cell":
        return 1
    elif type == "MergedCell":
        return 2
    else:
        return 0

# Print an explanatory error message using given msg
def excel_error(row_num, msg):
    print("{} error at line {}".format(msg, row_num))

# Delete spaces and invalid characters for variables. Convert turkish chars 
# to english. The returned string will be in camel case
def camelify(var_name):
    # Turkish characters
    turkish = {"ç": "c", "ğ":"g" , "ı":"i","ö":"o", "ş":"s", "ü":"u"}

    # If the variable starts with a number prefix it with 'v'
    if var_name[0].isnumeric():
        var_name = "v" + var_name

    # print("var_name: ", var_name)

    result = ""
    prev_ch = None
    for ch in var_name:
        # Alphanumerical chars are acceptable
        if ch.isalnum():
            pass
        # Convert the turkish chars to english
        elif ch in turkish:
            ch = turkish[ch]
        # Skip spaces and non-alphanumerical characters
        else:
            prev_ch = ch 
            continue

        # If previous char is skipped, capitalize this char
        if prev_ch == " ":
            ch = ch.upper()

        # Add the char to the result and update the previous char
        result += ch
        prev_ch = ch

    return result

# Gives control to the user or automatically handles the problem when a divergence
# btw. document byte count and program byte count occurs
def byte_error(row_num, diff, row, last_row, last_row_num):
    # Print excel error
    excel_error(row_num, "Byte field")
    # Print last two rows
    print_last_rows(row, last_row, row_num, last_row_num)
    print("Document byte count and program byte count dows not match")
    print("If you want to manually handle the error, print 'y'. " + \
            "Print  'n'  if you want the program itself to write the bytes" + \
            " according its byte count")
    
    inp = input()

    if inp.lower() == 'y':
        return False
    elif inp.lower() == 'n':
        process_row.byte_diff = diff
        return True
    else: # Invalid option
        byte_error(row_num, diff, row, last_row, last_row_num)

# Write the notes field of this row to the C file as a comment
def process_comment(row, out_file, depth = 0, bit = False):
    # Get the notes column
    notes = row[5].value

    # Get the bit column
    bits = row[1].value

    # print(notes)

    if not bit:
        out_file.write("{}// {}\n".format("\t"*depth,notes))
    elif bit:
        out_file.write("{}// Bits {} // {}\n".format("\t"*depth,bits,notes))

def process_bit_rows(bit_rows, out_file):
    for row_ind in range(1,len(bit_rows)+1):

        (var_name, bit_size, row) = bit_rows[-row_ind]

        # Non-comment (Cell type) rows
        if var_name is not None:
            # Write the variable to out_file
            out_file.write("\t\tuint8_t {} : {}; ".format(var_name, bit_size))

            # Process the notes field as a comment
            process_comment(row, out_file, bit = True)

            # Process comments of this non-comment row
            row_ind-=1

            (var_name, bit_size, row) = bit_rows[-row_ind]

            while var_name is None:
                process_comment(row, out_file, depth= 2)
                row_ind -= 1
                (var_name, bit_size, row) = bit_rows[-row_ind]

    out_file.write("\t};\n")



# Given the row, output file and number of the current row, processes the row and prints
# error message if an error is detected in excel. Return a boolean value telling whether 
# the line is successfully processed
def process_row(row, out_file, row_num, last_row = None, bit_rows = None, last_row_num = None):
    # Depth of the row. False is equivalent to 0
    result = False

    # Whether the bit rows of a byte have all been appended to the  list
    # 0 -> currently not processing a multibit byte
    # 1 -> currently processing a multibit byte, but not yet finished
    # 2 -> just finished the multibit byte
    bit_rows_completed = 0

    # Cells of the row
    byte = row[0].value
    bit = row[1].value
    data = row[2].value
    length = row[3].value

    # Split the byte to integers
    bytes = str(byte).split("-")
    bytes = list(map(lambda x: int(x),list(map(lambda x: x.strip(),bytes))))
    print("Processing bytes {}".format(bytes))

    # Byte size
    size = None

    # Whether the currently processed variable is single or multi-byte
    multi_byte = None

    # Check whether the written byte and byte count matches
    if bytes[0] != process_row.byte_count - process_row.byte_diff:
            if not byte_error(row_num, process_row.byte_count - bytes[0], row, last_row, last_row_num):
                result = False
                return (result, bit_rows_completed)
            
            

    # Name of the variable in the document in camelified form
    var_name = camelify(data)

    # Check if the variable is reserved
    is_reserved = var_name == "Reserved"
    if is_reserved:
        process_row.reserved_count+=1
        var_name = "{}{}".format(var_name, process_row.reserved_count)

    # Depending on the byte size, print the int type to out_file
    if len(bytes) == 1: # Single byte, 8-bit
        multi_byte = False

        # Byte size
        size = 1

        # Size of the variable in bits
        bit_size = 0    

        # The byte is divided into multiple variables
        if bit != "7-0":

            # Currently processing a multibit byte
            bit_rows_completed = 1

            # Split the bit to integers
            bits = str(bit).split("-")
            bits = list(map(lambda x: int(x),list(map(lambda x: x.strip(), bits))))

            # If there is a discontinuation btw. bits, give excel error
            # print("prev bit: {} and current bit: {}".format(process_row.prev_bit, int(bits[0])))
            # print("prev byte: {} and current byte: {}".format(process_row.prev_byte, bytes[0]))
            # print("byte count pre addition: {}".format(process_row.byte_count))


            # print("prev byte {}".format(process_row.prev_byte))
            # print("byte[0] {}".format(bytes[0]))
            # print("prev bit {}".format(process_row.prev_bit))
            # print("bit [0] {}".format(bits[0]))
            
            # If there is a previous bit
            # Either the current byte and the byte of previous variable are same and their bits are consecutive
            # Or their bytes are consecutive and their bits and previous ends with 0 and next starts with 7
            if not(process_row.prev_bit == -1 or \
                (process_row.prev_byte == bytes[0] and process_row.prev_bit == bits[0]+1) \
                or (process_row.prev_byte == bytes[0]-1 and process_row.prev_bit==0 and bits[0]==7)):
                
                # Print excel error
                excel_error(row_num, "Bit field")
                # Print last two rows
                print_last_rows(row, last_row, row_num, last_row_num)
                result = False
                return (result, bit_rows_completed)
            
            # Start of a new struct
            if bits[0] == 7:
                out_file.write("\tstruct // Byte {}\n\t{{\n".format(process_row.byte_count))

            # Whether the variable is multibit
            multi_bit = None

            if len(bits) == 1:
                multi_bit = False
                # print("Sinle bit Last bit: {}".format(bits))
            elif len(bits) == 2:
                multi_bit = True
                # print("Multi bit Last bit: {}".format(bits))
            else:
                # Print excel error
                excel_error(row_num, "Bit field")
                # Print last two rows
                print_last_rows(row, last_row, row_num, last_row_num)
                result = False
                return (result, bit_rows_completed)
            
            print("Processing bits {}".format(bits))
            
            # Calculate the size of the variable in bits
            if multi_bit:
                bit_size = bits[0] - bits[1] + 1 
            else:
                bit_size = 1

            # Append the row to the bit_rows
            bit_rows.append((var_name, bit_size, row))

            # Write the variable to out_file
            # out_file.write("\t\tuint8_t {} : {}; ".format(var_name, bit_size))

            # Process the notes field as a comment
            # process_comment(row, out_file)

            # If the last bit is 0, close the struct and increase the byte count
            if (multi_bit and bits[1]== 0) or (not multi_bit and bits[0]==0) :
                # out_file.write("\n\t};\n")

                # All the bits of the byte has been counted
                bit_rows_completed = 2

                # Increase the byte count
                process_row.byte_count += size

            # Hold the previous bit
            process_row.prev_bit = bits[1] if multi_bit else bits[0] 

            result = 2
          
        # Single variable for the byte
        else:
            bit_size = 8
            # Write the variable to out_file
            out_file.write("\tuint8_t {}; // Byte {} ".format(var_name, process_row.byte_count))

            # Process the notes field as a comment
            process_comment(row, out_file)

            # Increase the byte count
            process_row.byte_count += size

            result = 1

        
        # Check that the length column in the document is True
        if not is_reserved and int(length) != bit_size:
            # Print excel error
            excel_error(row_num, "Length field")
            # Print last two rows
            print_last_rows(row, last_row, row_num, last_row_num)
            result = False
            return (result, bit_rows_completed)

    else: # Multi byte
        multi_byte = True

        size = bytes[1] - bytes[0] + 1

        # Check that the length column in the document is True
        if not is_reserved and int(length) != 8*size:
            # Print excel error
            excel_error(row_num, "Length field")
            # Print last two rows
            print_last_rows(row, last_row, row_num, last_row_num)
            return (result, bit_rows_completed)

        # Print out the previous and current bytes
        # print("prev byte: {} and current byte: {}".format(process_row.prev_byte, bytes[1]))
        # print("byte count pre addition: {}".format(process_row.byte_count))

        if size == 2: # 16-bit
            # check if the bit field is correct
            if bit != "15-0":
                # Print excel error
                excel_error(row_num, "Bit column")
                # Print last two rows
                print_last_rows(row, last_row, row_num, last_row_num)
                return (result, bit_rows_completed)
            
            out_file.write("\tuint16_t {}; // Byte {}-{} ".format(var_name, process_row.byte_count, process_row.byte_count+size-1)) # Write the type to output file

        elif size == 4: # 32-bit
            # check if the bit field is correct
            if bit != "31-0":
                # Print excel error
                excel_error(row_num, "Bit column")
                # Print last two rows
                print_last_rows(row, last_row, row_num, last_row_num)
                return (result, bit_rows_completed)
            out_file.write("\tuint32_t {}; // Byte {}-{} ".format(var_name, process_row.byte_count, process_row.byte_count+size-1)) # Write the type to output file

        elif size == 8: # 64-bit
            # check if the bit field is correct
            if bit != "63-0":
                # Print excell error
                excel_error(row_num, "Bit column")
                # Print last two rows
                print_last_rows(row, last_row, row_num, last_row_num)
                return (result, bit_rows_completed)
            out_file.write("\tuint64_t {}; // Byte {}-{} ".format(var_name, process_row.byte_count, process_row.byte_count+size-1)) # Write the type to output file

        else: # Unacceptable bit size: Error
            # Print excel error
            excel_error( row_num, msg = "Byte column")
            # Print last two rows
            print_last_rows(row, last_row, row_num, last_row_num)
            return (result, bit_rows_completed)
        
        # Process the notes field as a comment
        process_comment(row, out_file)

        # Increase the byte count
        process_row.byte_count += size

        result = 1
    

    # A static variable to hold the previous byte
    process_row.prev_byte = bytes[1] if multi_byte else bytes[0]

    return (result, bit_rows_completed)

# Print the content of the row
def print_row(row):
    for cell in row:
        print(cell.value, end= "   ")
    print()

def print_last_rows(row, last_row, row_num, last_row_num):
    print("Last two rows:")
    if last_row is not None:
        print("line {}:".format( last_row_num), end = " ")
        print_row(last_row)
    
    if row is not None:
        print("line {}:".format(row_num), end = " ")
        print_row(row)

# Usage:
# python3 first.py <option> -i <input file> -o <output file> -l <line start> <end>
def driver():
    # Correct usage
    usage = "Correct usage is:\n" + "python3 dedocumenter.py (-d | -c) -i <input file>" + \
        " [-o <output file>] -l <line start> <end> [-r <reserved variable no>]" + \
        " [-b <byte count>] [-p]"

    # Length of the argument list
    n = len(sys.argv)

    # At least option and input file must be specified
    if n < 5:
        print("Not enough arguments")
        print(usage)
        return False
    
    option = sys.argv[1]

    overwrite = None

    # Delete/overwrite option
    if option == "-d":
        overwrite = True
    # Continue
    elif option == "-c":
        overwrite = False
    # Invalid option
    else:
        print("invalid option {}. Option should be either -d or -c".format(option))
        print(usage)
        return False

    

    excel_path = None
    out_file_path = None
    start = None
    end = None
    res = None
    byte_start = None
    total_packet = False # Whether the given line interval includes the total 
                        # message packet or not

    # Iterate over arguments
    a = 2
    while a < n:
        # Input flag
        if sys.argv[a] == "-i":
            if a+1 >= n: # Not enough arguments
                print("Expected an input file address after -i flag")
                return False
            excel_path = sys.argv[a+1]
            a+=2
            continue
        elif sys.argv[a] == "-o":
            if a+1 >= n: # Not enough arguments
                print("Expected an output file address after -o flag")
                return False
            out_file_path = sys.argv[a+1]
            a+=2
            continue
        elif sys.argv[a] == "-l":
            if a+2 >= n: # Not enough arguments
                print("Expected the line start and end values after -l flag")
                return False
            start = int(sys.argv[a+1])
            end = int(sys.argv[a+2])
            a+=3
            continue
        elif sys.argv[a] == "-r":
            if a+1 >= n: # Not enough arguments
                print("Expected the reserved variable no after -r flag")
                return False
            res = int(sys.argv[a+1])
            a+=2
            continue
        elif sys.argv[a] == "-b":
            if a+1 >= n: # Not enough arguments
                print("Expected the byte count after -b flag")
                return False
            byte_start = int(sys.argv[a+1])
            a+=2
            continue
        elif sys.argv[a] == "-p":
            total_packet = True
            a+=1
            continue
        else:
            print("Invalid command line argument")
            print(usage)
            return False

    # Check if there is an input file
    if excel_path == None:
        print("No input file specified")
        print(usage)
        return False

    # Open workbook
    wb = openpyxl.load_workbook(excel_path)

    # Open sheet
    sheet = wb.active

    # Input file name w/o extension
    file = excel_path.split(".",1)[0]

    if out_file_path == None:
        out_file_path = file + ".h"
        print("No output file specified. Defaulting to {}".format(out_file_path))

    if start == None:
        print("No line start and end is given")
        print(usage)
        return False
    
    if res == None:
        if overwrite:
            print("No reserved variable no is specified. Defaulting to 0")
            res = 0
        else:
            print("In continuation mode last reserved variable no must be given")
            print(usage)
            return False
        
    if byte_start == None:
        if overwrite:
            print("No byte count is specified. Defaulting to 0")
            byte_start = 0
        else:
            print("In continuation mode last starting byte count must be given")
            print(usage)
            return False
        
    
    # If overwrite, delete the content of the file first
    if overwrite:
        open(out_file_path, 'w').close()

    # Open the output file
    out_file = open(out_file_path, 'a')

    # Number of the current row
    row_num = start

    # Set previous byte and previous bit fields and byte_count of process row
    process_row.prev_byte = -1 # Number of the byte that is last processed
    process_row.prev_bit = -1 # Number of the bit that is last processed
    process_row.byte_count = byte_start # Total number of the bytes covered so far
    process_row.reserved_count = res # Total number of reserved variables so far
    process_row.byte_diff = 0 # Difference between byte count in the input document
                            # and program's own byte count

    # Only start the massage structure in overwrite form
    if overwrite:
        # Print the beginning of structure
        out_file.write("typedef struct\n{\n");

    # Variable to hold whetehr the current row of excel is processed successfully
    row_depth = True

    # Store the last row for debugging
    last_row = None
    last_row_num = None

    # A list to hold multi-bit bytes
    bit_rows = []

    # Whether the multibit byte is completely processed
    bit_rows_completed = 0

    # Iterate the rows of the excel file
    for row in sheet.iter_rows(min_row = start, max_row = end, max_col = 6):
        
        # Get the cell type
        type_checked = check_type(row[0])

        # Process the Cell type cells
        if type_checked == 1:
            # Process rows of multibit bytes including comment rows
            if bit_rows_completed == 2:
                process_bit_rows(bit_rows, out_file)
                bit_rows = []
                bit_rows_completed = 0

            # Process the row
            (row_depth, bit_rows_completed) = process_row(row, out_file, row_num, last_row, bit_rows, last_row_num)
            if not row_depth:
                # If the row is not processed successfully, print the last two rows
                # and terminate the program
                break

            # Store the last Cell type row
            last_row = row
            last_row_num = row_num

        # Process MergedCell type cells as comment 
        elif type_checked == 2:

            # Comments that are not part of a multibit byte
            if bit_rows_completed == 0:
                process_comment(row, out_file, row_depth)

            # Comments that are part of a multibit byte
            elif bit_rows_completed == 1:
                bit_rows.append((None, None, row))
        
        row_num+=1

    # Do not forget to process rows of multibit bytes at the last row of excel
    if bit_rows_completed == 2:
        process_bit_rows(bit_rows, out_file)
        bit_rows = []
        bit_rows_completed = 0

    # If the total packet is processed to its last line
    if total_packet and row_num > end:
        # Close the struct
        out_file.write("}} {};\n".format(file))
        print("All of the message packet is processed")

    # Close the output file
    out_file.close()

if __name__=="__main__":
    
    driver()
    

    
