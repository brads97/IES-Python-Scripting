"""
Python script utilising the external libraries imported below to autmatically produce 
Excel spreadsheet of heating loads after an HTG file has been created from CIBSE Loads.

"""
import iesve
import xlsxwriter
import os
import numpy as np
import pandas as pd
from ies_file_picker import IesFilePicker
from operator import truediv

"""
Extracting information from ModelIt

"""

# Selecting the current IES project and initialising.
project = iesve.VEProject.get_current_project()

# Only rooms will be selected, any shading or roofs will be removed.
models = project.models
realmodel = models[0]

# Create list of VEBody objects from VEModel object.
bodies = realmodel.get_bodies(False)

# Creating variable lists to later be appended into excel.
# Clear variables by setting to empty list.
name_list = []
area_list = []

for body in bodies:
     # Create VERoomData object from VEBody object.
     # Skip any bodies that aren't thermal rooms
     if body.type != iesve.VEBody_type.room:
        continue
     room_data = body.get_room_data(type=0)
     general_room_data = room_data.get_general()
     
     # Create VERoomData object from VEBody object.
     room_name = general_room_data['name']
     floor_area = round(general_room_data['floor_area'], 1)
     name_list.append(room_name)
     area_list.append(floor_area)
     
print("Room Name -", name_list)
print("Room Area -", area_list)

"""
Extracting information from vista pro HTG file

"""

# Selecting the model results in VistaPro.
file_name = IesFilePicker.pick_vista_file([("HTG File","*.HTG")], "Navigate to and select a HTG File")
results = iesve.ResultsReader()
results.open_aps_data(file_name)

# Creating room_temp function which takes in the results and outputs the room temp

def room_temp(results):
    # Get the results data for room temp.
     room_temp = []
    # Initialise a list with one value 0 to catch a no rooms situation.
     for body in bodies:
        # Skip any bodies that aren't thermal rooms
        if body.type != iesve.VEBody_type.room:
            continue
        # Get the results data. z = room level data.
        set_point = results.get_room_results(body.id, 'Heating set point', 'Heating set point', 'z')
    
        # The temp results are a list.
        #print('Heating set point for', body.name, set_point)
       
        # Append max result value for each body to the list.
        room_temp.append(max(set_point))
        
    # Returns the value for each rooom.
     return (room_temp)

  
# Creating max_load_room function which takes in the results and outputs the maximum room load.
def max_load_room(results,bodies):
    max_load_room = []
    # Initialise a list with one value 0 to catch a no rooms situation.
    for body in bodies:
        # Skip any bodies that aren't thermal rooms
        if body.type != iesve.VEBody_type.room:
            continue
        # Get the results data. z = room level data.
        Heat_plant = results.get_room_results(body.id, 'Heating plant sensible load', 'Heating plant sensible load', 'z')
    
        # The temp results are a list.
        #print('Heating load for', body.name, Heat_plant)
       
        # Append max result value for each body to the list.
        max_load_room.append(round(max(Heat_plant),2))
        
    # Returns the maximum value for each rooom.
    return (max_load_room)


# Calling the functions above and printing outputs.
max_Heat_plant = max_load_room(results,bodies)
Set_Point = room_temp(results)

# Multiplying array by a design margin of 10%.
DM_Heat_plant = list(np.array(max_Heat_plant)*1.1)
print("Space Heating Load -", max_Heat_plant)
print("Design Load -", DM_Heat_plant)
print("Set Point -", Set_Point)

# Formatting outputs from results into a comblined list, to be input into a pandas DataFrame
elements = list(zip(name_list, area_list, Set_Point, max_Heat_plant, DM_Heat_plant))
df = pd.DataFrame(elements, columns = ['Name', 'Floor Area (m2)', 'Set Point (degC)', 'Space Heating Load (W)', '+10% Design Load (W)'])

# Dropping any rooms from the list with a Space heating load of <= 0. 
df_drop = df.drop(df[df['Space Heating Load (W)'] <= 0].index)

# Sort the name list in the dataframe alphabetically.
df_sort = df_drop.sort_values('Name')

# Adding Design load per m2 column by dividing the design load column by the floor area column.
df_sort['Design Load per m2 (W/m2)'] = df_sort['+10% Design Load (W)'] / df_sort['Floor Area (m2)']


"""
Now that we have the desired outputs to create our heating loads excel spreadsheet
We use the xlsx library imported earlier to format the DataFrame ito excel.

"""

# Writing into excel.
# Creating a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Heating Loads.xlsx', engine='xlsxwriter')
df_sort.to_excel(writer, sheet_name='Heating Loads',index=False, header=True)

# Convert the dataframe to an XlsxWriter Excel Object.
book = writer.book
sheet = writer.sheets['Heating Loads']

# Formatting rows and columns in excel.
# Setting column width and centre format
sheet_format = book.add_format()
sheet_format.set_align('Center')
sheet_format.set_num_format(2)
sheet.set_column('A:XFD', 25, sheet_format)

# Format for Set Point Column
setpoint_format = book.add_format()
setpoint_format.set_align('Center')
setpoint_format.set_num_format(0)

# Total string aligned to right and set to bold.
total_format = book.add_format()
total_format.set_align('right')
total_format.set_bold()

# Sum total set to centre, bold.
sum_format = book.add_format()
sum_format.set_bold()
sum_format.set_top(1)
sum_format.set_align('center')
sum_format.set_num_format(2)

# Column C  - Inserting temp set point again as no way to independently set rounding for this column, see setpoint_format.
sheet.write_column("C2", df_sort['Set Point (degC)'],setpoint_format)

# Column D  - Counting the number of rows within the df_sort function to determine the placement of 'Total'.
count_row = df_sort.shape[0]
numb_row = count_row + 1
sheet.write(numb_row,3,'Total', total_format)

# Column E - Design Load Total.
sheet.write(numb_row, 4,'=SUM(E2:E{})'.format(numb_row), sum_format)

# Column F - Design Load per m2.
sheet.write(numb_row, 5,'=AVERAGE(F2:F{})'.format(numb_row),sum_format)

# Close Excel workbook and use python to open it up in windows.
book.close()
os.startfile('Heating Loads.xlsx')