# python-openpyxl-cheatsheet

1. [Template](#template)
2. [Snippets](#snippets)

## Template

```python
## Full template

# Import traceback for error printing
import traceback

# Import openpyxl
from openpyxl import Workbook

# Declare the path to save the Workbook
path = 'hello_world.xlsx'

# Initialize the new Workbook
wb = Workbook()

# Connect to the active Worksheet
ws = wb.active

# Overall error handling to make sure to save and close the Workbook in the event of an error
try:
    # Write values to the Workbook using the two possible approaches
    ws['A1'] = 'Hello World'
    ws.cell(row=1, column=2).value = 'Hello Universe'

    # Declare variables for the area/range of the loop
    min_row = 2
    max_row = 5
    min_col = 1
    max_col = 2

    # Write values to the Workbook in a loop
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.value = 'Looping outer space'
            
# In the event of an error, print the error
except:
    # Print the error
    print(traceback.format_exc())
    
# No matter what, close and save the Workbook
finally:
    # Save the Workbook to the declared path
    wb.save(path)

    # Close the Workbook
    wb.close()
```

## Snippets

### Create Workbook
```python
# Import openpyxl
from openpyxl import Workbook

# Initialize the new Workbook
wb = Workbook()

# Connect to the active Worksheet
ws = wb.active
```

### Load Existing Workbook
```python
# Import openpyxl
from openpyxl import load_workbook

# Declare the path to the Workbook
path = 'hello_world.xlsx'

# Initialize the new Workbook
wb = load_workbook(filename=path)

# Connect to the active Worksheet
ws = wb.active

# Connect to another Worksheet using the name
ws_second = wb['Sheet2']
```

### Connect to Worksheets
```python
# Connect to the active Worksheet
ws = wb.active

# Connect to another Worksheet using the name
ws_second = wb['Sheet2']
```

### Create Worksheets
```python
# Create new Worksheet at the end (default)
ws = wb.create_sheet("NewSheet1")

# Create new Worksheet as the first
ws_second = wb.create_sheet("NewSheet2", 0)
```

### Rename Worksheets
```python
# Set the name (title) of the Worksheet
ws.title = "New Title"
```

### Write Values
```python
# Write values to the Workbook using the two possible approaches
ws['A1'] = 'Hello World'
ws.cell(row=1, column=2).value = 'Hello Universe'

# Declare variables for the area/range of the loop
min_row = 2
max_row = 5
min_col = 1
max_col = 2

# Write values to the Workbook in a loop
for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
    for cell in row:
        cell.value = 'Looping outer space'
```

### Append Values
```python
# Declare variables for the area/range of the loop
rows = 10

# Append values to the Workbook in a loop
for index in range(rows - 1):
    ws.append([[index, "Hello World"], [index, "test"]])
```

### Save and close Workbook
```python
# Declare the path to save the Workbook
path = 'hello_world.xlsx'

# Save the Workbook to the declared path
wb.save(path)

# Close the Workbook
wb.close()
```
