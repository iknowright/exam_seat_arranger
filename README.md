# Exam Seat Arranger

A python script for seat arrangement for examinations.

## Features
* Support custom layout
  * Arrange seat pattern as your wish
* Simple setup and configure procedure
* Generate a new file based on input layout

## Getting Started

### Prerequisites

* python
  * openpyxl - For interaction of Excel/Sheets

### Configuration

#### Installing Package
```lan=shell
pip install openpyxl
```

#### Configure *seats.xlsx*
```
# file structure
# seats.xlsx
-- sheets
    -- students
    -- stats
    -- {room Name}
    -- {room Name}
    -- {room Name}
    -- {room Name}
```

For *students*:
- Fill in the students name in Column A [A2:A{n}], the of the columns will be ignored.

For *stats*:

- Fill in the `Col("Exam Rooms")` with the sheets' names (room names) and specify which column to start in the romm sheet by adding entry in `Col("Start Column")` and which column to end in the room sheet by adding entry in `Col("End Column")`. Finally, set a row range for arranger to scan for each room by adding the number to `Col("Rows")`.

- About the rest of the column are automatically calculate by the function inside the corresponding cells. So you do not need to change anything there.

For a *{room}*:

- Before *Row 5 ( Default )*, you can put anything above the row. Normally, there are whiteboard and stage in front of the seats, you can draw that out id you want.

- After *Row 5 ( Default )*, you can start to arrange students' seats by putting text `"x"` to the cell. After that, it will be nice if you put some color for the seats cell to indicate table area.

#### Done configuration
After the above configurations, you can now run the seat arranger.

## Running the script

Execute the command below with generate an arranged file called *arranged_{input_file}* -> *arranged_seats.xlsx*

```lan=shell
python seat_arranger -i seats.xlsx -n {number_of_rooms}
```

* `-i` argument for input file
* `-n` argument for number of rooms to arrange seats


## Full example
*seats.xlsx* with sheets `['students', 'stats', '65105', 'A3201', '4264']`.

* `students`
![](https://i.imgur.com/f16W2Aw.png)

* `stats`
![](https://i.imgur.com/nGYmerm.png)

* before arrangement `65105`
![](https://i.imgur.com/Ipd4d7i.png)

Execution

`python seat_arranger -i seats.xlsx -n`

New file

*arranged_seats.xlsx*

* after arrangement `65105`
![](https://i.imgur.com/s587ATB.png)



## Future Works
* Pure seat arrangement based on student list and table layout
    * Generate seat layout without setting the `"x"` on the seats
* Add for seat arrangement patterns

## Inspiration and reference
[TOC-Seat-Arranger](https://github.com/Sirius207/TOC-Seat-Arranger) by Po-Chun

## Contribution

* ChaiShi
