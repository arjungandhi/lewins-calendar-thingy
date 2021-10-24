# Requirements
It's recommended it install and use pipenv for python version management
https://pypi.org/project/pipenv


if you don't want to do that use pip to install openpyxl and ics
`pip install openpyxl ics`

# Usage

## Initialize pipenv env (skip if libraries installed with pip)
1. Navigate to the root of the repo
2. `pipenv install` this creates a virtual enviroment and installs all the packages you need
3. `pipenv shell` initializes the virtual env and puts you in it

## Convert the excel file to ics
1. go into the src folder

`python main.py --file /path/to/excelfile.xlsx`

2. This will generate a file called events.ics

## Viewing the events
go to a calendar provider of choice.

### Some random guys code (kinda ugly)
Just upload your file here and navigate to the term you want to see the events should all be there
https://larrybolt.github.io/online-ics-feed-viewer/

### Google calendar (less ugly)

1. create a new calendar use the + button next to other calendars in the sidebar

2. name the calendar and make it

3. click the import & export button in the sidebar (in settings but you should see it in the sidebar)

4. upload the ics, select the newly made calendar as where to import it 



Navigate to the date range in the term you want to look at, all the events *should* be there. I included a the row number in the name so you can more easily figure out which row you need to modify.



